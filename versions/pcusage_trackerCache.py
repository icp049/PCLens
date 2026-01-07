import os
import sys
import json
import hashlib
from pathlib import Path
from datetime import datetime
from collections import defaultdict
import threading

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import tkinter as tk
from tkinter import ttk, filedialog

import customtkinter as ctk
from tkcalendar import DateEntry

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")


# =========================
# Cache helpers
# =========================
APP_NAME = "PCUsageVisualizer"
CACHE_VERSION = 1  # bump this if you change processing logic and want to invalidate old caches

def get_cache_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    p = Path(base) / APP_NAME / "cache"
    p.mkdir(parents=True, exist_ok=True)
    return p

def file_signature(path: str) -> dict:
    st = os.stat(path)
    return {
        "abs_path": os.path.abspath(path),
        "size": int(st.st_size),
        "mtime": int(st.st_mtime),
        "cache_version": CACHE_VERSION,
    }

def cache_key(sig: dict) -> str:
    # stable key derived from file path (so different excels cache separately)
    h = hashlib.sha256(sig["abs_path"].encode("utf-8")).hexdigest()[:16]
    return h

def cache_paths(sig: dict) -> tuple[Path, Path]:
    key = cache_key(sig)
    cache_dir = get_cache_dir()
    data_path = cache_dir / f"{key}.pkl"
    meta_path = cache_dir / f"{key}.json"
    return data_path, meta_path

def try_load_cache(excel_path: str):
    global df, site_timelines, monthly_site_pcs

    sig = file_signature(excel_path)
    data_path, meta_path = cache_paths(sig)

    if not data_path.exists() or not meta_path.exists():
        return None

    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
    except Exception:
        return None

    # must match signature exactly (including CACHE_VERSION)
    if meta.get("signature") != sig:
        return None

    try:
        payload = pd.read_pickle(data_path)
    except Exception:
        return None

    df = payload["df"]
    site_timelines = payload["site_timelines"]

    monthly_site_pcs.clear()
    # restore into defaultdict(lambda: defaultdict(set))
    restored = defaultdict(lambda: defaultdict(set))
    for site, month_map in payload["monthly_site_pcs"].items():
        for m_str, pcs_list in month_map.items():
            restored[site][pd.Period(m_str)] = set(pcs_list)
    monthly_site_pcs.update(restored)

    return meta

def save_cache(excel_path: str):
    global df, site_timelines, monthly_site_pcs

    sig = file_signature(excel_path)
    data_path, meta_path = cache_paths(sig)

    # Convert Period keys to strings for pickle portability
    monthly_serial = {}
    for site, month_map in monthly_site_pcs.items():
        monthly_serial[site] = {}
        for m, pcs_set in month_map.items():
            monthly_serial[site][str(m)] = sorted(list(pcs_set))

    payload = {
        "df": df,
        "site_timelines": site_timelines,
        "monthly_site_pcs": monthly_serial
    }

    pd.to_pickle(payload, data_path)

    meta = {
        "signature": sig,
        "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "excel_name": os.path.basename(sig["abs_path"]),
        "excel_path": sig["abs_path"],
    }
    meta_path.write_text(json.dumps(meta, indent=2), encoding="utf-8")
    return meta


# =========================
# Global state
# =========================
df = None
site_timelines = {}
scatter = None
filtered_data_for_hover = None
annotation = None
monthly_site_pcs = defaultdict(lambda: defaultdict(set))


# =========================
# Export
# =========================
def export_plot_data():
    if not site_timelines or df is None:
        status_label.configure(text="⚠ Cannot export — no data loaded.")
        return

    export_rows = []
    thresholds = list(range(10, 101, 10))
    all_months = sorted(df['Login Time'].dt.to_period('M').dropna().unique())

    export_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Save Threshold Plot Data"
    )

    if not export_path:
        status_label.configure(text="❌ Export canceled.")
        return

    export_win = ctk.CTkToplevel(root)
    export_win.title("Exporting...")
    export_win.geometry("300x140")
    export_win.transient(root)
    export_win.grab_set()

    cancel_flag = threading.Event()

    label = ctk.CTkLabel(export_win, text="Exporting... 0%")
    label.pack(pady=(10, 5))

    progress = ctk.CTkProgressBar(export_win, mode="determinate")
    progress.pack(fill="x", padx=20, pady=(0, 10))
    progress.set(0)

    def handle_cancel():
        cancel_flag.set()
        if export_win.winfo_exists():
            export_win.destroy()
        status_label.configure(text="❌ Export canceled.")

    cancel_btn = ctk.CTkButton(export_win, text="❌ Cancel", command=handle_cancel)
    cancel_btn.pack(pady=(0, 10))

    def do_export():
        try:
            total_minutes = 0
            for month in all_months:
                month_start = month.to_timestamp()
                month_end = (month_start + pd.offsets.MonthBegin(1))
                for site, (timeline, _) in site_timelines.items():
                    filtered_timeline = timeline[(timeline.index >= month_start) & (timeline.index < month_end)]
                    for pct in thresholds:
                        monthly_pcs = monthly_site_pcs.get(site, {}).get(month, set())
                        monthly_count = len(monthly_pcs)
                        if monthly_count == 0:
                            continue
                        required_count = int(np.ceil((pct / 100) * monthly_count)) or 1
                        total_minutes += len(filtered_timeline[filtered_timeline['ActivePCs'] >= required_count])

            if total_minutes == 0:
                root.after(0, lambda: status_label.configure(text="⚠ No data to export."))
                return

            completed_minutes = 0

            for month in all_months:
                if cancel_flag.is_set():
                    return
                month_start = month.to_timestamp()
                month_end = (month_start + pd.offsets.MonthBegin(1))
                month_str = str(month)

                for site, (timeline, _) in site_timelines.items():
                    if cancel_flag.is_set():
                        return

                    filtered_timeline = timeline[(timeline.index >= month_start) & (timeline.index < month_end)]

                    for pct in thresholds:
                        if cancel_flag.is_set():
                            return

                        monthly_pcs = monthly_site_pcs.get(site, {}).get(month, set())
                        monthly_count = len(monthly_pcs)
                        if monthly_count == 0:
                            continue

                        required_count = int(np.ceil((pct / 100) * monthly_count)) or 1
                        filtered_minutes = filtered_timeline[filtered_timeline['ActivePCs'] >= required_count].index

                        for minute in filtered_minutes:
                            export_rows.append({
                                'Branch': site,
                                'Threshold (%)': pct,
                                'Timestamp': minute,
                                'PCs Used': f"{required_count} of {monthly_count}",
                                'Month': month_str
                            })
                            completed_minutes += 1
                            if completed_minutes % 200 == 0 or completed_minutes == total_minutes:
                                progress_val = completed_minutes / total_minutes
                                root.after(0, lambda v=progress_val: progress.set(v))
                                root.after(0, lambda v=progress_val: label.configure(
                                    text=f"Exporting... {int(v * 100)}%"))

            if not cancel_flag.is_set():
                pd.DataFrame(export_rows).to_excel(export_path, index=False)
                root.after(0, lambda: status_label.configure(
                    text=f"📤 Exported all-month plot data to: {export_path}"
                ))
        except Exception as e:
            root.after(0, lambda: status_label.configure(text=f"❌ Export failed: {e}"))
        finally:
            if export_win.winfo_exists():
                root.after(0, export_win.destroy)

    threading.Thread(target=do_export, daemon=True).start()


# =========================
# Plot helpers
# =========================
def _show_no_data_title(site, start_date, end_date, pct_required, month_info=None, note="No data for selected range."):
    end_inclusive = (end_date - pd.Timedelta(days=1)).date()
    title = f"PC Usage at {site} from {start_date.date()} to {end_inclusive} (≥{pct_required}%)"

    ax.clear()
    ax.set_title(title, fontsize=14)
    ax.set_xlabel("Time of Day")
    ax.set_ylabel("Date")
    ax.grid(True, linestyle='--', alpha=0.3)
    ax.text(0.5, 0.5, f"⚠ {note}", transform=ax.transAxes, ha="center", va="center", fontsize=12)

    fig.tight_layout()
    plot_canvas.draw()

    status_detail_text.configure(state="normal")
    status_detail_text.delete("1.0", "end")
    if month_info:
        status_detail_text.insert("end", "📊 Monthly PC Counts & Thresholds\n")
        status_detail_text.insert("end", "\n".join(month_info))
    status_detail_text.configure(state="disabled")


def update_plot():
    global scatter, filtered_data_for_hover, annotation

    if df is None or not site_timelines:
        status_label.configure(text="⚠ Please load an Excel file first.")
        return

    site = site_var.get()
    if not site or site not in site_timelines:
        status_label.configure(text="⚠ Please select a branch that has data loaded.")
        return

    pct_required = int(percent_var.get())

    start_str = start_date_var.get()
    end_str = end_date_var.get()
    if not start_str or not end_str:
        status_label.configure(text="⚠ Please select both start and end dates.")
        return

    try:
        start_date = pd.to_datetime(start_str)
        end_date = pd.to_datetime(end_str) + pd.Timedelta(days=1)  # exclusive internally
    except Exception:
        status_label.configure(text="⚠ Invalid date format.")
        return

    month_range = pd.period_range(start_date.to_period('M'),
                                  (end_date - pd.Timedelta(days=1)).to_period('M'))

    timeline, _ = site_timelines[site]

    filtered_parts = []
    month_info = []
    change_notes = []
    change_details = []
    prev_set = None


    for m in month_range:
        pcs_now = set(monthly_site_pcs.get(site, {}).get(m, set()))
        if prev_set is not None:
            added = sorted(pcs_now - prev_set)
            removed = sorted(prev_set - pcs_now)
            if added:
                change_notes.append(f"NOTE: {len(added)} PC(s) added starting {m.strftime('%B %Y')}.")
                change_details.append(f"  + {m.strftime('%b %Y')}: " + (", ".join(added[:20]) + (f', +{len(added)-20} more' if len(added) > 20 else '')))
            if removed:
                change_notes.append(f"NOTE: {len(removed)} PC(s) removed starting {m.strftime('%B %Y')}.")
                change_details.append(f"  - {m.strftime('%b %Y')}: " + (", ".join(removed[:20]) + (f', +{len(removed)-20} more' if len(removed) > 20 else '')))
        prev_set = pcs_now

   
    for m in month_range:
        pcs_in_month = monthly_site_pcs.get(site, {}).get(m, set())
        monthly_pc_count = len(pcs_in_month)
        if monthly_pc_count == 0:
            continue

        monthly_required = int(np.ceil((pct_required / 100) * monthly_pc_count)) or 1
        month_info.append(f"- {m.strftime('%b %Y')}: {monthly_pc_count} PCs → Required: {monthly_required}")

        month_start = m.to_timestamp()
        month_end = (month_start + pd.offsets.MonthBegin(1))
        month_timeline = timeline[(timeline.index >= month_start) & (timeline.index < month_end)]
        month_filtered = month_timeline[month_timeline['ActivePCs'] >= monthly_required].copy()
        if not month_filtered.empty:
            filtered_parts.append(month_filtered)

    if not filtered_parts:
        status_label.configure(text="⚠ No data to show for selected range.")
        _show_no_data_title(site, start_date, end_date, pct_required, month_info)
        return

    filtered = pd.concat(filtered_parts)

    # ---------- Plot ----------
    ax.clear()

    annotation = ax.annotate(
        text='',
        xy=(0, 0),
        xytext=(15, 15),
        textcoords='offset points',
        bbox=dict(boxstyle='round', fc='w'),
        arrowprops=dict(arrowstyle='->'),
        zorder=11
    )
    annotation.set_visible(False)

    all_dates = pd.date_range(start=start_date, end=end_date - pd.Timedelta(days=1)).date

    filtered = filtered.reset_index()
    filtered['HourFloat'] = filtered['Timestamp'].dt.hour + filtered['Timestamp'].dt.minute / 60.0
    filtered['YFloat'] = filtered['Timestamp'].map(lambda x: x.toordinal())
    filtered['FullDateTime'] = filtered['Timestamp']

    scatter = ax.scatter(filtered['HourFloat'], filtered['YFloat'], s=80, picker=20, zorder=1)
    filtered_data_for_hover = filtered

    xticks = np.arange(0, 24.01, 1)
    ax.set_xticks(xticks)
    ax.set_xticklabels([f"{int(h):02}:00" for h in xticks], rotation=45)

    y_ticks = [d.toordinal() for d in all_dates]
    y_labels = [d.strftime('%b-%d') for d in all_dates]

    total_days = len(all_dates)
    step = 1
    if total_days > 40:
        step = 7
    elif total_days > 25:
        step = 5
    elif total_days > 15:
        step = 3
    elif total_days > 7:
        step = 2

    reduced_ticks = y_ticks[::step]
    reduced_labels = y_labels[::step]
    ax.set_yticks(reduced_ticks)
    ax.set_yticklabels(reduced_labels, fontsize=8)
    ax.set_ylim([y_ticks[-1], y_ticks[0]])

    end_inclusive = (end_date - pd.Timedelta(days=1)).date()
    ax.set_title(f"PC Usage at {site} from {start_date.date()} to {end_inclusive} (≥{pct_required}%)", fontsize=14)
    ax.set_xlabel("Time of Day")
    ax.set_ylabel("Date")
    ax.grid(True, linestyle='--', alpha=0.3)

    fig.tight_layout()
    plot_canvas.draw()

    # ---------- Build right-side details ----------
    qualified_minutes = set(filtered['Timestamp'])

    all_logins = df[df['Location'] == site]
    unique_pcs = sorted(all_logins['Resource'].unique())

    qualified_times = sorted(qualified_minutes)
    ranges = []
    if qualified_times:
        start = prev = qualified_times[0]
        for t in qualified_times[1:]:
            if (t - prev).seconds == 60:
                prev = t
            else:
                ranges.append((start, prev))
                start = prev = t
        ranges.append((start, prev))

    table_lines = []
    table_lines.append(f"{'Month':<12} {'PCs':>5} {f'Required PCs at {pct_required}%':>20}")
    for m in month_range:
        pcs_count = len(monthly_site_pcs.get(site, {}).get(m, set()))
        if pcs_count == 0:
            continue
        req_count = int(np.ceil((pct_required / 100) * pcs_count)) or 1
        table_lines.append(f"{m.strftime('%b %Y'):<12} {pcs_count:>5} {req_count:>20}")

    sections = []
    sections.append("📊 Monthly PC Counts & Thresholds")
    sections.append("```")
    sections.extend(table_lines)
    sections.append("```")

    if change_notes:
        sections.append("\n🔄 Changes in PC Counts")
        for note in change_notes:
            if "added" in note:
                sections.append(f"  ✅ {note}")
            elif "removed" in note:
                sections.append(f"  ❌ {note}")
            else:
                sections.append(f"  ⚠ {note}")

    if change_details:
        sections.append("  ↳ PCs changed (by name):")
        sections.extend(change_details)

    sections.append(f"\n🕒 Qualified Time Ranges (≥{pct_required}% PCs active)")
    if ranges:
        for s, e in ranges[:50]:
            if s.date() == e.date():
                sections.append(f"  • {s.strftime('%Y-%m-%d %H:%M')} → {e.strftime('%H:%M')}")
            else:
                sections.append(f"  • {s.strftime('%Y-%m-%d %H:%M')} → {e.strftime('%Y-%m-%d %H:%M')}")
        if len(ranges) > 50:
            sections.append(f"  ...and {len(ranges) - 50} more.")
    else:
        sections.append("  ⚠ No qualified time ranges found.")

    status_detail_text.configure(state="normal")
    status_detail_text.delete("1.0", "end")
    status_detail_text.insert("end", "\n".join(sections))
    status_detail_text.configure(state="disabled")

    status_label.configure(text="✔ Plot updated.")


def on_hover(event):
    global scatter, filtered_data_for_hover, annotation

    if scatter is None or filtered_data_for_hover is None or df is None:
        return
    if event.inaxes != ax:
        return

    contains, info = scatter.contains(event)
    if contains and info and "ind" in info and len(info["ind"]) > 0:
        ind_list = info["ind"]

        mouse_x, mouse_y = event.xdata, event.ydata
        nearest_idx = min(
            ind_list,
            key=lambda i: (filtered_data_for_hover.iloc[i]['HourFloat'] - mouse_x) ** 2 +
                          (filtered_data_for_hover.iloc[i]['YFloat'] - mouse_y) ** 2
        )

        row = filtered_data_for_hover.iloc[nearest_idx]
        x = row['HourFloat']
        y = row['YFloat']
        timestamp = row['FullDateTime']

        site = site_var.get()
        all_logins_site = df[df['Location'] == site]
        mask = (all_logins_site['Login Time'] <= timestamp) & (all_logins_site['Logout Time'] >= timestamp)
        active_pcs = sorted(all_logins_site.loc[mask, 'Resource'].unique())

        MAX_NAMES = 12
        if len(active_pcs) > MAX_NAMES:
            shown = ", ".join(active_pcs[:MAX_NAMES])
            extra = len(active_pcs) - MAX_NAMES
            pcs_text = f"{shown}, +{extra} more"
        else:
            pcs_text = ", ".join(active_pcs) if active_pcs else "None"

        hover_text = f"{timestamp.strftime('%b %d @ %H:%M')}\nActive PCs: {len(active_pcs)}\n{pcs_text}"

        annotation.xy = (x, y)
        annotation.set_text(hover_text)
        annotation.get_bbox_patch().set_facecolor("lightyellow")
        annotation.get_bbox_patch().set_edgecolor("gray")
        annotation.set_alpha(0.9)
        annotation.set_zorder(11)
        annotation.set_visible(True)
        fig.canvas.draw_idle()
    else:
        if annotation is not None:
            annotation.set_visible(False)
            fig.canvas.draw_idle()


# =========================
# Loading + caching
# =========================
def reset_ui():
    global df, site_timelines
    df = None
    site_timelines.clear()
    monthly_site_pcs.clear()

    status_label.configure(text="📂 Please load an Excel file to begin.")
    cache_label.configure(text="Cache: (none)")
    site_dropdown.configure(values=[])
    start_date_var.set("")
    end_date_var.set("")
    site_var.set("")

    status_detail_text.configure(state="normal")
    status_detail_text.delete("1.0", "end")
    status_detail_text.configure(state="disabled")

    ax.clear()
    ax.axis('off')
    plot_canvas.draw()


def load_and_initialize():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        return

    cancelled = threading.Event()

    loading_win = ctk.CTkToplevel(root)
    loading_win.title("Loading...")
    loading_win.geometry("320x160")
    loading_win.transient(root)
    loading_win.grab_set()

    def handle_cancel():
        cancelled.set()
        if loading_win.winfo_exists():
            loading_win.destroy()
        reset_ui()

    loading_win.protocol("WM_DELETE_WINDOW", handle_cancel)

    label = ctk.CTkLabel(loading_win, text="Loading... 0%")
    label.pack(pady=(10, 5))

    progress = ctk.CTkProgressBar(loading_win, mode="determinate")
    progress.pack(fill="x", padx=20, pady=(0, 10))
    progress.set(0)

    cancel_btn = ctk.CTkButton(loading_win, text="❌ Cancel", command=handle_cancel)
    cancel_btn.pack(pady=(0, 10))

    def do_load():
        try:
            reset_ui()

            # ✅ Try cache first
            meta = try_load_cache(file_path)
            if meta:
                root.after(0, lambda: populate_ui_after_load(meta, from_cache=True))
                return

            # ❌ No cache -> build
            load_data(file_path, progress, label, cancelled)
            if cancelled.is_set():
                return

            meta2 = save_cache(file_path)
            root.after(0, lambda: populate_ui_after_load(meta2, from_cache=False))

        except Exception as e:
            root.after(0, lambda: status_label.configure(text=f"❌ Failed to load: {e}"))
            root.after(0, reset_ui)
        finally:
            if loading_win.winfo_exists():
                root.after(0, loading_win.destroy)

    threading.Thread(target=do_load, daemon=True).start()


def refresh_rebuild():
    file_path = filedialog.askopenfilename(
        title="Select Excel File (Force Rebuild)",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        return

    cancelled = threading.Event()

    loading_win = ctk.CTkToplevel(root)
    loading_win.title("Rebuilding Cache...")
    loading_win.geometry("320x160")
    loading_win.transient(root)
    loading_win.grab_set()

    def handle_cancel():
        cancelled.set()
        if loading_win.winfo_exists():
            loading_win.destroy()
        reset_ui()

    loading_win.protocol("WM_DELETE_WINDOW", handle_cancel)

    label = ctk.CTkLabel(loading_win, text="Rebuilding... 0%")
    label.pack(pady=(10, 5))

    progress = ctk.CTkProgressBar(loading_win, mode="determinate")
    progress.pack(fill="x", padx=20, pady=(0, 10))
    progress.set(0)

    cancel_btn = ctk.CTkButton(loading_win, text="❌ Cancel", command=handle_cancel)
    cancel_btn.pack(pady=(0, 10))

    def do_refresh():
        try:
            reset_ui()
            load_data(file_path, progress, label, cancelled)
            if cancelled.is_set():
                return
            meta = save_cache(file_path)
            root.after(0, lambda: populate_ui_after_load(meta, from_cache=False, forced=True))
        except Exception as e:
            root.after(0, lambda: status_label.configure(text=f"❌ Refresh failed: {e}"))
        finally:
            if loading_win.winfo_exists():
                root.after(0, loading_win.destroy)

    threading.Thread(target=do_refresh, daemon=True).start()


def populate_ui_after_load(meta, from_cache: bool, forced: bool = False):
    # dropdowns
    site_dropdown.configure(values=sorted(df['Location'].unique()))
    if df["Location"].nunique() > 0:
        site_var.set(sorted(df["Location"].unique())[0])

    min_date = df["Login Time"].min().date()
    max_date = df["Login Time"].max().date()
    start_date_picker.config(mindate=min_date, maxdate=max_date)
    end_date_picker.config(mindate=min_date, maxdate=max_date)
    start_date_picker.set_date(min_date)
    end_date_picker.set_date(max_date)

    cache_label.configure(text=f"Cache: ✅ {meta['last_updated']} ({meta['excel_name']})")

    if forced:
        status_label.configure(text="🔄 Forced rebuild complete. Ready to visualize.")
    else:
        status_label.configure(text=("⚡ Loaded from cache. Ready to visualize." if from_cache else
                                     "✔ File loaded + cached. Ready to visualize."))


def load_data(file_path, progress_widget, label_widget, cancelled):
    global df, site_timelines, monthly_site_pcs

    REQUIRED = {"login time", "logout time", "resource", "location"}

    preview = pd.read_excel(file_path, header=None, nrows=50)
    header_row_idx = None
    for i in range(len(preview)):
        row_vals = preview.iloc[i].astype(str).str.strip().str.lower()
        cells = set(v for v in row_vals if v and v != "nan")
        if REQUIRED.issubset(cells):
            header_row_idx = i
            break

    if header_row_idx is None:
        root.after(0, lambda: status_label.configure(
            text="❌ Could not find header row containing: Login Time, Logout Time, Resource, Location"
        ))
        return

    raw = pd.read_excel(file_path, header=header_row_idx)
    norm_to_orig = {str(c).strip().lower(): c for c in raw.columns}

    missing = [col for col in REQUIRED if col not in norm_to_orig]
    if missing:
        root.after(0, lambda: status_label.configure(
            text=f"❌ Missing required column(s): {', '.join([m.title() for m in missing])}"
        ))
        return

    df = raw[[norm_to_orig["login time"],
              norm_to_orig["logout time"],
              norm_to_orig["resource"],
              norm_to_orig["location"]]].copy()

    df.rename(columns={
        norm_to_orig["login time"]: "Login Time",
        norm_to_orig["logout time"]: "Logout Time",
        norm_to_orig["resource"]: "Resource",
        norm_to_orig["location"]: "Location",
    }, inplace=True)

    dt_format = '%m/%d/%Y %I:%M %p'
    df["Login Time"] = pd.to_datetime(df["Login Time"], format=dt_format, errors="coerce")
    df["Logout Time"] = pd.to_datetime(df["Logout Time"], format=dt_format, errors="coerce")
    df["Resource"] = df["Resource"].astype(str).str.strip()
    df["Location"] = df["Location"].astype(str).str.strip()

    before = len(df)
    df.dropna(subset=["Login Time", "Logout Time", "Resource", "Location"], inplace=True)
    dropped = before - len(df)
    if dropped > 0:
        root.after(0, lambda: status_label.configure(text=f"⚠ Skipped {dropped} rows with missing required values."))

    bad_locations = ["express", "outreach", "laptop"]
    df = df[~df["Location"].str.lower().str.contains("|".join(bad_locations), na=False)]

    df = df[df["Logout Time"] >= df["Login Time"]]
    if df.empty:
        root.after(0, lambda: status_label.configure(text="⚠ No valid rows after cleaning."))
        return

    min_date = df["Login Time"].min().date()
    max_date = df["Login Time"].max().date()

    root.after(0, lambda: start_date_picker.config(mindate=min_date, maxdate=max_date))
    root.after(0, lambda: end_date_picker.config(mindate=min_date, maxdate=max_date))
    root.after(0, lambda: start_date_picker.set_date(min_date))
    root.after(0, lambda: end_date_picker.set_date(max_date))

    start = df["Login Time"].min().replace(day=1)
    end = (df["Login Time"].max() + pd.Timedelta(days=1)).replace(day=1) + pd.offsets.MonthEnd(0)
    full_range = pd.date_range(start=start, end=end, freq="min")
    timeline_template = pd.DataFrame({"Timestamp": full_range}).set_index("Timestamp")

    site_timelines.clear()
    monthly_site_pcs.clear()

    total_rows = len(df)
    processed_rows = 0

    for site in df["Location"].unique():
        if cancelled.is_set():
            return

        site_df = df[df["Location"] == site]
        pcs = sorted(site_df["Resource"].unique())
        num_pcs = len(pcs)

        timeline = timeline_template.copy()
        timeline["ActivePCs"] = 0

        for _, row in site_df.iterrows():
            if cancelled.is_set():
                return

            login, logout = row["Login Time"], row["Logout Time"]

            if pd.notna(login):
                month = login.to_period("M")
                monthly_site_pcs[site][month].add(row["Resource"])

            if pd.notna(login) and pd.notna(logout):
                active_minutes = pd.date_range(start=login, end=logout, freq="min")
                active_minutes = active_minutes.intersection(timeline.index)
                timeline.loc[active_minutes, "ActivePCs"] += 1

            processed_rows += 1
            if processed_rows % 50 == 0 or processed_rows == total_rows:
                val = processed_rows / total_rows
                root.after(0, lambda v=val: progress_widget.set(v))
                root.after(0, lambda v=val: label_widget.configure(text=f"Processing data... {int(v * 100)}%"))

        site_timelines[site] = (timeline, num_pcs)


# =========================
# UI
# =========================
root = ctk.CTk()
root.title("PC Usage Visualizer")
root.geometry("1000x1000")
root.minsize(1000, 1000)

scrollable_frame = ctk.CTkScrollableFrame(root)
scrollable_frame.pack(fill="both", expand=True, padx=10, pady=10)

plot_frame = ctk.CTkFrame(scrollable_frame)
plot_frame.pack(fill="both", expand=True, padx=10, pady=(10, 0))
plot_frame.pack_propagate(True)

fig, ax = plt.subplots(figsize=(12, 6))
plot_canvas = FigureCanvasTkAgg(fig, master=plot_frame)
fig.canvas.mpl_connect('motion_notify_event', on_hover)
canvas_widget = plot_canvas.get_tk_widget()
canvas_widget.pack(fill="both", expand=True)

toolbar_container = tk.Frame(scrollable_frame)
toolbar_container.pack(fill=tk.X, padx=10, pady=(0, 10))

toolbar = NavigationToolbar2Tk(plot_canvas, toolbar_container)
toolbar.update()
toolbar.pack(side=tk.BOTTOM, fill=tk.X)

ax.clear()
ax.axis('off')
plot_canvas.draw()

status_label = ctk.CTkLabel(scrollable_frame, text="📂 Please load an Excel file to begin.", text_color="gray")
status_label.pack(fill=ctk.X, padx=10, pady=(5, 5))

cache_label = ctk.CTkLabel(scrollable_frame, text="Cache: (none)", text_color="gray")
cache_label.pack(fill=ctk.X, padx=10, pady=(0, 10))

btn_row = ctk.CTkFrame(scrollable_frame)
btn_row.pack(fill="x", padx=10, pady=(0, 10))

load_btn = ctk.CTkButton(btn_row, text="📥 Import File (Use Cache)", command=load_and_initialize)
load_btn.pack(side="left", padx=(0, 10))

refresh_btn = ctk.CTkButton(btn_row, text="🔄 Force Rebuild (Ignore Cache)", command=refresh_rebuild)
refresh_btn.pack(side="left")

detail_container = ctk.CTkFrame(scrollable_frame, height=150)
detail_container.pack(fill=ctk.X, padx=10, pady=(0, 10))
detail_container.pack_propagate(False)

status_detail_text = ctk.CTkTextbox(detail_container, wrap="none", font=("Courier New", 14))
status_detail_text.pack(fill="both", expand=True, padx=10, pady=10)
status_detail_text.configure(state="disabled")

frame = ctk.CTkFrame(scrollable_frame)
frame.pack(side=ctk.TOP, fill=ctk.X, padx=10, pady=10)
frame.grid_columnconfigure(0, weight=1)
frame.grid_columnconfigure(1, weight=1)

left_frame = ctk.CTkFrame(frame)
left_frame.grid(row=0, column=0, sticky='w')

ctk.CTkLabel(left_frame, text="Start Date:").pack(side=ctk.LEFT, padx=(0, 5))
start_date_var = tk.StringVar()
start_date_picker = DateEntry(left_frame, textvariable=start_date_var, width=12, background='darkblue', foreground='white', borderwidth=2)
start_date_picker.pack(side=ctk.LEFT)

ctk.CTkLabel(left_frame, text="  End Date:").pack(side=ctk.LEFT, padx=(10, 5))
end_date_var = tk.StringVar()
end_date_picker = DateEntry(left_frame, textvariable=end_date_var, width=12, background='darkblue', foreground='white', borderwidth=2)
end_date_picker.pack(side=ctk.LEFT)

ctk.CTkLabel(left_frame, text="   Select Branch:").pack(side=ctk.LEFT, padx=(10, 5))
site_var = ctk.StringVar()
site_dropdown = ttk.Combobox(left_frame, textvariable=site_var, state='readonly')
site_dropdown.pack(side=ctk.LEFT)

apply_btn = ctk.CTkButton(
    frame,
    text="✅ Apply",
    command=update_plot,
    fg_color="green",
    hover_color="#006400"
)
apply_btn.grid(row=0, column=2, padx=(20, 0), sticky='e')

export_btn = ctk.CTkButton(frame, text="📤 Export for PowerBI", command=export_plot_data)
export_btn.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(10, 0))

percent_var = tk.IntVar(value=100)
slider_frame = ctk.CTkFrame(frame)
slider_frame.grid(row=0, column=1, padx=(20, 0), sticky='e')

slider_label = ctk.CTkLabel(slider_frame, text="Select Threshold:", text_color="white", font=ctk.CTkFont(size=12))
slider_label.pack(pady=(0, 5))

percent_slider = ctk.CTkSlider(
    slider_frame,
    from_=10,
    to=100,
    number_of_steps=9,
    variable=percent_var,
    command=lambda val: percent_var.set(int(float(val)))
)
percent_slider.set(100)
percent_slider.pack(fill='x')

tick_frame = ctk.CTkFrame(slider_frame)
tick_frame.pack(fill='x', pady=(2, 0))
tick_frame.grid_columnconfigure(tuple(range(10)), weight=1)

for idx, i in enumerate(range(10, 101, 10)):
    lbl = ctk.CTkLabel(tick_frame, text=str(i), text_color="white", font=ctk.CTkFont(size=10))
    lbl.grid(row=0, column=idx, sticky='n')

root.mainloop()
