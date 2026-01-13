import os
import json
import hashlib
import shutil
import time
import pickle


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk
from tkinter import ttk, filedialog
from datetime import datetime
from collections import defaultdict
import threading
import customtkinter as ctk

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")
from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk

# Add this at the top
from tkcalendar import DateEntry


#######cache feature helper ###########


APP_NAME = "PCUsageVisualizer"
CACHE_VERSION = "v5"


def _default_month_map():
    return defaultdict(set)


monthly_site_pcs = defaultdict(_default_month_map)


def get_app_dir():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    d = os.path.join(base, APP_NAME)
    os.makedirs(d, exist_ok=True)
    return d


def get_data_dir():
    d = os.path.join(get_app_dir(), "data")
    os.makedirs(d, exist_ok=True)
    return d


def get_cache_dir():
    d = os.path.join(get_app_dir(), "cache")
    os.makedirs(d, exist_ok=True)
    return d


def get_config_path():
    return os.path.join(get_app_dir(), "config.json")


def load_config():
    p = get_config_path()
    if not os.path.exists(p):
        return {}
    try:
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg: dict):
    with open(get_config_path(), "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)


def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()[:16]


def cache_paths_for_hash(file_hash: str):
    base = os.path.join(get_cache_dir(), f"{CACHE_VERSION}_{file_hash}")
    return {
        "df": base + "_df.pkl",
        "monthly": base + "_monthly.pkl",
        "sites": base + "_sites.pkl",
        "timelines": base + "_timelines.npz",
    }


def cache_exists(file_hash: str) -> bool:
    p = cache_paths_for_hash(file_hash)
    missing = [k for k, path in p.items() if not os.path.exists(path)]
    if missing:
        print("CACHE MISS hash=", file_hash, "missing=", missing)
        # ✅ delete partial cache to prevent getting stuck forever
        for path in p.values():
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass
        return False
    print("CACHE HIT hash=", file_hash)
    return True


def copy_to_current(original_path: str) -> str:
    ext = os.path.splitext(original_path)[1].lower()
    if ext not in [".xlsx", ".xls"]:  # match your current loader
        raise ValueError("Please select an Excel file (.xlsx or .xls).")

    dest = os.path.join(get_data_dir(), "current.xlsx")  # normalize
    shutil.copy2(original_path, dest)
    return dest


def safe_read_retry(func, retries=10, delay=0.25):
    """
    Excel can hold a lock briefly while saving.
    Retry a few times to avoid 'Permission denied' / sharing violations.
    """
    last_exc = None
    for _ in range(retries):
        try:
            return func()
        except Exception as e:
            last_exc = e
            time.sleep(delay)
    raise last_exc


#######cache feature helper ###########

# Global state
df = None
site_timelines = {}
scatter = None
filtered_data_for_hover = None
annotation = None


details_popout_win = None
details_popout_text = None

export_pending_after_load = False


def open_details_popout():
    global details_popout_win, details_popout_text

    # If already open, just bring it to front
    if details_popout_win is not None and details_popout_win.winfo_exists():
        details_popout_win.lift()
        details_popout_win.focus_force()
        return

    details_popout_win = ctk.CTkToplevel(root)
    site_name = site_var.get() if "site_var" in globals() else ""
    details_popout_win.title(f"Details — {site_name}" if site_name else "Details")
    details_popout_win.geometry("900x600")

    # Try to maximize on Windows
    try:
        details_popout_win.state("zoomed")
    except Exception:
        pass

    # --- Mirror the same "Details" card look ---
    detail_container = ctk.CTkFrame(details_popout_win, corner_radius=16)
    detail_container.pack(fill="both", expand=True, padx=10, pady=10)

    # Header bar
    detail_header = ctk.CTkFrame(detail_container, fg_color="transparent")
    detail_header.pack(fill="x", padx=14, pady=(12, 6))

    details_title = ctk.CTkLabel(
        detail_header,
        text="Details",
        font=ctk.CTkFont(size=16, weight="bold"),
    )
    details_title.pack(side="left")

    details_subtitle = ctk.CTkLabel(
        detail_header,
        text="(Pop Out)",
        text_color="#9aa0a6",
        font=ctk.CTkFont(size=12),
    )
    details_subtitle.pack(side="left", padx=(12, 0))

    def refresh_popout():
        if details_popout_text is None or not details_popout_text.winfo_exists():
            return
        content = status_detail_text.get("1.0", "end-1c")
        details_popout_text.configure(state="normal")
        details_popout_text.delete("1.0", "end")
        details_popout_text.insert("end", content)
        details_popout_text.configure(state="disabled")

    def copy_popout():
        try:
            txt = status_detail_text.get("1.0", "end-1c")
            details_popout_win.clipboard_clear()
            details_popout_win.clipboard_append(txt)
        except Exception:
            pass

    ctk.CTkButton(detail_header, text="📋 Copy", width=90, command=copy_popout).pack(
        side="right", padx=(8, 0)
    )
    ctk.CTkButton(
        detail_header, text="🔄 Refresh", width=110, command=refresh_popout
    ).pack(side="right")

    # Separator line
    sep = ctk.CTkFrame(detail_container, height=1, fg_color="#2a2a2a")
    sep.pack(fill="x", padx=14, pady=(0, 10))

    # Text area (match your Details textbox)
    details_popout_text = ctk.CTkTextbox(
        detail_container,
        wrap="word",
        font=("Segoe UI", 12),
        corner_radius=12,
    )
    details_popout_text.pack(fill="both", expand=True, padx=14, pady=(0, 14))
    details_popout_text.configure(state="disabled")

    # Fill it once immediately
    refresh_popout()

    def on_close():
        global details_popout_win, details_popout_text
        try:
            details_popout_win.destroy()
        except Exception:
            pass
        details_popout_win = None
        details_popout_text = None

    details_popout_win.protocol("WM_DELETE_WINDOW", on_close)


def export_plot_data():
    if not site_timelines:
        status_label.config(text="⚠ Cannot export — no data loaded.")
        return

    export_rows = []
    thresholds = list(range(10, 101, 10))
    all_months = sorted(df["Login Time"].dt.to_period("M").dropna().unique())

    # ✅ Prompt file path only once here
    # export_path = filedialog.asksaveasfilename(
    #     defaultextension=".xlsx",
    #     filetypes=[("Excel files", "*.xlsx")],
    #     title="Save Threshold Plot Data",
    # )
    export_path = os.path.join(get_app_dir(), "exports")
    os.makedirs(export_path, exist_ok=True)
    export_path = os.path.join(
        export_path, f"powerbi_export_{load_config().get('last_hash','unknown')}.xlsx"
    )

    if not export_path:
        root.after(0, lambda: status_label.configure(text="❌ Export canceled."))
        return

    # Set up export progress window
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
        export_win.destroy()
        root.after(0, lambda: status_label.configure(text="❌ Export canceled."))

    cancel_btn = ctk.CTkButton(export_win, text="❌ Cancel", command=handle_cancel)
    cancel_btn.pack(pady=(0, 10))

    def do_export():
        try:
            # First pass to count total minutes to export
            total_minutes = 0
            for month in all_months:
                month_start = month.to_timestamp()
                month_end = month_start + pd.offsets.MonthBegin(1)
                for site, (timeline, num_pcs) in site_timelines.items():
                    filtered_timeline = timeline[
                        (timeline.index >= month_start) & (timeline.index < month_end)
                    ]
                    for pct in thresholds:
                        monthly_pcs = monthly_site_pcs.get(site, {}).get(month, set())
                        monthly_count = len(monthly_pcs)
                        required_count = int(np.ceil((pct / 100) * monthly_count)) or 1
                        total_minutes += len(
                            filtered_timeline[
                                filtered_timeline["ActivePCs"] >= required_count
                            ]
                        )

            if total_minutes == 0:
                root.after(
                    0, lambda: status_label.configure(text="⚠ No data to export.")
                )
                return

            completed_minutes = 0

            for month in all_months:
                if cancel_flag.is_set():
                    return
                month_start = month.to_timestamp()
                month_end = month_start + pd.offsets.MonthBegin(1)
                month_str = str(month)

                for site, (timeline, num_pcs) in site_timelines.items():
                    if cancel_flag.is_set():
                        return

                    filtered_timeline = timeline[
                        (timeline.index >= month_start) & (timeline.index < month_end)
                    ]
                    all_logins = df[
                        (df["Location"] == site)
                        & (df["Login Time"] >= month_start)
                        & (df["Login Time"] < month_end)
                    ]
                    unique_pcs = sorted(all_logins["Resource"].unique())

                    for pct in thresholds:
                        if cancel_flag.is_set():
                            return

                        monthly_pcs = monthly_site_pcs.get(site, {}).get(month, set())
                        monthly_count = len(monthly_pcs)
                        if monthly_count == 0:
                            continue  # skip if no PCs in that month

                        required_count = int(np.ceil((pct / 100) * monthly_count)) or 1
                        filtered_minutes = filtered_timeline[
                            filtered_timeline["ActivePCs"] >= required_count
                        ].index

                        for minute in filtered_minutes:
                            export_rows.append(
                                {
                                    "Branch": site,
                                    "Threshold (%)": pct,
                                    "Timestamp": minute,
                                    "PCs Used": f"{required_count} of {monthly_count}",
                                    "Month": month_str,
                                }
                            )
                            completed_minutes += 1
                            if (
                                completed_minutes % 100 == 0
                                or completed_minutes == total_minutes
                            ):
                                progress_val = completed_minutes / total_minutes
                                root.after(0, lambda v=progress_val: progress.set(v))
                                root.after(
                                    0,
                                    lambda v=progress_val: label.configure(
                                        text=f"Exporting... {int(v * 100)}%"
                                    ),
                                )

            if not cancel_flag.is_set():
                pd.DataFrame(export_rows).to_excel(export_path, index=False)
                root.after(
                    0,
                    lambda: status_label.configure(
                        text=f"📤 Exported all-month plot data to: {export_path}"
                    ),
                )

        except Exception as e:
            root.after(0, lambda: status_label.configure(text=f"❌ Export failed: {e}"))
        finally:
            if export_win.winfo_exists():
                root.after(0, export_win.destroy)

    threading.Thread(target=do_export, daemon=True).start()


def _show_no_data_title(
    site,
    start_date,
    end_date,
    pct_required,
    month_info=None,
    note="No data for selected range.",
):

    end_inclusive = (end_date - pd.Timedelta(days=1)).date()
    title = f"PC Usage at {site} from {start_date.date()}  to {end_inclusive} (≥{pct_required}%)"

    ax.clear()
    ax.set_title(title, fontsize=14)
    ax.set_xlabel("Time of Day")
    ax.set_ylabel("Date")
    ax.grid(True, linestyle="--", alpha=0.3)

    ax.text(
        0.5,
        0.5,
        f"⚠ {note}",
        transform=ax.transAxes,
        ha="center",
        va="center",
        fontsize=12,
    )

    fig.tight_layout()
    plot_canvas.draw()

    status_detail_text.configure(state="normal")
    status_detail_text.delete("1.0", "end")
    if month_info:
        status_detail_text.insert("end", "📊 Monthly PC Counts & Thresholds\n")
        status_detail_text.insert("end", "\n".join(month_info))
    status_detail_text.configure(state="disabled")


def keep_only_continuous_runs(index_like, min_len=5):
    """
    Return only timestamps that belong to runs of >= min_len consecutive minutes.
    """
    idx = pd.DatetimeIndex(index_like).sort_values().unique()
    if len(idx) == 0:
        return idx

    s = pd.Series(idx)
    grp = s.diff().ne(pd.Timedelta(minutes=1)).cumsum()

    kept = []
    for _, g in s.groupby(grp):
        if len(g) >= min_len:
            kept.append(pd.DatetimeIndex(g.values))

    if not kept:
        return pd.DatetimeIndex([])
    return kept[0].append(kept[1:]) if len(kept) > 1 else kept[0]


def update_plot():
    global scatter, filtered_data_for_hover, annotation

    site = site_var.get()
    if not site or site not in site_timelines:
        root.after(
            0,
            lambda: status_label.configure(
                text="⚠ Please select a branch that has data loaded."
            ),
        )
        return

    pct_required = percent_var.get()

    start_str = start_date_var.get()
    end_str = end_date_var.get()
    if not start_str or not end_str:
        root.after(
            0,
            lambda: status_label.configure(
                text="⚠ Please select both start and end dates."
            ),
        )
        return

    try:
        start_date = pd.to_datetime(start_str)
        # we treat end_date as exclusive internally, but show inclusive in title
        end_date = pd.to_datetime(end_str) + pd.Timedelta(days=1)
    except Exception:
        root.after(0, lambda: status_label.configure(text="⚠ Invalid date format."))
        return

    # Build month range over the selected dates
    month_range = pd.period_range(
        start_date.to_period("M"), (end_date - pd.Timedelta(days=1)).to_period("M")
    )

    timeline, _ = site_timelines[site]

    filtered_parts = []
    month_info = []  # For status panel
    change_notes = []
    change_details = []
    prev_set = None

    # Detect changes in monthly PC counts (for notes)
    for m in month_range:
        pcs_now = set(monthly_site_pcs.get(site, {}).get(m, set()))
        if prev_set is not None:
            added = sorted(pcs_now - prev_set)
            removed = sorted(prev_set - pcs_now)
            if added:
                change_notes.append(
                    f"NOTE: {len(added)} PC(s) added starting {m.strftime('%B %Y')}."
                )
                change_details.append(
                    f"  + {m.strftime('%b %Y')}: "
                    + (
                        ", ".join(added[:20])
                        + (f", +{len(added)-20} more" if len(added) > 20 else "")
                    )
                )
            if removed:
                change_notes.append(
                    f"NOTE: {len(removed)} PC(s) removed starting {m.strftime('%B %Y')}."
                )
                change_details.append(
                    f"  - {m.strftime('%b %Y')}: "
                    + (
                        ", ".join(removed[:20])
                        + (f", +{len(removed)-20} more" if len(removed) > 20 else "")
                    )
                )
        prev_set = pcs_now

    # Filter by each month with its own required PC count (based on that month’s active PCs)
    for m in month_range:
        pcs_in_month = monthly_site_pcs.get(site, {}).get(m, set())
        monthly_pc_count = len(pcs_in_month)
        if monthly_pc_count == 0:
            continue

        monthly_required = int(np.ceil((pct_required / 100) * monthly_pc_count)) or 1
        month_info.append(
            f"- {m.strftime('%b %Y')}: {monthly_pc_count} PCs → Required: {monthly_required}"
        )

        month_start = m.to_timestamp()
        month_end = month_start + pd.offsets.MonthBegin(1)

        month_timeline = timeline[
            (timeline.index >= month_start) & (timeline.index < month_end)
        ]
        month_filtered = month_timeline[
            month_timeline["ActivePCs"] >= monthly_required
        ].copy()
        if not month_filtered.empty:
            filtered_parts.append(month_filtered)

    # If nothing qualifies, show a clear "no data" refresh on the plot and exit
    if not filtered_parts:
        root.after(
            0,
            lambda: status_label.configure(
                text="⚠ No data to show for selected range."
            ),
        )
        _show_no_data_title(site, start_date, end_date, pct_required, month_info)
        return

    filtered = pd.concat(filtered_parts)

    min_run = int(min_run_var.get() or 5)
    kept_idx = keep_only_continuous_runs(filtered.index, min_len=min_run)
    filtered = filtered.loc[kept_idx]

    if filtered.empty:
        root.after(
            0,
            lambda: status_label.configure(
                text=f"⚠ No {min_run}-minute continuous blocks found for selected range."
            ),
        )
        _show_no_data_title(
            site,
            start_date,
            end_date,
            pct_required,
            month_info,
            note=f"No {min_run}-minute continuous blocks met the threshold.",
        )
        return

    # ---------- Plot ----------
    ax.clear()

    # Hover annotation
    annotation = ax.annotate(
        text="",
        xy=(0, 0),
        xytext=(15, 15),
        textcoords="offset points",
        bbox=dict(boxstyle="round", fc="w"),
        arrowprops=dict(arrowstyle="->"),
        zorder=11,
    )
    annotation.set_visible(False)

    all_dates = pd.date_range(
        start=start_date, end=end_date - pd.Timedelta(days=1)
    ).date

    filtered.index.name = "Timestamp"
    filtered = filtered.reset_index()

    filtered["HourFloat"] = (
        filtered["Timestamp"].dt.hour + filtered["Timestamp"].dt.minute / 60.0
    )

    filtered["YFloat"] = filtered["Timestamp"].map(lambda x: x.toordinal())
    filtered["FullDateTime"] = filtered["Timestamp"]

    # Draw scatter
    # (no custom colors to keep default matplotlib style consistent)
    scatter = ax.scatter(
        filtered["HourFloat"], filtered["YFloat"], s=80, picker=20, zorder=1
    )
    filtered_data_for_hover = filtered

    # X axis
    xticks = np.arange(0, 24.01, 1)
    ax.set_xticks(xticks)
    ax.set_xticklabels([f"{int(h):02}:00" for h in xticks], rotation=45)

    # Y axis (dates)
    y_ticks = [d.toordinal() for d in all_dates]
    y_labels = [d.strftime("%b-%d") for d in all_dates]

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
    ax.set_ylim([y_ticks[-1], y_ticks[0]])  # invert

    # ---- Title now includes site, date range, and threshold ----
    end_inclusive = (end_date - pd.Timedelta(days=1)).date()
    ax.set_title(
        f"PC Usage at {site} from {start_date.date()} to {end_inclusive} (≥{pct_required}%)",
        fontsize=14,
    )

    ax.set_xlabel("Time of Day")
    ax.set_ylabel("Date")
    ax.grid(True, linestyle="--", alpha=0.3)
    fig.tight_layout()
    plot_canvas.draw()

    # ---------- Build right-side details ----------
    # Qualified minutes (for ranges)
    qualified_minutes = set(filtered["Timestamp"])

    all_logins = df[df["Location"] == site]
    unique_pcs = sorted(all_logins["Resource"].unique())

    pc_contributions = defaultdict(set)
    for pc in unique_pcs:
        pc_sessions = all_logins[all_logins["Resource"] == pc]
        for _, row in pc_sessions.iterrows():
            login, logout = row["Login Time"], row["Logout Time"]
            if pd.notna(login) and pd.notna(logout):
                session_minutes = pd.date_range(start=login, end=logout, freq="min")
                overlaps = qualified_minutes.intersection(session_minutes)
                if overlaps:
                    pc_contributions[pc].update(overlaps)

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

    # Monthly PC threshold table text
    table_lines = []
    table_lines.append(
        f"{'Month':<12} {'PCs':>5} {f'Required PCs at {pct_required}%':>20}"
    )
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

    # 👇 Add the names (Resources) that changed
    if change_details:
        sections.append("  ↳ PCs changed (by name):")
        sections.extend(change_details)

    sections.append(f"\n🕒 Qualified Time Ranges (≥{pct_required}% PCs active)")
    if ranges:
        for s, e in ranges[:50]:
            if s.date() == e.date():
                sections.append(
                    f"  • {s.strftime('%Y-%m-%d %H:%M')} → {e.strftime('%H:%M')}"
                )
            else:
                sections.append(
                    f"  • {s.strftime('%Y-%m-%d %H:%M')} → {e.strftime('%Y-%m-%d %H:%M')}"
                )
        if len(ranges) > 50:
            sections.append(f"  ...and {len(ranges) - 50} more.")
    else:
        sections.append("  ⚠ No qualified time ranges found.")

    status_detail_text.configure(state="normal")
    status_detail_text.delete("1.0", "end")
    status_detail_text.insert("end", "\n".join(sections))
    status_detail_text.configure(state="disabled")

    sync_details_popout()

    if details_popout_win is not None and details_popout_win.winfo_exists():
        details_popout_win.title(f"Details — {site}")

    root.after(0, lambda: status_label.configure(text="✔ Plot updated."))


def on_hover(event):
    global scatter, filtered_data_for_hover, annotation

    if scatter is None or filtered_data_for_hover is None:
        return

    if event.inaxes != ax:
        return

    contains, info = scatter.contains(event)
    if contains and info and "ind" in info and len(info["ind"]) > 0:
        ind_list = info["ind"]

        # Compute which point is closest to the mouse cursor
        mouse_x, mouse_y = event.xdata, event.ydata
        nearest_idx = min(
            ind_list,
            key=lambda i: (filtered_data_for_hover.iloc[i]["HourFloat"] - mouse_x) ** 2
            + (filtered_data_for_hover.iloc[i]["YFloat"] - mouse_y) ** 2,
        )

        row = filtered_data_for_hover.iloc[nearest_idx]

        x = row["HourFloat"]
        y = row["YFloat"]
        timestamp = row["FullDateTime"]

        print(f"[Hover Debug] Nearest Point Index: {nearest_idx}")
        print(f"[Hover Debug] HourFloat: {x}, YFloat: {y}")
        print(f"[Hover Debug] FullDateTime: {timestamp}")
        print(f"[Hover Debug] Row:\n{row}\n")

        # add the pc names active here
        site = site_var.get()
        all_logins_site = df[df["Location"] == site]
        mask = (all_logins_site["Login Time"] <= timestamp) & (
            all_logins_site["Logout Time"] >= timestamp
        )
        active_pcs = sorted(all_logins_site.loc[mask, "Resource"].unique())

        # Build hover text (limit long lists for readability)
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
        annotation.set_visible(False)
        fig.canvas.draw_idle()


def load_current_file():
    cfg = load_config()
    current_path = cfg.get("current_file") or os.path.join(
        get_data_dir(), "current.xlsx"
    )

    if not os.path.exists(current_path):
        root.after(
            0,
            lambda: status_label.configure(
                text="📂 No saved file. Click Import File to begin."
            ),
        )
        return

    cancelled = threading.Event()

    # ✅ IMPORTANT: reset on the main thread BEFORE starting thread
    reset_ui()

    loading_win = ctk.CTkToplevel(root)
    loading_win.title("Loading...")
    loading_win.geometry("300x140")
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
            file_hash = safe_read_retry(lambda: sha256_file(current_path))
            print("HASH:", file_hash)

            ok = load_processed_cache(file_hash)
            print("load_processed_cache returned:", ok)

            if ok:
                root.after(0, init_ui_from_df)
                global export_pending_after_load
                if export_pending_after_load:
                    export_pending_after_load = False
                    root.after(0, export_plot_data)
                root.after(
                    0,
                    lambda: status_label.configure(
                        text="⚡ Loaded saved file from cache."
                    ),
                )
                cfg = load_config()
                cfg["current_file"] = current_path
                cfg["last_hash"] = file_hash
                save_config(cfg)
                return

            load_data(current_path, progress, label, cancelled)

        except Exception as e:
            root.after(
                0, lambda: status_label.configure(text=f"❌ Failed to load: {e}")
            )
            root.after(0, reset_ui)
        finally:
            try:
                if loading_win.winfo_exists():
                    root.after(0, loading_win.destroy)
            except Exception:
                pass

    threading.Thread(target=do_load, daemon=True).start()


def init_ui_from_df():
    min_date = df["Login Time"].min().date()
    max_date = df["Login Time"].max().date()

    start_date_picker.config(mindate=min_date, maxdate=max_date)
    end_date_picker.config(mindate=min_date, maxdate=max_date)
    start_date_picker.set_date(min_date)
    end_date_picker.set_date(max_date)

    site_dropdown.configure(values=sorted(df["Location"].unique()))
    if df["Location"].nunique() > 0:
        site_var.set(sorted(df["Location"].unique())[0])


def _npz_scalar(x):
    """Return a python scalar from np.load() values (handles 0d/1d/2d arrays)."""
    if isinstance(x, np.ndarray):
        try:
            return x.item()
        except Exception:
            y = x.reshape(-1)[0]
            return y.item() if hasattr(y, "item") else y
    return x


def save_processed_cache(file_hash: str):
    # Uses your globals: df, site_timelines, monthly_site_pcs
    p = cache_paths_for_hash(file_hash)

    # df + monthly_site_pcs are fine pickled
    with open(p["df"], "wb") as f:
        pickle.dump(df, f, protocol=pickle.HIGHEST_PROTOCOL)

    with open(p["monthly"], "wb") as f:
        pickle.dump(monthly_site_pcs, f, protocol=pickle.HIGHEST_PROTOCOL)

    # Save site list
    sites = sorted(site_timelines.keys())
    with open(p["sites"], "wb") as f:
        pickle.dump(sites, f, protocol=pickle.HIGHEST_PROTOCOL)

    # Save timelines efficiently (only ActivePCs arrays + shared time index)
    # (requires numpy already imported)
    if not sites:
        raise ValueError("No sites to cache.")

    timeline0, _ = site_timelines[sites[0]]
    start_ns = int(timeline0.index[0].value)  # nanoseconds since epoch
    length = len(timeline0.index)

    arrays = {
        "start_ns": np.array([start_ns], dtype=np.int64),
        "length": np.array([length], dtype=np.int64),
    }

    for site in sites:
        timeline, _ = site_timelines[site]
        arrays[f"ap__{site}"] = timeline["ActivePCs"].to_numpy(dtype=np.int16)

    np.savez_compressed(p["timelines"], **arrays)


def load_processed_cache(file_hash: str) -> bool:
    global df, site_timelines, monthly_site_pcs

    if not cache_exists(file_hash):
        return False

    p = cache_paths_for_hash(file_hash)

    try:
        with open(p["df"], "rb") as f:
            df = pickle.load(f)

        with open(p["monthly"], "rb") as f:
            loaded = pickle.load(f)
            monthly_site_pcs.clear()
            monthly_site_pcs.update(loaded)

        with open(p["sites"], "rb") as f:
            sites = pickle.load(f)

        npz = np.load(p["timelines"], allow_pickle=False)
        start_ns = int(npz["start_ns"].reshape(-1)[0])
        start_ts = pd.Timestamp(start_ns, unit="ns")
        length = int(npz["length"].reshape(-1)[0])
        index = pd.date_range(start=start_ts, periods=length, freq="min")
        index.name = "Timestamp"

        site_timelines.clear()
        for site in sites:
            key = f"ap__{site}"
            if key not in npz.files:
                raise KeyError(f"Missing timeline key in npz: {key}")
            arr = npz[key]
            timeline = pd.DataFrame({"ActivePCs": arr}, index=index)
            num_pcs = int(df[df["Location"] == site]["Resource"].nunique())
            site_timelines[site] = (timeline, num_pcs)

        print("CACHE LOAD OK:", "df rows =", len(df), "sites =", len(site_timelines))
        return True

    except Exception as e:
        print("CACHE LOAD FAILED:", repr(e))
        # optional: delete corrupt cache so next run rebuilds
        for path in p.values():
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass
        df = None
        site_timelines.clear()
        monthly_site_pcs.clear()
        return False


def load_and_initialize():
    global export_pending_after_load
    original_path = filedialog.askopenfilename(
        title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not original_path:
        return

    try:
        current_path = copy_to_current(original_path)
        cfg = load_config()
        cfg["current_file"] = current_path
        cfg.pop("last_hash", None)  # force re-check
        save_config(cfg)

        export_pending_after_load = True
    except Exception as e:
        root.after(0, lambda: status_label.configure(text=f"❌ Import failed: {e}"))
        return

    # ✅ One loader only
    load_current_file()

    # cancelled = threading.Event()  # shared flag to cancel loading

    # Create loading window before thread starts
    # loading_win = ctk.CTkToplevel(root)
    # loading_win.title("Loading...")
    # loading_win.geometry("300x140")
    # loading_win.transient(root)
    # loading_win.grab_set()

    # def handle_cancel():
    #     cancelled.set()
    #     loading_win.destroy()
    #     reset_ui()

    # loading_win.protocol("WM_DELETE_WINDOW", handle_cancel)  # Close window

    # label = ctk.CTkLabel(loading_win, text="Processing Data... 0%")
    # label.pack(pady=(10, 5))

    # progress = ctk.CTkProgressBar(loading_win, mode="determinate")
    # progress.pack(fill="x", padx=20, pady=(0, 10))
    # progress.set(0)

    # cancel_btn = ctk.CTkButton(loading_win, text="❌ Cancel", command=handle_cancel)
    # cancel_btn.pack(pady=(0, 10))

    # def do_load():
    #     try:
    #         reset_ui()
    #         load_data(file_path, progress, label, cancelled)
    #     except Exception as e:
    #         root.after(0, lambda: status_label.configure(text=f"❌ Failed to load: {e}"))
    #         root.after(0, reset_ui)
    #     finally:
    #         if loading_win.winfo_exists():
    #             root.after(0, loading_win.destroy)

    # threading.Thread(target=do_load, daemon=True).start()


def reset_ui():
    global df, site_timelines
    df = None
    site_timelines.clear()
    monthly_site_pcs.clear()
    root.after(
        0, lambda: status_label.configure(text="📂 Please load an Excel file to begin.")
    )
    site_dropdown.configure(values=[])
    start_date_var.set("")
    end_date_var.set("")
    site_var.set("")
    status_detail_text.configure(state="normal")
    status_detail_text.delete("1.0", "end")
    status_detail_text.configure(state="disabled")
    ax.clear()
    ax.axis("off")
    plot_canvas.draw()


def load_data(file_path, progress_widget, label_widget, cancelled):
    global df, site_timelines, monthly_site_pcs

    # ---- 1) Detect which row contains the header ----
    import math

    REQUIRED = {"login time", "logout time", "resource", "location"}

    # Read a small preview without headers so we can search for the header row
    preview = pd.read_excel(file_path, header=None, nrows=50)  # adjust nrows if needed

    header_row_idx = None
    for i in range(len(preview)):
        row_vals = preview.iloc[i].astype(str).str.strip().str.lower()
        # Build a set of non-empty cells
        cells = set(v for v in row_vals if v and v != "nan")
        if REQUIRED.issubset(cells):
            header_row_idx = i
            break

    if header_row_idx is None:
        root.after(
            0,
            lambda: status_label.configure(
                text="❌ Could not find header row containing: Login Time, Logout Time, Resource, Location"
            ),
        )
        return

    # ---- 2) Re-read with the correct header row ----
    raw = pd.read_excel(file_path, header=header_row_idx)

    # Normalize column names (trim + lowercase) for matching,
    # but keep original names to preserve data.
    norm_to_orig = {str(c).strip().lower(): c for c in raw.columns}

    # Ensure required columns exist (case/space-insensitive)
    missing = [col for col in REQUIRED if col not in norm_to_orig]
    if missing:
        root.after(
            0,
            lambda: status_label.configure(
                text=f"❌ Missing required column(s) after header detection: {', '.join([m.title() for m in missing])}"
            ),
        )
        return

    # ---- 3) Keep only the needed columns and rename to canonical names ----
    df = raw[
        [
            norm_to_orig["login time"],
            norm_to_orig["logout time"],
            norm_to_orig["resource"],
            norm_to_orig["location"],
        ]
    ].copy()

    df.rename(
        columns={
            norm_to_orig["login time"]: "Login Time",
            norm_to_orig["logout time"]: "Logout Time",
            norm_to_orig["resource"]: "Resource",
            norm_to_orig["location"]: "Location",
        },
        inplace=True,
    )

    # ---- 4) Parse/clean + drop blanks (this prevents crashes) ----
    dt_format = "%m/%d/%Y %I:%M %p"
    df["Login Time"] = pd.to_datetime(
        df["Login Time"], format=dt_format, errors="coerce"
    )
    df["Logout Time"] = pd.to_datetime(
        df["Logout Time"], format=dt_format, errors="coerce"
    )
    df["Resource"] = df["Resource"].astype(str).str.strip()
    df["Location"] = df["Location"].astype(str).str.strip()

    # Drop rows with any required blanks or unparsable dates
    before = len(df)
    df.dropna(
        subset=["Login Time", "Logout Time", "Resource", "Location"], inplace=True
    )
    dropped = before - len(df)
    if dropped > 0:
        root.after(
            0,
            lambda: status_label.configure(
                text=f"⚠ Skipped {dropped} rows with missing required values."
            ),
        )

    # remove some parts of the express, laptop and outreach
    bad_locations = ["express", "outreach", "laptop"]
    df = df[~df["Location"].str.lower().str.contains("|".join(bad_locations))]

    # Optional: ensure Login <= Logout
    df = df[df["Logout Time"] >= df["Login Time"]]
    if df.empty:
        root.after(
            0, lambda: status_label.configure(text="⚠ No valid rows after cleaning.")
        )
        return

    # ---- 5) (your existing code continues unchanged from here) ----
    df["Month"] = df["Login Time"].dt.to_period("M")
    available_months = sorted(df["Month"].dropna().unique())

    min_date = df["Login Time"].min().date()
    max_date = df["Login Time"].max().date()
    root.after(0, lambda: start_date_picker.config(mindate=min_date, maxdate=max_date))
    root.after(0, lambda: end_date_picker.config(mindate=min_date, maxdate=max_date))

    root.after(0, lambda: start_date_picker.set_date(min_date))
    root.after(0, lambda: end_date_picker.set_date(max_date))

    start = df["Login Time"].min().replace(day=1)
    end = (df["Login Time"].max() + pd.Timedelta(days=1)).replace(
        day=1
    ) + pd.offsets.MonthEnd(0)
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

            # Track monthly PC presence
            if pd.notna(login):
                month = login.to_period("M")
                monthly_site_pcs[site][month].add(row["Resource"])

            if pd.notna(login) and pd.notna(logout):
                active_minutes = pd.date_range(start=login, end=logout, freq="min")
                active_minutes = active_minutes.intersection(timeline.index)  # align
                timeline.loc[active_minutes, "ActivePCs"] += 1

            processed_rows += 1
            if processed_rows % 50 == 0 or processed_rows == total_rows:
                val = processed_rows / total_rows
                root.after(0, lambda v=val: progress_widget.set(v))
                root.after(
                    0,
                    lambda v=val: label_widget.configure(
                        text=f"Processing data... {int(v*100)}%"
                    ),
                )

        site_timelines[site] = (timeline, num_pcs)

    root.after(
        0, lambda: site_dropdown.configure(values=sorted(df["Location"].unique()))
    )
    if df["Location"].nunique() > 0:
        root.after(0, lambda: site_var.set(sorted(df["Location"].unique())[0]))

    try:
        global export_pending_after_load

        # ✅ Export once after a NEW import, before caching
        if export_pending_after_load:
            export_pending_after_load = False
            root.after(0, export_plot_data)

        file_hash = sha256_file(file_path)
        save_processed_cache(file_hash)

        cfg = load_config()
        cfg["current_file"] = file_path
        cfg["last_hash"] = file_hash
        save_config(cfg)
    except Exception as e:
        root.after(
            0,
            lambda: status_label.configure(
                text=f"✔ File loaded. (Cache save failed: {e})"
            ),
        )
        return

    root.after(
        0,
        lambda: status_label.configure(
            text="✔ File loaded and cached. Ready to visualize data."
        ),
    )


def sync_details_popout():
    if details_popout_text is None:
        return
    if not details_popout_text.winfo_exists():
        return

    content = status_detail_text.get("1.0", "end-1c")
    details_popout_text.configure(state="normal")
    details_popout_text.delete("1.0", "end")
    details_popout_text.insert("end", content)
    details_popout_text.configure(state="disabled")


# ---------------- TKINTER UI ----------------
root = ctk.CTk()
root.title("PC Usage Visualizer")
root.geometry("1000x1000")
root.minsize(1000, 1000)

scrollable_frame = ctk.CTkScrollableFrame(root)
scrollable_frame.pack(fill="both", expand=True, padx=10, pady=10)

plot_frame = ctk.CTkFrame(scrollable_frame)
plot_frame.pack(fill="both", expand=True, padx=10, pady=(10, 0))
plot_frame.pack_propagate(True)  # Let children grow

fig, ax = plt.subplots(figsize=(12, 6))


plot_canvas = FigureCanvasTkAgg(fig, master=plot_frame)
fig.canvas.mpl_connect("motion_notify_event", on_hover)
canvas_widget = plot_canvas.get_tk_widget()
canvas_widget.pack(fill="both", expand=True)  # Auto-scaling!

plot_frame.pack(fill="both", expand=True, padx=10, pady=(10, 5))  # ⬅ add bottom padding

toolbar_container = tk.Frame(scrollable_frame)  # Not inside plot_frame
toolbar_container.pack(fill=tk.X, padx=10, pady=(0, 10))

toolbar = NavigationToolbar2Tk(plot_canvas, toolbar_container)
toolbar.update()
toolbar.pack(side=tk.BOTTOM, fill=tk.X)


ax.clear()
ax.axis("off")
plot_canvas.draw()

status_label = ctk.CTkLabel(
    scrollable_frame, text="📂 Please load an Excel file to begin.", text_color="gray"
)
status_label.pack(fill=ctk.X, padx=10, pady=(5, 10))

load_btn = ctk.CTkButton(
    scrollable_frame, text="Import File", command=load_and_initialize
)
load_btn.pack(pady=(0, 10))

detail_container = ctk.CTkFrame(scrollable_frame, corner_radius=16)
detail_container.pack(fill=ctk.X, padx=10, pady=(0, 10))

# Header bar
detail_header = ctk.CTkFrame(detail_container, fg_color="transparent")
detail_header.pack(fill="x", padx=14, pady=(12, 6))

details_title = ctk.CTkLabel(
    detail_header,
    text="Details",
    font=ctk.CTkFont(size=16, weight="bold"),
)
details_title.pack(side="left")

details_subtitle = ctk.CTkLabel(
    detail_header,
    text="",
    text_color="#9aa0a6",
    font=ctk.CTkFont(size=12),
)
details_subtitle.pack(side="left", padx=(12, 0))


def copy_details():
    try:
        txt = status_detail_text.get("1.0", "end-1c")
        root.clipboard_clear()
        root.clipboard_append(txt)
        status_label.configure(text="📋 Details copied to clipboard.")
    except Exception:
        pass


ctk.CTkButton(detail_header, text="📋 Copy", width=90, command=copy_details).pack(
    side="right", padx=(8, 0)
)
ctk.CTkButton(
    detail_header, text="⤢ Pop Out", width=100, command=open_details_popout
).pack(side="right")

# Separator line
sep = ctk.CTkFrame(detail_container, height=1, fg_color="#2a2a2a")
sep.pack(fill="x", padx=14, pady=(0, 10))

# Text area (modern)
status_detail_text = ctk.CTkTextbox(
    detail_container,
    wrap="word",
    font=("Segoe UI", 12),
    corner_radius=12,
    height=220,
)
status_detail_text.pack(fill="both", expand=True, padx=14, pady=(0, 14))
status_detail_text.configure(state="disabled")


frame = ctk.CTkFrame(scrollable_frame)
frame.pack(side=ctk.TOP, fill=ctk.X, padx=10, pady=10)
frame.grid_columnconfigure(0, weight=1)
frame.grid_columnconfigure(1, weight=1)


left_frame = ctk.CTkFrame(frame)
left_frame.grid(row=0, column=0, sticky="w")

ctk.CTkLabel(left_frame, text="Start Date:").pack(side=ctk.LEFT, padx=(0, 5))
start_date_var = tk.StringVar()
start_date_picker = DateEntry(
    left_frame,
    textvariable=start_date_var,
    width=12,
    background="darkblue",
    foreground="white",
    borderwidth=2,
)
start_date_picker.pack(side=ctk.LEFT)

ctk.CTkLabel(left_frame, text="  End Date:").pack(side=ctk.LEFT, padx=(10, 5))
end_date_var = tk.StringVar()
end_date_picker = DateEntry(
    left_frame,
    textvariable=end_date_var,
    width=12,
    background="darkblue",
    foreground="white",
    borderwidth=2,
)
end_date_picker.pack(side=ctk.LEFT)


ctk.CTkLabel(left_frame, text="   Select Branch:").pack(side=ctk.LEFT, padx=(10, 5))
site_var = ctk.StringVar()
site_dropdown = ttk.Combobox(left_frame, textvariable=site_var, state="readonly")
site_dropdown.pack(side=ctk.LEFT)


apply_btn = ctk.CTkButton(
    frame,
    text="✅ Apply",
    command=update_plot,
    fg_color="green",  # This sets the button's fill color
    hover_color="#006400",  # Optional: a darker green on hover
)
apply_btn.grid(row=0, column=2, padx=(20, 0), sticky="e")

# export_btn = ctk.CTkButton(
#     frame, text="📤 Export for PowerBI", command=export_plot_data
# )
# export_btn.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(10, 0))


percent_var = tk.IntVar()
min_run_var = tk.IntVar(value=5)
slider_frame = ctk.CTkFrame(frame)
slider_frame.grid(row=0, column=1, padx=(20, 0), sticky="e")


slider_label = ctk.CTkLabel(
    slider_frame,
    text="Select Threshold:",
    text_color="white",
    font=ctk.CTkFont(size=12),
)
slider_label.pack(pady=(0, 5))

percent_slider = ctk.CTkSlider(
    slider_frame,
    from_=10,
    to=100,
    number_of_steps=9,
    variable=percent_var,
    command=lambda val: percent_var.set(int(float(val))),
)
percent_slider.set(100)
percent_slider.pack(fill="x")

tick_frame = ctk.CTkFrame(slider_frame)
tick_frame.pack(fill="x", pady=(2, 0))

tick_frame.grid_columnconfigure(tuple(range(10)), weight=1)


for idx, i in enumerate(range(10, 101, 10)):
    lbl = ctk.CTkLabel(
        tick_frame, text=str(i), text_color="white", font=ctk.CTkFont(size=10)
    )
    lbl.grid(row=0, column=idx, sticky="n")



ctk.CTkLabel(
    slider_frame,
    text="Min Continuous Minutes:",
    text_color="white",
    font=ctk.CTkFont(size=12),
).pack(pady=(10, 5))

ALLOWED_MIN_RUNS = [5, 10, 15, 20, 30, 40, 50, 60]

def _snap_min_run(v):
    v = int(v)
    snapped = min(ALLOWED_MIN_RUNS, key=lambda x: abs(x - v))
    min_run_var.set(snapped)

# --- Min continuous minutes slider ---
min_run_slider = ctk.CTkSlider(
    slider_frame,
    from_=min(ALLOWED_MIN_RUNS),
    to=max(ALLOWED_MIN_RUNS),
    number_of_steps=len(ALLOWED_MIN_RUNS) - 1,
    variable=min_run_var,
    command=_snap_min_run,
)
min_run_slider.set(5)
min_run_slider.pack(fill="x")

# Tick labels
min_tick_frame = ctk.CTkFrame(slider_frame)
min_tick_frame.pack(fill="x", pady=(2, 0))
min_tick_frame.grid_columnconfigure(tuple(range(len(ALLOWED_MIN_RUNS))), weight=1)

for idx, val in enumerate(ALLOWED_MIN_RUNS):
    ctk.CTkLabel(
        min_tick_frame,
        text=str(val),
        text_color="white",
        font=ctk.CTkFont(size=10),
    ).grid(row=0, column=idx)




load_current_file()

root.mainloop()
