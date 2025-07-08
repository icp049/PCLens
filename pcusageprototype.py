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



# Global state
df = None
site_timelines = {}

def export_plot_data():
    if not site_timelines:
        status_label.config(text="⚠ Cannot export — no data loaded.")
        return

    export_rows = []
    thresholds = list(range(10, 101, 10))
    all_months = sorted(df['Login Time'].dt.to_period('M').dropna().unique())

    # ✅ Prompt file path only once here
    export_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Save Threshold Plot Data"
    )

    if not export_path:
        status_label.configure(text="❌ Export canceled.")
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
        status_label.configure(text="❌ Export canceled.")

    cancel_btn = ctk.CTkButton(export_win, text="❌ Cancel", command=handle_cancel)
    cancel_btn.pack(pady=(0, 10))

    def do_export():
        try:
            # First pass to count total minutes to export
            total_minutes = 0
            for month in all_months:
                month_start = month.to_timestamp()
                month_end = (month_start + pd.offsets.MonthBegin(1))
                for site, (timeline, num_pcs) in site_timelines.items():
                    filtered_timeline = timeline[(timeline.index >= month_start) & (timeline.index < month_end)]
                    for pct in thresholds:
                        required_count = int(np.ceil((pct / 100) * num_pcs)) or 1
                        total_minutes += len(filtered_timeline[filtered_timeline['ActivePCs'] >= required_count])

            if total_minutes == 0:
                root.after(0, lambda: status_label.configure(text="⚠ No data to export."))
                return

            completed_minutes = 0

            for month in all_months:
                if cancel_flag.is_set(): return
                month_start = month.to_timestamp()
                month_end = (month_start + pd.offsets.MonthBegin(1))
                month_str = str(month)

                for site, (timeline, num_pcs) in site_timelines.items():
                    if cancel_flag.is_set(): return

                    filtered_timeline = timeline[(timeline.index >= month_start) & (timeline.index < month_end)]
                    all_logins = df[(df['Site'] == site) &
                                    (df['Login Time'] >= month_start) &
                                    (df['Login Time'] < month_end)]
                    unique_pcs = sorted(all_logins['Resource'].unique())

                    for pct in thresholds:
                        if cancel_flag.is_set(): return

                        required_count = int(np.ceil((pct / 100) * num_pcs)) or 1
                        filtered_minutes = filtered_timeline[filtered_timeline['ActivePCs'] >= required_count].index

                        for minute in filtered_minutes:
                            export_rows.append({
                                'Branch': site,
                                'Threshold (%)': pct,
                                'Timestamp': minute,
                                'PCs Used': f"{required_count} of {len(unique_pcs)}",
                                'Month': month_str
                            })
                            completed_minutes += 1
                            if completed_minutes % 100 == 0 or completed_minutes == total_minutes:
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






def update_plot():
    site = site_var.get()
    month_str = month_var.get()
    if not site or site not in site_timelines or not month_str:
        return

    selected_month = pd.Period(month_str).to_timestamp()
    month_start = selected_month.replace(day=1)
    month_end = (month_start + pd.offsets.MonthBegin(1))

    timeline, num_pcs = site_timelines[site]
    timeline = timeline[(timeline.index >= month_start) & (timeline.index < month_end)]
    pct_required = percent_var.get()
    required_count = int(np.ceil((pct_required / 100) * num_pcs)) or 1

    filtered = timeline[timeline['ActivePCs'] >= required_count].copy()
    ax.clear()

    all_dates = pd.date_range(start=month_start, end=month_end - pd.Timedelta(days=1)).date

    if filtered.empty:
        ax.set_title(f"[{site}] No overlaps with ≥{pct_required}% active PCs", fontsize=14)
        ax.set_xticks(np.arange(0, 24, 1))
        ax.set_xticklabels([f"{h:02}:00" for h in range(24)], rotation=45)
        ax.set_ylim([all_dates[-1], all_dates[0]])
        ax.set_yticks(all_dates)
        ax.set_yticklabels([d.strftime('%b-%d') for d in all_dates], fontsize=8)
        ax.grid(True, linestyle='--', alpha=0.3)
        fig.tight_layout()
        plot_canvas.draw()
        status_label.configure(text="⚠ No data to show")
        status_detail_label.configure(text="")
        return

    filtered['Date'] = filtered.index.date
    filtered['HourFloat'] = filtered.index.hour + filtered.index.minute / 60.0
    ax.scatter(filtered['HourFloat'], filtered['Date'], s=5, color='green')
    xticks = np.arange(0, 24.01, 1)
    ax.set_xticks(xticks)
    ax.set_xticklabels([f"{int(h):02}:00" for h in xticks], rotation=45)
    ax.set_ylim([all_dates[-1], all_dates[0]])
    ax.set_yticks(all_dates)
    ax.set_yticklabels([d.strftime('%b-%d') for d in all_dates], fontsize=8)
    ax.set_title(f"[{site}] Time of Day with {pct_required}% PCs Active", fontsize=14)
    ax.set_xlabel("Time of Day")
    ax.set_ylabel("Date")
    ax.grid(True, linestyle='--', alpha=0.3)
    fig.tight_layout()
    plot_canvas.draw()

    qualified_minutes = set(filtered.index)
    all_logins = df[df['Site'] == site]
    unique_pcs = sorted(all_logins['Resource'].unique())

    pc_contributions = defaultdict(set)
    for pc in unique_pcs:
        pc_sessions = all_logins[all_logins['Resource'] == pc]
        for _, row in pc_sessions.iterrows():
            login, logout = row['Login Time'], row['Logout Time']
            if pd.notna(login) and pd.notna(logout):
                session_minutes = pd.date_range(start=login, end=logout, freq='min')
                overlaps = qualified_minutes.intersection(session_minutes)
                if overlaps:
                    pc_contributions[pc].update(overlaps)

    pc_contribution_counts = {pc: len(minutes) for pc, minutes in pc_contributions.items()}
    sorted_pcs = sorted(pc_contribution_counts.items(), key=lambda x: x[1], reverse=True)
    active_pc_names = [pc for pc, _ in sorted_pcs[:required_count]]

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

    range_lines = []
    for s, e in ranges[:50]:
        same_day = s.date() == e.date()
        if same_day:
            range_lines.append(f"- {s.strftime('%Y-%m-%d %H:%M')} to {e.strftime('%H:%M')}")
        else:
            range_lines.append(f"- {s.strftime('%Y-%m-%d %H:%M')} to {e.strftime('%Y-%m-%d %H:%M')}")
    if len(ranges) > 50:
        range_lines.append(f"...and {len(ranges) - 50} more.")

    time_block = "\n\n🕒 Qualified Time Ranges (≥{}% PCs active):\n".format(pct_required)
    time_block += "\n".join(range_lines)

    status_detail_label.configure(
        text=f"🖥 {len(active_pc_names)}/{len(unique_pcs)} PCs contributed during qualified time blocks (≥{pct_required}%)\n" + time_block
    )
    status_label.configure(text="✔ Plot updated.")



def load_and_initialize():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not file_path:
        return  # Cancelled

    cancelled = threading.Event()  # shared flag to cancel loading

    # Create loading window before thread starts
    loading_win = ctk.CTkToplevel(root)
    loading_win.title("Loading...")
    loading_win.geometry("300x140")
    loading_win.transient(root)
    loading_win.grab_set()

    def handle_cancel():
        cancelled.set()
        loading_win.destroy()
        reset_ui()

    loading_win.protocol("WM_DELETE_WINDOW", handle_cancel)  # Close window

    label = ctk.CTkLabel(loading_win, text="Processing Data... 0%")
    label.pack(pady=(10, 5))

    progress = ctk.CTkProgressBar(loading_win, mode="determinate")
    progress.pack(fill="x", padx=20, pady=(0, 10))
    progress.set(0)

    cancel_btn = ctk.CTkButton(loading_win, text="❌ Cancel", command=handle_cancel)
    cancel_btn.pack(pady=(0, 10))

    def do_load():
        try:
            reset_ui() 
            load_data(file_path, progress, label, cancelled)
        except Exception as e:
            root.after(0, lambda: status_label.configure(text=f"❌ Failed to load: {e}"))
            root.after(0, reset_ui)
        finally:
            if loading_win.winfo_exists():
                root.after(0, loading_win.destroy)

    threading.Thread(target=do_load, daemon=True).start()
    
def reset_ui():
    global df, site_timelines
    df = None
    site_timelines.clear()
    status_label.configure(text="📂 Please load an Excel file to begin.")
    month_dropdown.configure(values=[])
    site_dropdown.configure(values=[])
    month_var.set("")
    site_var.set("")
    status_detail_label.configure(text="")
    ax.clear()
    ax.axis('off')
    plot_canvas.draw()

    
def load_data(file_path, progress_widget, label_widget, cancelled):
    global df, site_timelines

    df = pd.read_excel(file_path)
    dt_format = '%m/%d/%Y %I:%M %p'
    df['Login Time'] = pd.to_datetime(df['Login Time'], format=dt_format, errors='coerce')
    df['Logout Time'] = pd.to_datetime(df['Logout Time'], format=dt_format, errors='coerce')
    df['Resource'] = df['Resource'].astype(str).str.strip()
    df['Site'] = df['Site'].astype(str).str.strip()

    df['Month'] = df['Login Time'].dt.to_period('M')
    available_months = sorted(df['Month'].dropna().unique())

    root.after(0, lambda: month_dropdown.configure(values=[str(m) for m in available_months]))
    if available_months:
        root.after(0, lambda: month_var.set(str(available_months[0])))

    start = df['Login Time'].min().replace(day=1)
    end = (df['Login Time'].max() + pd.Timedelta(days=1)).replace(day=1) + pd.offsets.MonthEnd(0)
    full_range = pd.date_range(start=start, end=end, freq='min')
    timeline_template = pd.DataFrame({'Timestamp': full_range})
    timeline_template.set_index('Timestamp', inplace=True)

    site_timelines.clear()

    # Setup progress tracking
    total_rows = len(df)
    processed_rows = 0

    for site in df['Site'].unique():
        if cancelled.is_set():
           return  # Exit early if user cancelled
        
        site_df = df[df['Site'] == site]
        pcs = sorted(site_df['Resource'].unique())
        num_pcs = len(pcs)

        timeline = timeline_template.copy()
        timeline['ActivePCs'] = 0

        for _, row in site_df.iterrows():
            if cancelled.is_set():
              return 
            login, logout = row['Login Time'], row['Logout Time']
            if pd.notna(login) and pd.notna(logout):
                active_minutes = pd.date_range(start=login, end=logout, freq='min')
                for minute in active_minutes:
                    if minute in timeline.index:
                        timeline.loc[minute, 'ActivePCs'] += 1

            processed_rows += 1
            if processed_rows % 50 == 0 or processed_rows == total_rows:
                progress_val = processed_rows / total_rows
                root.after(0, lambda val=progress_val: progress_widget.set(val))
                root.after(0, lambda val=progress_val: label_widget.configure(
                    text=f"Processing data... {int(val * 100)}%"))

        site_timelines[site] = (timeline, num_pcs)

    # Final GUI updates
    root.after(0, lambda: site_dropdown.configure(values=sorted(df['Site'].unique())))
    if df['Site'].nunique() > 0:
        root.after(0, lambda: site_var.set(sorted(df['Site'].unique())[0]))
    root.after(0, lambda: status_label.configure(text="✔ File loaded. Ready to visualize data."))


# ---------------- TKINTER UI ----------------
root = ctk.CTk()
root.title("PC Usage Visualizer")
root.geometry("1000x1000")
root.minsize(1000, 1000)

scrollable_frame = ctk.CTkScrollableFrame(root)
scrollable_frame.pack(fill="both", expand=True, padx=10, pady=10)

plot_frame = ctk.CTkFrame(scrollable_frame, width=1200, height=600)
plot_frame.pack(fill=ctk.X, padx=10)
plot_frame.pack_propagate(False)

fig, ax = plt.subplots(figsize=(12, 6))
plot_canvas = FigureCanvasTkAgg(fig, master=plot_frame)
canvas_widget = plot_canvas.get_tk_widget()
canvas_widget.pack()
canvas_widget.config(width=1200, height=600)

ax.clear()
ax.axis('off')
plot_canvas.draw()

status_label = ctk.CTkLabel(scrollable_frame, text="📂 Please load an Excel file to begin.", text_color="gray")
status_label.pack(fill=ctk.X, padx=10, pady=(5, 10))

load_btn = ctk.CTkButton(scrollable_frame, text="Import File", command=load_and_initialize)
load_btn.pack(pady=(0, 10))

detail_container = ctk.CTkFrame(scrollable_frame, height=150)
detail_container.pack(fill=ctk.X, padx=10, pady=(0, 10))
detail_container.pack_propagate(False)

status_detail_scroll = ctk.CTkScrollableFrame(detail_container)
status_detail_scroll.pack(fill=ctk.BOTH, expand=True)

status_detail_label = ctk.CTkLabel(status_detail_scroll, text="", text_color="white", justify="left", anchor="w")
status_detail_label.pack(fill=ctk.X, padx=10, pady=10)

frame = ctk.CTkFrame(scrollable_frame)
frame.pack(side=ctk.TOP, fill=ctk.X, padx=10, pady=10)
frame.grid_columnconfigure(0, weight=1)
frame.grid_columnconfigure(1, weight=1)


left_frame = ctk.CTkFrame(frame)
left_frame.grid(row=0, column=0, sticky='w')

ctk.CTkLabel(left_frame, text="Select Month:").pack(side=ctk.LEFT, padx=(0, 5))
month_var = ctk.StringVar()
month_dropdown = ttk.Combobox(left_frame, textvariable=month_var, state='readonly')
month_dropdown.pack(side=ctk.LEFT)

ctk.CTkLabel(left_frame, text="   Select Branch:").pack(side=ctk.LEFT, padx=(10, 5))
site_var = ctk.StringVar()
site_dropdown = ttk.Combobox(left_frame, textvariable=site_var, state='readonly')
site_dropdown.pack(side=ctk.LEFT)


apply_btn = ctk.CTkButton(
    frame,
    text="✅ Apply",
    command=update_plot,
    fg_color="green",       # This sets the button's fill color
    hover_color="#006400"   # Optional: a darker green on hover
)
apply_btn.grid(row=0, column=2, padx=(20, 0), sticky='e')

export_btn = ctk.CTkButton(frame, text="📤 Export for PowerBI", command=export_plot_data)
export_btn.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(10, 0))


percent_var = tk.IntVar()
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
