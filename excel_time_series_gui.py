"""GUI app to visualize time-series data from an Excel file."""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from pandas.api.types import is_numeric_dtype, is_datetime64_any_dtype


class TimeSeriesExcelViewer:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel Time Series Viewer")
        self.root.geometry("1200x760")
        self.root.minsize(1000, 640)

        self.excel_path = tk.StringVar(
            value=r"d:\Digitaltwinimplementation\Code_test\waterflow_check\CMU814.xlsx"
        )
        self.sheet_var = tk.StringVar()
        self.time_col_var = tk.StringVar()
        self.value_col_var = tk.StringVar()
        self.start_time_var = tk.StringVar()
        self.end_time_var = tk.StringVar()
        self.y_min_var = tk.StringVar()
        self.y_max_var = tk.StringVar()
        self.crop_filename_var = tk.StringVar(value="cropped_output.xlsx")
        self.auto_y_var = tk.BooleanVar(value=True)

        self.df_all = None
        self.df_plot = None
        self.figure = None
        self.canvas = None
        self.toolbar = None
        self.time_is_datetime = False
        self.current_ax = None
        self.current_time_values = []
        self.slider_sync_in_progress = False

        self._build_ui()
        self.load_excel()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        main.rowconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)

        control = ttk.LabelFrame(main, text="Controls", padding=10)
        control.grid(row=0, column=0, sticky="nsw", padx=(0, 10))

        plot_frame = ttk.LabelFrame(main, text="Plot", padding=6)
        plot_frame.grid(row=0, column=1, sticky="nsew")
        plot_frame.rowconfigure(0, weight=1)
        plot_frame.columnconfigure(0, weight=1)
        self.plot_container = plot_frame

        r = 0
        ttk.Label(control, text="Excel file:").grid(row=r, column=0, sticky="w")
        ttk.Entry(control, textvariable=self.excel_path, width=44).grid(
            row=r + 1, column=0, columnspan=2, sticky="ew", pady=(2, 4)
        )
        ttk.Button(control, text="Browse...", command=self.choose_excel).grid(
            row=r + 1, column=2, padx=(6, 0), sticky="ew"
        )
        ttk.Button(control, text="Load", command=self.load_excel).grid(
            row=r + 2, column=0, columnspan=3, sticky="ew", pady=(0, 8)
        )

        r += 3
        ttk.Label(control, text="Sheet:").grid(row=r, column=0, sticky="w")
        self.sheet_combo = ttk.Combobox(
            control, textvariable=self.sheet_var, state="readonly", width=36
        )
        self.sheet_combo.grid(row=r + 1, column=0, columnspan=2, sticky="ew", pady=(2, 6))
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _e: self.load_sheet())

        r += 2
        ttk.Label(control, text="Time column:").grid(row=r, column=0, sticky="w")
        self.time_combo = ttk.Combobox(
            control, textvariable=self.time_col_var, state="readonly", width=36
        )
        self.time_combo.grid(row=r + 1, column=0, columnspan=2, sticky="ew", pady=(2, 6))
        self.time_combo.bind("<<ComboboxSelected>>", lambda _e: self.update_time_bounds())

        r += 2
        ttk.Label(control, text="Value column:").grid(row=r, column=0, sticky="w")
        self.value_combo = ttk.Combobox(
            control, textvariable=self.value_col_var, state="readonly", width=36
        )
        self.value_combo.grid(row=r + 1, column=0, columnspan=2, sticky="ew", pady=(2, 8))

        r += 2
        ttk.Separator(control, orient="horizontal").grid(
            row=r, column=0, columnspan=3, sticky="ew", pady=6
        )
        r += 1

        ttk.Label(control, text="Start time:").grid(row=r, column=0, sticky="w")
        ttk.Label(control, text="End time:").grid(row=r, column=1, sticky="w")
        ttk.Entry(control, textvariable=self.start_time_var, width=18).grid(
            row=r + 1, column=0, sticky="ew", pady=(2, 6), padx=(0, 4)
        )
        ttk.Entry(control, textvariable=self.end_time_var, width=18).grid(
            row=r + 1, column=1, sticky="ew", pady=(2, 6), padx=(4, 0)
        )

        r += 2
        ttk.Label(control, text="Range sliders (start/end):").grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(0, 2)
        )
        r += 1
        self.start_idx_scale = tk.Scale(
            control,
            from_=0,
            to=1,
            orient="horizontal",
            resolution=1,
            label="Start",
            command=self._on_start_slider,
        )
        self.start_idx_scale.grid(row=r, column=0, columnspan=3, sticky="ew", pady=(0, 4))
        r += 1
        self.end_idx_scale = tk.Scale(
            control,
            from_=0,
            to=1,
            orient="horizontal",
            resolution=1,
            label="End",
            command=self._on_end_slider,
        )
        self.end_idx_scale.grid(row=r, column=0, columnspan=3, sticky="ew", pady=(0, 8))

        r += 1
        ttk.Checkbutton(
            control, text="Auto Y range", variable=self.auto_y_var
        ).grid(row=r, column=0, sticky="w")
        r += 1

        ttk.Label(control, text="Y min:").grid(row=r, column=0, sticky="w")
        ttk.Label(control, text="Y max:").grid(row=r, column=1, sticky="w")
        ttk.Entry(control, textvariable=self.y_min_var, width=18).grid(
            row=r + 1, column=0, sticky="ew", pady=(2, 8), padx=(0, 4)
        )
        ttk.Entry(control, textvariable=self.y_max_var, width=18).grid(
            row=r + 1, column=1, sticky="ew", pady=(2, 8), padx=(4, 0)
        )

        r += 2
        ttk.Button(control, text="Plot", command=self.plot_data).grid(
            row=r, column=0, columnspan=3, sticky="ew", pady=(0, 6)
        )
        ttk.Button(control, text="Reset ranges", command=self.update_time_bounds).grid(
            row=r + 1, column=0, columnspan=3, sticky="ew"
        )
        ttk.Button(control, text="Crop", command=self.crop_data).grid(
            row=r + 2, column=0, columnspan=3, sticky="ew", pady=(6, 0)
        )

        r += 3
        ttk.Label(control, text="Crop file name:").grid(row=r, column=0, sticky="w")
        ttk.Entry(control, textvariable=self.crop_filename_var, width=28).grid(
            row=r, column=1, columnspan=2, sticky="ew", padx=(4, 0)
        )

        for c in range(3):
            control.columnconfigure(c, weight=1)

        self.info_label = ttk.Label(
            main,
            text="Tip: Leave start/end empty to use full range.",
            foreground="#444",
            anchor="w",
        )
        self.info_label.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))

    def choose_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.excel_path.set(path)

    def load_excel(self):
        path = self.excel_path.get().strip()
        if not path:
            messagebox.showerror("Error", "Please provide an Excel file path.")
            return

        try:
            excel = pd.ExcelFile(path)
            sheets = excel.sheet_names
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to open Excel file:\n{exc}")
            return

        if not sheets:
            messagebox.showerror("Error", "No sheets found in this Excel file.")
            return

        self.sheet_combo["values"] = sheets
        self.sheet_var.set(sheets[0])
        self.load_sheet()

    def load_sheet(self):
        try:
            self.df_all = pd.read_excel(self.excel_path.get().strip(), sheet_name=self.sheet_var.get())
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to read selected sheet:\n{exc}")
            return

        if self.df_all is None or self.df_all.empty:
            messagebox.showwarning("Warning", "Selected sheet is empty.")
            return

        columns = [str(c) for c in self.df_all.columns]
        self.time_combo["values"] = columns
        self.value_combo["values"] = columns

        default_time = "time" if "time" in columns else columns[0]
        self.time_col_var.set(default_time)

        value_candidates = [c for c in columns if c != default_time]
        self.value_col_var.set(value_candidates[0] if value_candidates else default_time)

        self.update_time_bounds()

    def _parse_time_input(self, text: str):
        text = text.strip()
        if text == "":
            return None

        if self.time_is_datetime:
            return pd.to_datetime(text)

        try:
            return float(text)
        except ValueError:
            return pd.to_datetime(text)

    def _parse_time_input_with_type(self, text: str, is_datetime: bool):
        text = text.strip()
        if text == "":
            return None
        if is_datetime:
            return pd.to_datetime(text)
        return float(text)

    def _parse_time_series(self, raw_time: pd.Series):
        if is_datetime64_any_dtype(raw_time):
            return pd.to_datetime(raw_time, errors="coerce"), True
        if is_numeric_dtype(raw_time):
            return pd.to_numeric(raw_time, errors="coerce"), False

        numeric_time = pd.to_numeric(raw_time, errors="coerce")
        datetime_time = pd.to_datetime(raw_time, errors="coerce")
        numeric_ratio = numeric_time.notna().mean()
        datetime_ratio = datetime_time.notna().mean()

        if numeric_ratio >= 0.9 and numeric_ratio >= datetime_ratio:
            return numeric_time, False
        if datetime_ratio >= 0.9:
            return datetime_time, True
        raise ValueError("Time column must be numeric or datetime-like.")

    def _clean_plot_df(self):
        if self.df_all is None:
            raise ValueError("No data loaded.")

        time_col = self.time_col_var.get()
        value_col = self.value_col_var.get()
        if not time_col or not value_col:
            raise ValueError("Please select time and value columns.")

        if time_col not in self.df_all.columns or value_col not in self.df_all.columns:
            raise ValueError("Selected columns are not in the current sheet.")

        df = self.df_all[[time_col, value_col]].copy()
        df = df.dropna(subset=[time_col, value_col])

        parsed_time, is_datetime = self._parse_time_series(df[time_col])
        df[time_col] = parsed_time
        self.time_is_datetime = is_datetime

        df[value_col] = pd.to_numeric(df[value_col], errors="coerce")
        df = df.dropna(subset=[time_col, value_col]).sort_values(by=time_col)
        if df.empty:
            raise ValueError("No valid numeric data in selected columns.")
        return df, time_col, value_col

    def crop_data(self):
        try:
            if self.df_all is None:
                raise ValueError("No data loaded.")

            time_col = self.time_col_var.get()
            if not time_col:
                raise ValueError("Please select a time column.")
            if time_col not in self.df_all.columns:
                raise ValueError("Selected time column is not in the current sheet.")

            parsed_time, is_datetime = self._parse_time_series(self.df_all[time_col])
            start_val = self._parse_time_input_with_type(self.start_time_var.get(), is_datetime)
            end_val = self._parse_time_input_with_type(self.end_time_var.get(), is_datetime)

            mask = parsed_time.notna()
            if start_val is not None:
                mask &= parsed_time >= start_val
            if end_val is not None:
                mask &= parsed_time <= end_val

            cropped = self.df_all.loc[mask].copy()
            if cropped.empty:
                raise ValueError("No rows remain after applying time filter.")

            source_path = self.excel_path.get().strip()
            base_name = os.path.splitext(os.path.basename(source_path))[0]
            configured_name = self.crop_filename_var.get().strip()
            if configured_name:
                default_name = configured_name
                if not default_name.lower().endswith(".xlsx"):
                    default_name += ".xlsx"
            else:
                default_name = f"{base_name}_{self.sheet_var.get()}_cropped.xlsx"
            save_path = filedialog.asksaveasfilename(
                title="Save cropped Excel file",
                defaultextension=".xlsx",
                initialfile=default_name,
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            )
            if not save_path:
                return

            cropped.to_excel(save_path, index=False, sheet_name=str(self.sheet_var.get())[:31] or "Cropped")
            self.crop_filename_var.set(os.path.basename(save_path))
            self.info_label.config(
                text=f"Cropped {len(cropped)} rows and saved to '{os.path.basename(save_path)}'.",
                foreground="#1f6f43",
            )
        except Exception as exc:
            messagebox.showerror("Crop Error", str(exc))
            self.info_label.config(text="Crop failed. Check time range and selections.", foreground="#b42318")

    def update_time_bounds(self):
        try:
            df, time_col, _value_col = self._clean_plot_df()
        except Exception:
            return

        self.current_time_values = list(df[time_col].values)
        t_min = df[time_col].min()
        t_max = df[time_col].max()

        if self.time_is_datetime:
            self.start_time_var.set(str(t_min))
            self.end_time_var.set(str(t_max))
        else:
            self.start_time_var.set(f"{float(t_min):g}")
            self.end_time_var.set(f"{float(t_max):g}")

        self.y_min_var.set("")
        self.y_max_var.set("")
        self._setup_sliders()

    def _setup_sliders(self):
        count = len(self.current_time_values)
        if count == 0:
            return

        max_idx = count - 1
        self.slider_sync_in_progress = True
        self.start_idx_scale.configure(to=max_idx)
        self.end_idx_scale.configure(to=max_idx)
        self.start_idx_scale.set(0)
        self.end_idx_scale.set(max_idx)
        self.slider_sync_in_progress = False

    def _format_time_value(self, value):
        if self.time_is_datetime:
            return pd.Timestamp(value).strftime("%Y-%m-%d %H:%M:%S")
        return f"{float(value):g}"

    def _on_start_slider(self, _value):
        if self.slider_sync_in_progress or not self.current_time_values:
            return
        start_idx = int(float(self.start_idx_scale.get()))
        end_idx = int(float(self.end_idx_scale.get()))
        if start_idx > end_idx:
            self.slider_sync_in_progress = True
            self.end_idx_scale.set(start_idx)
            self.slider_sync_in_progress = False
            end_idx = start_idx
        self.start_time_var.set(self._format_time_value(self.current_time_values[start_idx]))
        self.end_time_var.set(self._format_time_value(self.current_time_values[end_idx]))
        self.plot_data()

    def _on_end_slider(self, _value):
        if self.slider_sync_in_progress or not self.current_time_values:
            return
        start_idx = int(float(self.start_idx_scale.get()))
        end_idx = int(float(self.end_idx_scale.get()))
        if end_idx < start_idx:
            self.slider_sync_in_progress = True
            self.start_idx_scale.set(end_idx)
            self.slider_sync_in_progress = False
            start_idx = end_idx
        self.start_time_var.set(self._format_time_value(self.current_time_values[start_idx]))
        self.end_time_var.set(self._format_time_value(self.current_time_values[end_idx]))
        self.plot_data()

    def _sync_sliders_to_time_entries(self):
        if not self.current_time_values:
            return

        start_val = self._parse_time_input(self.start_time_var.get())
        end_val = self._parse_time_input(self.end_time_var.get())
        if start_val is None or end_val is None:
            return

        if self.time_is_datetime:
            series_num = pd.to_datetime(pd.Series(self.current_time_values)).view("int64")
            start_num = pd.Timestamp(start_val).value
            end_num = pd.Timestamp(end_val).value
        else:
            series_num = pd.to_numeric(pd.Series(self.current_time_values), errors="coerce")
            start_num = float(start_val)
            end_num = float(end_val)

        start_idx = int((series_num - start_num).abs().idxmin())
        end_idx = int((series_num - end_num).abs().idxmin())
        if start_idx > end_idx:
            start_idx, end_idx = end_idx, start_idx

        self.slider_sync_in_progress = True
        self.start_idx_scale.set(start_idx)
        self.end_idx_scale.set(end_idx)
        self.slider_sync_in_progress = False

    def _ensure_canvas(self):
        if self.figure is None:
            self.figure = plt.Figure(figsize=(8, 5), dpi=100)
            self.canvas = FigureCanvasTkAgg(self.figure, master=self.plot_container)
            self.canvas.get_tk_widget().grid(row=0, column=0, sticky="nsew")
            self.toolbar = NavigationToolbar2Tk(self.canvas, self.plot_container, pack_toolbar=False)
            self.toolbar.update()
            self.toolbar.grid(row=1, column=0, sticky="ew")
            self.canvas.mpl_connect("scroll_event", self._on_scroll)

    def _on_scroll(self, event):
        if self.current_ax is None or event.inaxes != self.current_ax:
            return

        cur_min, cur_max = self.current_ax.get_xlim()
        center = event.xdata if event.xdata is not None else (cur_min + cur_max) / 2.0
        scale = 1 / 1.25 if event.button == "up" else 1.25

        new_min = center - (center - cur_min) * scale
        new_max = center + (cur_max - center) * scale
        if new_max <= new_min:
            return

        self.current_ax.set_xlim(new_min, new_max)
        self.canvas.draw_idle()

        if self.time_is_datetime:
            left = pd.Timestamp(mdates.num2date(new_min)).tz_localize(None)
            right = pd.Timestamp(mdates.num2date(new_max)).tz_localize(None)
            self.start_time_var.set(left.strftime("%Y-%m-%d %H:%M:%S"))
            self.end_time_var.set(right.strftime("%Y-%m-%d %H:%M:%S"))
        else:
            self.start_time_var.set(f"{new_min:g}")
            self.end_time_var.set(f"{new_max:g}")

        self._sync_sliders_to_time_entries()

    def plot_data(self):
        try:
            df, time_col, value_col = self._clean_plot_df()

            start_val = self._parse_time_input(self.start_time_var.get())
            end_val = self._parse_time_input(self.end_time_var.get())

            if start_val is not None:
                df = df[df[time_col] >= start_val]
            if end_val is not None:
                df = df[df[time_col] <= end_val]

            if df.empty:
                raise ValueError("No rows remain after applying time filter.")

            self._ensure_canvas()
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            ax.plot(df[time_col], df[value_col], linewidth=1.5, color="#0d6efd")
            self.current_ax = ax
            ax.set_title(f"{value_col} vs {time_col}")
            ax.set_xlabel(time_col)
            ax.set_ylabel(value_col)
            ax.grid(True, alpha=0.3)
            if self.time_is_datetime:
                ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d %H:%M:%S"))

            if not self.auto_y_var.get():
                y_min_txt = self.y_min_var.get().strip()
                y_max_txt = self.y_max_var.get().strip()
                y_min = float(y_min_txt) if y_min_txt else None
                y_max = float(y_max_txt) if y_max_txt else None

                if y_min is not None and y_max is not None and y_min >= y_max:
                    raise ValueError("Y min must be less than Y max.")

                current_min, current_max = ax.get_ylim()
                ax.set_ylim(
                    y_min if y_min is not None else current_min,
                    y_max if y_max is not None else current_max,
                )

            self.figure.autofmt_xdate()
            self.figure.tight_layout()
            self.canvas.draw()
            self._sync_sliders_to_time_entries()

            self.info_label.config(
                text=f"Plotted {len(df)} points from sheet '{self.sheet_var.get()}'.",
                foreground="#1f6f43",
            )
        except Exception as exc:
            messagebox.showerror("Plot Error", str(exc))
            self.info_label.config(text="Plot failed. Check selections and ranges.", foreground="#b42318")


def main():
    root = tk.Tk()
    app = TimeSeriesExcelViewer(root)
    root.mainloop()


if __name__ == "__main__":
    main()
