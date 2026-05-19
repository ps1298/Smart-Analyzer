import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

import matplotlib
matplotlib.use("TkAgg")

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class CorporateReportBuilder:

    def __init__(self, root):

        self.root = root
        self.root.title("Corporate Data Analyzer")
        self.root.geometry("1200x760")
        self.root.resizable(True, True)

        # =====================================================
        # THEME VARIABLES
        # =====================================================

        self.current_theme = "dark"

        # DARK THEME
        self.bg_main = "#1E1E1E"
        self.bg_panel = "#2B2B2B"
        self.bg_header = "#111827"

        self.text_color = "#F3F4F6"

        self.accent_blue = "#2563EB"
        self.accent_green = "#16A34A"
        self.accent_purple = "#9333EA"

        self.border_color = "#3F3F46"

        # LIGHT THEME
        self.light_bg_main = "#F3F4F6"
        self.light_bg_panel = "#FFFFFF"
        self.light_bg_header = "#E5E7EB"

        self.light_text = "#111827"

        self.root.configure(bg=self.bg_main)

        # =====================================================
        # VARIABLES
        # =====================================================

        self.file_path = ""
        self.df = None
        self.report_df = None
        self.canvas = None
        self.current_figure = None

        self.agg_map = {
            "Sum": "sum",
            "Mean": "mean",
            "Average": "mean",
            "Max": "max",
            "Min": "min",
            "Count": "count",
            "Median": "median",
        }

        # =====================================================
        # STYLE
        # =====================================================

        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.configure_treeview_dark()

        self._build_ui()

    # =====================================================
    # BUILD UI
    # =====================================================

    def _build_ui(self):

        # =====================================================
        # HEADER
        # =====================================================

        self.header = tk.Label(
            self.root,
            text="Corporate Data Analyzer",
            font=("Segoe UI", 20, "bold"),
            bg=self.bg_header,
            fg="white",
            pady=15
        )

        self.header.pack(fill="x")

        # =====================================================
        # THEME TOGGLE
        # =====================================================

        theme_frame = tk.Frame(
            self.root,
            bg=self.bg_main
        )

        theme_frame.pack(fill="x", padx=15, pady=5)

        self.theme_btn = tk.Button(
            theme_frame,
            text="☀ Light Mode",
            command=self.toggle_theme,
            bg=self.accent_blue,
            fg="white",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
            cursor="hand2"
        )

        self.theme_btn.pack(anchor="e")

        # =====================================================
        # FILE FRAME
        # =====================================================

        file_frame = tk.Frame(
            self.root,
            bg=self.bg_main
        )

        file_frame.pack(fill="x", padx=15, pady=10)

        tk.Label(
            file_frame,
            text="Select CSV/Excel:",
            bg=self.bg_main,
            fg=self.text_color,
            font=("Segoe UI", 10)
        ).pack(side="left")

        tk.Button(
            file_frame,
            text="Browse",
            command=self.browse_file,
            width=10,
            bg=self.accent_blue,
            fg="white",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
            cursor="hand2"
        ).pack(side="left", padx=8)

        tk.Button(
            file_frame,
            text="Read",
            command=self.read_file,
            width=10,
            bg=self.accent_green,
            fg="white",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
            cursor="hand2"
        ).pack(side="left")

        self.file_lbl = tk.Label(
            file_frame,
            text="No file selected",
            fg="#60A5FA",
            bg=self.bg_main,
            font=("Segoe UI", 10)
        )

        self.file_lbl.pack(side="left", padx=12)

        # =====================================================
        # FILE INFO
        # =====================================================

        self.info_frame = tk.LabelFrame(
            self.root,
            text="File Information",
            padx=10,
            pady=10,
            bg=self.bg_panel,
            fg=self.text_color,
            font=("Segoe UI", 10, "bold"),
            bd=2
        )

        self.info_frame.pack(fill="x", padx=15, pady=8)

        self.info_text = tk.Text(
            self.info_frame,
            height=5,
            bg="#111827",
            fg="white",
            insertbackground="white",
            relief="flat",
            font=("Consolas", 10)
        )

        self.info_text.pack(fill="x")

        # =====================================================
        # CONTROLS
        # =====================================================

        self.controls = tk.LabelFrame(
            self.root,
            text="Build Report",
            padx=10,
            pady=10,
            bg=self.bg_panel,
            fg=self.text_color,
            font=("Segoe UI", 10, "bold"),
            bd=2
        )

        self.controls.pack(fill="x", padx=15, pady=8)

        row1 = tk.Frame(self.controls, bg=self.bg_panel)
        row1.pack(fill="x", pady=5)

        tk.Label(
            row1,
            text="Group By:",
            bg=self.bg_panel,
            fg=self.text_color,
            font=("Segoe UI", 10)
        ).pack(side="left")

        self.group_col_var = tk.StringVar()

        self.group_col_cb = ttk.Combobox(
            row1,
            textvariable=self.group_col_var,
            state="disabled",
            width=25
        )

        self.group_col_cb.pack(side="left", padx=8)

        tk.Label(
            row1,
            text="Aggregation:",
            bg=self.bg_panel,
            fg=self.text_color,
            font=("Segoe UI", 10)
        ).pack(side="left", padx=(15, 0))

        self.agg_var = tk.StringVar()

        self.agg_cb = ttk.Combobox(
            row1,
            textvariable=self.agg_var,
            state="disabled",
            values=list(self.agg_map.keys()),
            width=18
        )

        self.agg_cb.pack(side="left", padx=8)

        tk.Label(
            row1,
            text="Value Column:",
            bg=self.bg_panel,
            fg=self.text_color,
            font=("Segoe UI", 10)
        ).pack(side="left", padx=(15, 0))

        self.value_col_var = tk.StringVar()

        self.value_col_cb = ttk.Combobox(
            row1,
            textvariable=self.value_col_var,
            state="disabled",
            width=25
        )

        self.value_col_cb.pack(side="left", padx=8)

        # =====================================================
        # BUTTON ROW
        # =====================================================

        row2 = tk.Frame(self.controls, bg=self.bg_panel)
        row2.pack(fill="x", pady=10)

        tk.Button(
            row2,
            text="Preview Report",
            command=self.preview_report,
            width=18,
            bg=self.accent_green,
            fg="white",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
            cursor="hand2"
        ).pack(side="left", padx=5)

        tk.Label(
            row2,
            text="Export As:",
            bg=self.bg_panel,
            fg=self.text_color,
            font=("Segoe UI", 10)
        ).pack(side="left", padx=(15, 0))

        self.export_format_var = tk.StringVar(value="Excel (.xlsx)")

        self.export_format_cb = ttk.Combobox(
            row2,
            textvariable=self.export_format_var,
            state="readonly",
            values=["Excel (.xlsx)", "CSV (.csv)"],
            width=18
        )

        self.export_format_cb.pack(side="left", padx=8)

        tk.Button(
            row2,
            text="Export Report",
            command=self.export_report,
            width=16,
            bg=self.accent_blue,
            fg="white",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
            cursor="hand2"
        ).pack(side="left", padx=5)

        # =====================================================
        # CHART BUILDER
        # =====================================================

        self.chart_frame = tk.LabelFrame(
            self.root,
            text="Chart Builder",
            padx=10,
            pady=10,
            bg=self.bg_panel,
            fg=self.text_color,
            font=("Segoe UI", 10, "bold"),
            bd=2
        )

        self.chart_frame.pack(fill="x", padx=15, pady=8)

        c_row = tk.Frame(self.chart_frame, bg=self.bg_panel)
        c_row.pack(fill="x")

        tk.Label(
            c_row,
            text="Chart Type:",
            bg=self.bg_panel,
            fg=self.text_color,
            font=("Segoe UI", 10)
        ).pack(side="left")

        self.chart_type_var = tk.StringVar(value="Bar")

        self.chart_type_cb = ttk.Combobox(
            c_row,
            textvariable=self.chart_type_var,
            state="readonly",
            values=["Bar", "Column", "Pie", "Line"],
            width=15
        )

        self.chart_type_cb.pack(side="left", padx=8)

        tk.Button(
            c_row,
            text="Preview Chart",
            command=self.preview_chart,
            width=15,
            bg=self.accent_blue,
            fg="white",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
            cursor="hand2"
        ).pack(side="left", padx=5)

        tk.Button(
            c_row,
            text="Export Chart",
            command=self.export_chart,
            width=15,
            bg=self.accent_purple,
            fg="white",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
            cursor="hand2"
        ).pack(side="left", padx=5)

        # =====================================================
        # MAIN AREA
        # =====================================================

        main = tk.Frame(
            self.root,
            bg=self.bg_main
        )

        main.pack(fill="both", expand=True, padx=15, pady=10)

        # =====================================================
        # TABLE BOX
        # =====================================================

        self.table_box = tk.LabelFrame(
            main,
            text="Report Preview",
            padx=8,
            pady=8,
            bg=self.bg_panel,
            fg=self.text_color,
            font=("Segoe UI", 10, "bold")
        )

        self.table_box.pack(side="left", fill="both", expand=True)

        self.tree = ttk.Treeview(
            self.table_box,
            columns=("A", "B"),
            show="headings"
        )

        self.tree.heading("A", text="Group")
        self.tree.heading("B", text="Value")

        self.tree.column("A", width=220, anchor="w")
        self.tree.column("B", width=140, anchor="e")

        yscroll = ttk.Scrollbar(
            self.table_box,
            orient="vertical",
            command=self.tree.yview
        )

        self.tree.configure(yscrollcommand=yscroll.set)

        self.tree.pack(side="left", fill="both", expand=True)

        yscroll.pack(side="right", fill="y")

        # =====================================================
        # CHART BOX
        # =====================================================

        self.chart_box = tk.LabelFrame(
            main,
            text="Chart Preview",
            padx=8,
            pady=8,
            bg=self.bg_panel,
            fg=self.text_color,
            font=("Segoe UI", 10, "bold")
        )

        self.chart_box.pack(side="right", fill="both", expand=True)

        self.chart_container = tk.Frame(
            self.chart_box,
            bg=self.bg_panel
        )

        self.chart_container.pack(fill="both", expand=True)

    # =====================================================
    # THEME SWITCH
    # =====================================================

    def toggle_theme(self):

        if self.current_theme == "dark":

            self.current_theme = "light"

            self.root.configure(bg=self.light_bg_main)

            self.header.config(
                bg=self.light_bg_header,
                fg="black"
            )

            self.theme_btn.config(
                text="🌙 Dark Mode"
            )

            self.configure_treeview_light()

            self.apply_theme(
                bg_main=self.light_bg_main,
                bg_panel=self.light_bg_panel,
                text=self.light_text
            )

        else:

            self.current_theme = "dark"

            self.root.configure(bg=self.bg_main)

            self.header.config(
                bg=self.bg_header,
                fg="white"
            )

            self.theme_btn.config(
                text="☀ Light Mode"
            )

            self.configure_treeview_dark()

            self.apply_theme(
                bg_main=self.bg_main,
                bg_panel=self.bg_panel,
                text=self.text_color
            )

    # =====================================================
    # APPLY THEME
    # =====================================================

    def apply_theme(self, bg_main, bg_panel, text):

        frames = [
            self.info_frame,
            self.controls,
            self.chart_frame,
            self.table_box,
            self.chart_box
        ]

        for frame in frames:
            frame.config(bg=bg_panel, fg=text)

        self.info_text.config(
            bg="white" if self.current_theme == "light" else "#111827",
            fg="black" if self.current_theme == "light" else "white",
            insertbackground="black" if self.current_theme == "light" else "white"
        )

    # =====================================================
    # TREEVIEW THEMES
    # =====================================================

    def configure_treeview_dark(self):

        self.style.configure(
            "Treeview",
            background="#2B2B2B",
            foreground="white",
            fieldbackground="#2B2B2B",
            rowheight=28
        )

        self.style.configure(
            "Treeview.Heading",
            background="#111827",
            foreground="white",
            font=("Segoe UI", 10, "bold")
        )

    def configure_treeview_light(self):

        self.style.configure(
            "Treeview",
            background="white",
            foreground="black",
            fieldbackground="white",
            rowheight=28
        )

        self.style.configure(
            "Treeview.Heading",
            background="#E5E7EB",
            foreground="black",
            font=("Segoe UI", 10, "bold")
        )

    # =====================================================
    # HELPERS
    # =====================================================

    def _set_info(self, text):

        self.info_text.delete("1.0", tk.END)
        self.info_text.insert(tk.END, text)

    def _input_folder(self):

        if not self.file_path:
            return ""

        return os.path.dirname(os.path.abspath(self.file_path))

    def _safe_numeric_convert(self, series):

        return pd.to_numeric(
            series.astype(str).str.replace(",", "", regex=False),
            errors="coerce"
        )

    # =====================================================
    # FILE BROWSE
    # =====================================================

    def browse_file(self):

        path = filedialog.askopenfilename(
            title="Select CSV or Excel File",
            filetypes=[
                ("CSV Files", "*.csv"),
                ("Excel Files", "*.xlsx *.xls")
            ]
        )

        if path:
            self.file_path = path
            self.file_lbl.config(text=os.path.basename(path))

    # =====================================================
    # READ FILE
    # =====================================================

    def read_file(self):

        if not self.file_path:
            messagebox.showerror("Error", "Please select a file first.")
            return

        try:

            if self.file_path.lower().endswith(".csv"):
                self.df = pd.read_csv(self.file_path)
            else:
                self.df = pd.read_excel(self.file_path)

            rows, cols = self.df.shape

            headings = list(self.df.columns)

            info = (
                f"Rows: {rows}\n"
                f"Columns: {cols}\n\n"
                f"Column Headings:\n- "
                + "\n- ".join(map(str, headings))
            )

            self._set_info(info)

            text_cols = list(
                self.df.select_dtypes(include=["object"]).columns
            )

            numeric_cols = list(
                self.df.select_dtypes(include=["number"]).columns
            )

            for c in self.df.columns:

                if c in numeric_cols:
                    continue

                converted = self._safe_numeric_convert(self.df[c])

                if converted.notna().sum() >= max(3, int(0.6 * len(self.df))):
                    numeric_cols.append(c)

            self.group_col_cb.config(
                state="readonly",
                values=text_cols
            )

            self.value_col_cb.config(
                state="readonly",
                values=numeric_cols
            )

            self.agg_cb.config(state="readonly")

            if text_cols:
                self.group_col_var.set(text_cols[0])

            if numeric_cols:
                self.value_col_var.set(numeric_cols[0])

            self.agg_var.set("Sum")

            messagebox.showinfo(
                "Success",
                "File loaded successfully."
            )

        except Exception as e:

            messagebox.showerror(
                "Read Error",
                f"Could not read file.\n\n{e}"
            )

    # =====================================================
    # PREVIEW REPORT
    # =====================================================

    def preview_report(self):

        if self.df is None:
            messagebox.showerror("Error", "Please read the file first.")
            return

        group_col = self.group_col_var.get().strip()
        agg_label = self.agg_var.get().strip()
        value_col = self.value_col_var.get().strip()

        if not group_col or not agg_label or not value_col:
            messagebox.showerror(
                "Error",
                "Please select all report options."
            )
            return

        try:

            agg_func = self.agg_map.get(agg_label, "sum")

            df_work = self.df.copy()

            df_work[group_col] = (
                df_work[group_col]
                .astype(str)
                .str.strip()
                .str.title()
            )

            df_work[value_col] = self._safe_numeric_convert(
                df_work[value_col]
            )

            df_work = df_work.dropna(subset=[value_col])

            report = (
                df_work
                .groupby(group_col)[value_col]
                .agg(agg_func)
                .reset_index()
                .rename(columns={
                    group_col: "Group",
                    value_col: "Value"
                })
                .sort_values("Value", ascending=False)
            )

            self.report_df = report

            self._render_table(report)

        except Exception as e:

            messagebox.showerror(
                "Report Error",
                f"Could not build report.\n\n{e}"
            )

    # =====================================================
    # RENDER TABLE
    # =====================================================

    def _render_table(self, report):

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.tree.heading("A", text=report.columns[0])
        self.tree.heading("B", text=report.columns[1])

        for _, r in report.iterrows():

            a = str(r.iloc[0])

            try:
                b = f"{float(r.iloc[1]):,.2f}"
            except:
                b = str(r.iloc[1])

            self.tree.insert("", "end", values=(a, b))

    # =====================================================
    # EXPORT REPORT
    # =====================================================

    def export_report(self):

        if self.report_df is None or self.report_df.empty:
            messagebox.showerror("Error", "No report available.")
            return

        folder = self._input_folder()

        fmt = self.export_format_var.get()

        try:

            if fmt.startswith("Excel"):

                out_path = os.path.join(
                    folder,
                    "report_output.xlsx"
                )

                self.report_df.to_excel(
                    out_path,
                    index=False
                )

            else:

                out_path = os.path.join(
                    folder,
                    "report_output.csv"
                )

                self.report_df.to_csv(
                    out_path,
                    index=False
                )

            messagebox.showinfo(
                "Exported",
                f"Report exported successfully:\n{out_path}"
            )

        except Exception as e:

            messagebox.showerror(
                "Export Error",
                f"Could not export report.\n\n{e}"
            )

    # =====================================================
    # PREVIEW CHART
    # =====================================================

    def preview_chart(self):

        if self.report_df is None or self.report_df.empty:
            messagebox.showerror(
                "Error",
                "Please preview report first."
            )
            return

        chart_type = self.chart_type_var.get()

        data = self.report_df.head(10)

        try:

            for widget in self.chart_container.winfo_children():
                widget.destroy()

            fig = Figure(figsize=(5.5, 4.5), dpi=100)

            if self.current_theme == "dark":
                fig.patch.set_facecolor("#1E1E1E")
            else:
                fig.patch.set_facecolor("#FFFFFF")

            ax = fig.add_subplot(111)

            if self.current_theme == "dark":
                ax.set_facecolor("#2B2B2B")
                text_color = "white"
            else:
                ax.set_facecolor("#FFFFFF")
                text_color = "black"

            labels = data.iloc[:, 0].astype(str).tolist()
            values = data.iloc[:, 1].astype(float).tolist()

            if chart_type == "Bar":

                ax.barh(labels[::-1], values[::-1])

            elif chart_type == "Column":

                ax.bar(labels, values)
                ax.tick_params(axis="x", rotation=45)

            elif chart_type == "Line":

                ax.plot(labels, values, marker="o")
                ax.tick_params(axis="x", rotation=45)

            elif chart_type == "Pie":

                pie_n = 6 if len(values) > 6 else len(values)

                ax.pie(
                    values[:pie_n],
                    labels=labels[:pie_n],
                    autopct="%1.1f%%"
                )

                ax.axis("equal")

            ax.set_title(
                f"{chart_type} Chart",
                color=text_color
            )

            ax.tick_params(colors=text_color)

            ax.xaxis.label.set_color(text_color)
            ax.yaxis.label.set_color(text_color)

            for spine in ax.spines.values():
                spine.set_color(text_color)

            fig.tight_layout()

            self.current_figure = fig

            self.canvas = FigureCanvasTkAgg(
                fig,
                master=self.chart_container
            )

            self.canvas.draw()

            self.canvas.get_tk_widget().pack(
                fill="both",
                expand=True
            )

        except Exception as e:

            messagebox.showerror(
                "Chart Error",
                f"Could not generate chart.\n\n{e}"
            )

    # =====================================================
    # EXPORT CHART
    # =====================================================

    def export_chart(self):

        if self.current_figure is None:
            messagebox.showerror("Error", "No chart available.")
            return

        folder = self._input_folder()

        out_path = os.path.join(
            folder,
            "chart_output.png"
        )

        try:

            self.current_figure.savefig(
                out_path,
                dpi=200
            )

            messagebox.showinfo(
                "Exported",
                f"Chart exported successfully:\n{out_path}"
            )

        except Exception as e:

            messagebox.showerror(
                "Export Error",
                f"Could not export chart.\n\n{e}"
            )


# =====================================================
# MAIN
# =====================================================

if __name__ == "__main__":

    root = tk.Tk()

    app = CorporateReportBuilder(root)

    root.mainloop()