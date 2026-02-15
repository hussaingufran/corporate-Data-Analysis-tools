# ==========================================================
# Corporate Data Analysis Desktop Software (AI Powered)
# Built with Tkinter + Pandas + Matplotlib
# ==========================================================

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class DataAnalyzerApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Corporate Data Analysis Tool")
        self.root.geometry("1200x750")
        self.root.configure(bg="#f4f6f9")

        # Data Variables
        self.file_path = None
        self.df = None
        self.report_df = None
        self.chart_canvas = None

        self.create_widgets()

    # ======================================================
    # UI Layout
    # ======================================================
    def create_widgets(self):

        title = tk.Label(self.root, text="Corporate Data Analysis Software",
                         font=("Arial", 20, "bold"), bg="#f4f6f9")
        title.pack(pady=10)

        # File Frame
        file_frame = tk.LabelFrame(self.root, text="File Selection",
                                   font=("Arial", 12, "bold"), padx=10, pady=10)
        file_frame.pack(fill="x", padx=20, pady=5)

        tk.Button(file_frame, text="Browse File", command=self.browse_file, width=15).grid(row=0, column=0, padx=5)
        tk.Button(file_frame, text="Read File", command=self.read_file, width=15).grid(row=0, column=1, padx=5)

        self.file_label = tk.Label(file_frame, text="No file selected", fg="red")
        self.file_label.grid(row=0, column=2, padx=10)

        self.summary_label = tk.Label(file_frame, text="")
        self.summary_label.grid(row=1, column=0, columnspan=3, pady=5)

        # Report Builder Frame
        report_frame = tk.LabelFrame(self.root, text="Report Builder",
                                     font=("Arial", 12, "bold"), padx=10, pady=10)
        report_frame.pack(fill="x", padx=20, pady=5)

        tk.Label(report_frame, text="Group By:").grid(row=0, column=0)
        self.group_dropdown = ttk.Combobox(report_frame, state="readonly", width=20)
        self.group_dropdown.grid(row=0, column=1, padx=5)

        tk.Label(report_frame, text="Aggregation:").grid(row=0, column=2)
        self.agg_dropdown = ttk.Combobox(report_frame, state="readonly",
                                         values=["sum", "mean", "max", "min", "count", "median"],
                                         width=15)
        self.agg_dropdown.grid(row=0, column=3, padx=5)

        tk.Label(report_frame, text="Value Column:").grid(row=0, column=4)
        self.value_dropdown = ttk.Combobox(report_frame, state="readonly", width=20)
        self.value_dropdown.grid(row=0, column=5, padx=5)

        tk.Button(report_frame, text="Preview Report",
                  command=self.generate_report, width=15).grid(row=0, column=6, padx=10)

        tk.Button(report_frame, text="Export Report",
                  command=self.export_report, width=15).grid(row=0, column=7)

        # Chart Builder Frame
        chart_frame = tk.LabelFrame(self.root, text="Chart Builder",
                                    font=("Arial", 12, "bold"), padx=10, pady=10)
        chart_frame.pack(fill="x", padx=20, pady=5)

        tk.Label(chart_frame, text="Chart Type:").grid(row=0, column=0)

        self.chart_dropdown = ttk.Combobox(chart_frame, state="readonly",
                                           values=["Bar", "Column", "Line", "Pie"],
                                           width=15)
        self.chart_dropdown.grid(row=0, column=1, padx=5)

        tk.Button(chart_frame, text="Preview Chart",
                  command=self.generate_chart, width=15).grid(row=0, column=2, padx=10)

        tk.Button(chart_frame, text="Export Chart",
                  command=self.export_chart, width=15).grid(row=0, column=3)

        # Output Area
        self.output_frame = tk.Frame(self.root)
        self.output_frame.pack(fill="both", expand=True, padx=20, pady=10)

        self.tree = ttk.Treeview(self.output_frame)
        self.tree.pack(fill="both", expand=True)

    # ======================================================
    # File Handling
    # ======================================================
    def browse_file(self):
        self.file_path = filedialog.askopenfilename(
            filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx")]
        )
        if self.file_path:
            self.file_label.config(text=self.file_path, fg="green")

    def read_file(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select a file first.")
            return

        try:
            if self.file_path.endswith(".csv"):
                self.df = pd.read_csv(self.file_path)
            else:
                self.df = pd.read_excel(self.file_path)

            rows, cols = self.df.shape
            self.summary_label.config(
                text=f"Rows: {rows} | Columns: {cols} | Headings: {list(self.df.columns)}"
            )

            self.detect_columns()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file:\n{e}")

    # ======================================================
    # Column Detection
    # ======================================================
    def detect_columns(self):
        text_cols = self.df.select_dtypes(include=["object"]).columns.tolist()
        numeric_cols = self.df.select_dtypes(include=["int64", "float64"]).columns.tolist()

        self.group_dropdown["values"] = text_cols
        self.value_dropdown["values"] = numeric_cols

    # ======================================================
    # Report Generation
    # ======================================================
    def generate_report(self):
        if self.df is None:
            messagebox.showerror("Error", "Please read the file first.")
            return

        group_col = self.group_dropdown.get()
        agg_method = self.agg_dropdown.get()
        value_col = self.value_dropdown.get()

        if not group_col or not agg_method or not value_col:
            messagebox.showerror("Error", "Please select all dropdown values.")
            return

        try:
            self.report_df = (
                self.df.groupby(group_col)[value_col]
                .agg(agg_method)
                .reset_index()
                .sort_values(by=value_col, ascending=False)
            )

            self.display_table(self.report_df)

        except Exception as e:
            messagebox.showerror("Error", f"Report generation failed:\n{e}")

    # ======================================================
    # Display Table
    # ======================================================
    def display_table(self, dataframe):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(dataframe.columns)
        self.tree["show"] = "headings"

        for col in dataframe.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)

        for _, row in dataframe.iterrows():
            self.tree.insert("", "end", values=list(row))

    # ======================================================
    # Export Report
    # ======================================================
    def export_report(self):
        if self.report_df is None:
            messagebox.showerror("Error", "No report to export.")
            return

        folder = os.path.dirname(self.file_path)
        save_path = filedialog.asksaveasfilename(
            initialdir=folder,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")]
        )

        if save_path:
            if save_path.endswith(".csv"):
                self.report_df.to_csv(save_path, index=False)
            else:
                self.report_df.to_excel(save_path, index=False)

            messagebox.showinfo("Success", "Report exported successfully!")

    # ======================================================
    # Chart Generation
    # ======================================================
    def generate_chart(self):
        if self.report_df is None:
            messagebox.showerror("Error", "Generate report first.")
            return

        chart_type = self.chart_dropdown.get()

        if not chart_type:
            messagebox.showerror("Error", "Select chart type.")
            return

        if self.chart_canvas:
            self.chart_canvas.get_tk_widget().destroy()

        fig, ax = plt.subplots(figsize=(6, 4))

        x = self.report_df.iloc[:, 0]
        y = self.report_df.iloc[:, 1]

        if chart_type == "Bar":
            ax.bar(x, y)
        elif chart_type == "Column":
            ax.barh(x, y)
        elif chart_type == "Line":
            ax.plot(x, y)
        elif chart_type == "Pie":
            ax.pie(y, labels=x, autopct="%1.1f%%")

        ax.set_title("Report Chart")
        fig.tight_layout()

        self.chart_canvas = FigureCanvasTkAgg(fig, master=self.output_frame)
        self.chart_canvas.draw()
        self.chart_canvas.get_tk_widget().pack()

    # ======================================================
    # Export Chart
    # ======================================================
    def export_chart(self):
        if self.chart_canvas is None:
            messagebox.showerror("Error", "No chart to export.")
            return

        folder = os.path.dirname(self.file_path)
        save_path = filedialog.asksaveasfilename(
            initialdir=folder,
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")]
        )

        if save_path:
            self.chart_canvas.figure.savefig(save_path)
            messagebox.showinfo("Success", "Chart exported successfully!")


# ==========================================================
# Run Application
# ==========================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = DataAnalyzerApp(root)
    root.mainloop()
