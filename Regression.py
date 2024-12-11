import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import statsmodels.api as sm
from PIL import ImageGrab
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class RegressionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Regression Analysis Tool")
        self.root.geometry("1315x938")

        # Create a Canvas widget for scrolling
        self.canvas = tk.Canvas(root)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Create a scrollbar and attach it to the canvas
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.config(yscrollcommand=self.scrollbar.set)
        
        # Create a frame inside the canvas which will contain all the widgets
        self.main_frame = tk.Frame(self.canvas, bg="white")
        self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        
        self.main_frame.bind("<Configure>", self.on_frame_configure)

        # Buttons and input fields
        self.load_button = tk.Button(self.main_frame, text="Load File (Excel/CSV)", command=self.load_file)
        self.load_button.grid(row=0, column=0, padx=10, pady=10)

        self.load_button.config(bg="lightblue")
        self.load_button.config(relief="groove")
        self.load_button.config(font=("Verdana", 9, "italic bold"))
 
        # Dataset display
        self.tree_frame = tk.Frame(self.main_frame)
        self.tree_frame.grid(row=1, column=0, columnspan=4, padx=10, pady=10)
        
        self.tree = ttk.Treeview(self.tree_frame, show='headings')
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.scrollbar_tree = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.scrollbar_tree.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.config(yscrollcommand=self.scrollbar_tree.set)

        # Input fields for regression
        tk.Label(self.main_frame, text="Input Y Range (column number):").grid(row=2, column=0, sticky=tk.W)
        self.y_entry = tk.Entry(self.main_frame)
        self.y_entry.grid(row=2, column=1, pady=5)
        
        tk.Label(self.main_frame, text="Input X Range (comma-separated column numbers):").grid(row=3, column=0, sticky=tk.W)
        self.x_entry = tk.Entry(self.main_frame)
        self.x_entry.grid(row=3, column=1, pady=5)
        
        tk.Label(self.main_frame, text="Confidence Level (default 95%):").grid(row=4, column=0, sticky=tk.W)
        self.confidence_entry = tk.Entry(self.main_frame)
        self.confidence_entry.insert(0, "95")
        self.confidence_entry.grid(row=4, column=1, pady=5)
        
        self.run_button = tk.Button(self.main_frame, text="Perform Regression", command=self.perform_regression)
        self.run_button.grid(row=5, column=0, columnspan=2, pady=10)

        # Results Frame with Table for Coefficients
        self.results_frame = tk.Frame(self.main_frame)
        self.results_frame.grid(row=6, column=0, columnspan=4, padx=10, pady=10)
        
        self.results_table = ttk.Treeview(self.results_frame, columns=("Variable", "Coefficient", "Std Error", "t Stat", "P-value", "Lower 95%", "Upper 95%"), show='headings')
        self.results_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.results_scrollbar = ttk.Scrollbar(self.results_frame, orient=tk.VERTICAL, command=self.results_table.yview)
        self.results_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.results_table.config(yscrollcommand=self.results_scrollbar.set)
        
        # Set column headings for the result table
        self.results_table.heading("Variable", text="Variable")
        self.results_table.heading("Coefficient", text="Coefficient")
        self.results_table.heading("Std Error", text="Std Error")
        self.results_table.heading("t Stat", text="t Stat")
        self.results_table.heading("P-value", text="P-value")
        self.results_table.heading("Lower 95%", text="Lower 95%")
        self.results_table.heading("Upper 95%", text="Upper 95%")
        
        # Set column widths and row height
        self.results_table.column("Variable", width=250, anchor="center")
        self.results_table.column("Coefficient", width=250, anchor="center")
        self.results_table.column("Std Error", width=150, anchor="center")
        self.results_table.column("t Stat", width=150, anchor="center")
        self.results_table.column("P-value", width=150, anchor="center")
        self.results_table.column("Lower 95%", width=150, anchor="center")
        self.results_table.column("Upper 95%", width=150, anchor="center")
        
        # Text widget for displaying regression equation and recommendations
        self.results_text_frame = tk.Frame(self.main_frame)
        self.results_text_frame.grid(row=7, column=0, columnspan=4, padx=10, pady=10)
        
        self.results_text = tk.Text(self.results_text_frame, wrap=tk.WORD, width=150, height=10)
        self.results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.results_text_scrollbar = ttk.Scrollbar(self.results_text_frame, orient=tk.VERTICAL, command=self.results_text.yview)
        self.results_text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.results_text.config(yscrollcommand=self.results_text_scrollbar.set)
        
        self.results_text.config(state=tk.DISABLED)  # Make text box read-only
        
        # Button for dataset screenshot
        self.screenshot_button = tk.Button(self.main_frame, text="Capture Dataset Screenshot", command=self.capture_screenshot)
        self.screenshot_button.grid(row=8, column=0, pady=10)

        # Canvas for Scatter Plot
        self.plot_frame = tk.Frame(self.main_frame)
        self.plot_frame.grid(row=9, column=0, columnspan=4, padx=10, pady=10)

        self.canvas_plot = None  # Placeholder for the plot

    def on_frame_configure(self, event=None):
        """Update the scrollable region to the height of the main frame."""
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", ".csv"), ("Excel files", ".xlsx")])
        if not file_path:
            return
        
        try:
            # Load dataset
            if file_path.endswith(".csv"):
                self.data = pd.read_csv(file_path)
            elif file_path.endswith(".xlsx"):
                self.data = pd.read_excel(file_path)
            
            # Replace column names with numbers (1, 2, 3, ...) and show original names in the format
            self.data.columns = [f"{i + 1} ({col})" for i, col in enumerate(self.data.columns)]
            
            self.populate_treeview(self.data)
        except Exception as e:
            messagebox.showerror("Error", f"Could not load file: {e}")

    def populate_treeview(self, data):
        # Clear existing data in the Treeview
        self.tree.delete(*self.tree.get_children())
        
        # Add a Row Number column at the beginning
        self.tree["column"] = list(data.columns)
        self.tree["show"] = "headings"
        
        # Set the column headers
        for col in data.columns:
            self.tree.heading(col, text=col)
        
        # Insert rows
        for _, row in data.iterrows():
            self.tree.insert("", tk.END, values=list(row))
        
        # Resize columns to fit content
        for col in data.columns:
            self.tree.column(col, width=150, anchor="center")

    def perform_regression(self):
        try:
            y_col = int(self.y_entry.get().strip()) - 1
            x_cols = [int(col.strip()) - 1 for col in self.x_entry.get().split(",")]
            
            y = pd.to_numeric(self.data.iloc[:, y_col], errors='coerce')
            X = self.data.iloc[:, x_cols].apply(pd.to_numeric, errors='coerce')

            # Drop rows with missing values
            data = pd.concat([y, X], axis=1).dropna()
            y = data.iloc[:, 0]
            X = data.iloc[:, 1:]

            X = sm.add_constant(X)  # Add intercept
            model = sm.OLS(y, X).fit()
            
            self.display_results(model)
            self.plot_scatter(X, y)  # Update scatter plot after regression
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during regression: {e}")

    def display_results(self, model):
        # Clear any previous results in the result table
        for row in self.results_table.get_children():
            self.results_table.delete(row)
        
        # Extract regression coefficients summary
        summary_df = model.summary2().tables[1]
        
        # Insert regression coefficients into the table
        for index, row in summary_df.iterrows():
            self.results_table.insert("", "end", values=( 
                index,
                f"{row['Coef.']:.4f}",
                f"{row['Std.Err.']:.4f}",
                f"{row['t']:.4f}",
                f"{row['P>|t|']:.4f}",
                f"{row['[0.025']:.4f}",
                f"{row['0.975]']:.4f}"))
        
        # Display regression equation in the Text widget
        eq = "Regression Equation:\n\nY = "
        params = model.params
        eq += f"{params['const']:.4f} "
        for var in params.index[1:]:
            eq += f"+ {params[var]:.4f} * {var} "
        
        # Enable text widget to display the equation
        self.results_text.config(state=tk.NORMAL)
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, eq + "\n\n")
        
        # Display Variables to Purge in the Text widget
        recommendation = "\nVariables that need to Purge:\n\n"
        p_values = model.pvalues
        for var, p_value in p_values.items():
            if var != "const" and p_value > 0.05:
                recommendation += f"Variable {var} has a high p-value ({p_value:.4f}). Consider purging it.\n"
        if recommendation.strip() == "Variables that need to Purge:":
            recommendation += "No variables have high p-values for purging.\n"
        
        self.results_text.insert(tk.END, recommendation)
        self.results_text.config(state=tk.DISABLED)

    def plot_scatter(self, X, y):
        # Clear the previous plot if it exists
        if self.canvas_plot:
            self.canvas_plot.get_tk_widget().destroy()

        # Create a scatter plot
        fig, ax = plt.subplots(figsize=(6, 4))

        ax.scatter(X.iloc[:, 0], y, color="lightblue")
        ax.set_xlabel(X.columns[0])
        ax.set_ylabel("Y (Dependent Variable)")
        ax.set_title("Scatter Plot of X vs Y")

        # Embed the plot into the Tkinter window
        self.canvas_plot = FigureCanvasTkAgg(fig, master=self.plot_frame)
        self.canvas_plot.draw()
        self.canvas_plot.get_tk_widget().pack()

    def capture_screenshot(self):
        try:
            x = self.root.winfo_rootx()
            y = self.root.winfo_rooty()
            x1 = x + self.root.winfo_width()
            y1 = y + self.root.winfo_height()
            
            screenshot = ImageGrab.grab(bbox=(x, y, x1, y1))
            save_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
            if save_path:
                screenshot.save(save_path)
                messagebox.showinfo("Screenshot Captured", f"Screenshot saved to {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not capture screenshot: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RegressionApp(root)
    root.mainloop()
