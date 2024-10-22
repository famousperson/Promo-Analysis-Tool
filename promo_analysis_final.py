import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
from tkinter import font as tkfont
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
import os

class PromoAnalysisTool:
    def __init__(self, master):
        self.master = master
        self.master.title("M.J. BALE Promo Analysis Tool")
        self.master.geometry("1000x1100")
        self.master.configure(bg="#f0f0f0")  # Light gray background

        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("TButton", padding=10, font=('Helvetica', 12))
        self.style.configure("TCheckbutton", padding=5, font=('Helvetica', 16))  # Increased font size
        self.style.configure("TLabel", padding=5, font=('Helvetica', 12))

        # Custom style for modern toggle
        self.style.configure("Toggle.TCheckbutton",
                             indicatorsize=30,
                             padding=10,
                             relief="flat",
                             background="#f0f0f0",
                             foreground="black",
                             font=('Helvetica', 16))
        self.style.map("Toggle.TCheckbutton",
                      background=[('active', '#e0e0e0')],
                      indicatorcolor=[("selected", "#4CAF50"), ("!selected", "#ccc")])

        # Increase tab font size
        self.style.configure("TNotebook.Tab", font=('Helvetica', 14, 'bold'))

        self.promo_functions = {
            "$399 & $599 Suits": self.analyze_sublime_suits,
            "25% Off Chinos": self.analyze_chino_25_percent_off,
            "25% Off Coats/Outerwear": self.analyze_25_percent_off_coats,
            "25% Off Selected Styles": self.analyze_25_percent_off,
            "25% Off Tailoring": self.analyze_25_percent_off_winter_tailoring,
            "40% Off Tailoring": self.analyze_40_percent_off_tailoring,
            "50% Off 50 Styles": self.analyze_50_50,
            "Casual Bottom Multibuy": self.analyze_casual_bottom_multibuy,
            "Chino Multibuy": self.analyze_chino_multibuy,
            "FP Purchase": self.analyze_fp_purchase,
            "Gift Card": self.analyze_gift_card,
            "Knits Offer": self.analyze_25_percent_off_knits,
            "Linen Shirts Multibuy": self.analyze_linen_shirts_multibuy,
            "MD Purchase": self.analyze_md_purchase,
            "Polo Multibuy": self.analyze_polo_multibuy,
            "Promo Code": self.analyze_promo_code,
            "Shirts Multibuy": self.analyze_shirts_multibuy,
            "Suit Multibuy": self.analyze_suit_multibuy,
            "TAF25": self.analyze_taf25,
            "Tee Multibuy": self.analyze_tee_multibuy
        }

        self.create_widgets()

    def create_widgets(self):
        # Create a notebook (tabbed interface)
        self.notebook = ttk.Notebook(self.master)
        self.notebook.pack(expand=True, fill="both", padx=20, pady=20)

        # Create tabs
        self.promo_selection_tab = ttk.Frame(self.notebook)
        self.analysis_tab = ttk.Frame(self.notebook)
        self.results_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.promo_selection_tab, text="Promo Selection")
        self.notebook.add(self.analysis_tab, text="Analysis")
        self.notebook.add(self.results_tab, text="Results")

        # Promo Selection Tab
        self.create_promo_selection_widgets()

        # Analysis Tab
        self.create_analysis_widgets()

        # Results Tab
        self.create_results_widgets()

    def create_promo_selection_widgets(self):
        frame = ttk.Frame(self.promo_selection_tab, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(frame, text="Select Active Promotions:", font=("Helvetica", 18, "bold")).grid(column=0, row=0, sticky=tk.W, pady=(0, 20))

        self.promo_vars = {}
        default_checked = ["Chino Multibuy", "FP Purchase", "Gift Card", "Linen Shirts Multibuy",
                           "MD Purchase", "Polo Multibuy", "Promo Code", "Shirts Multibuy",
                           "Suit Multibuy", "Tee Multibuy"]

        for i, promo in enumerate(sorted(self.promo_functions.keys())):  # Sort promos alphabetically
            var = tk.BooleanVar(value=promo in default_checked)
            cb = ttk.Checkbutton(frame, text=promo, variable=var, style="Toggle.TCheckbutton")
            cb.grid(column=i % 3, row=i // 3 + 1, sticky=tk.W, padx=(0, 20), pady=(0, 10))
            self.promo_vars[promo] = var

    def create_analysis_widgets(self):
        frame = ttk.Frame(self.analysis_tab, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Create a modern-looking upload button
        upload_button = tk.Button(frame, text="Upload Excel File", command=self.analyze_file,
                                  font=('Helvetica', 14), bg="#4CAF50", fg="white",
                                  activebackground="#45a049", activeforeground="white",
                                  relief=tk.FLAT, padx=20, pady=10)
        upload_button.grid(column=0, row=0, pady=20)

        self.file_label = ttk.Label(frame, text="No file selected", style="TLabel")
        self.file_label.grid(column=0, row=1, pady=10)

        self.progress = ttk.Progressbar(frame, orient=tk.HORIZONTAL, length=400, mode='determinate', style="TProgressbar")
        self.progress.grid(column=0, row=2, pady=20)

        # Configure progress bar style
        self.style.configure("TProgressbar", thickness=25, troughcolor='#f0f0f0',
                             background='#4CAF50', bordercolor='#f0f0f0')

    def create_results_widgets(self):
        frame = ttk.Frame(self.results_tab, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Create a more visually appealing text widget
        self.result_text = tk.Text(frame, height=10, width=80, font=("Helvetica", 12),
                                   bg="#ffffff", fg="#333333", relief=tk.FLAT,
                                   padx=10, pady=10)
        self.result_text.grid(column=0, row=0, pady=20)

        # Add a scrollbar to the text widget
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.result_text.yview)
        scrollbar.grid(column=1, row=0, sticky='ns')
        self.result_text.configure(yscrollcommand=scrollbar.set)

        self.fig, self.ax = plt.subplots(figsize=(8, 6), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.fig, master=frame)
        self.canvas.get_tk_widget().grid(column=0, row=1, pady=20, columnspan=2)

        # Create a modern-looking save button
        self.save_button = tk.Button(frame, text="Save Results", command=self.save_results,
                                     font=('Helvetica', 14), bg="#008CBA", fg="white",
                                     activebackground="#007B9A", activeforeground="white",
                                     relief=tk.FLAT, padx=20, pady=10)
        self.save_button.grid(column=0, row=2, pady=20, columnspan=2)

        # Make the frame expandable
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

    def analyze_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        self.file_label.config(text=f"Selected file: {file_path}")

        df = pd.read_excel(file_path)

        # Add new columns
        df.insert(0, "Promo Type", "")
        df.insert(1, "Tier Group", "")

        # Update progress bar
        total_steps = len(self.promo_vars) + 4  # +4 for additional steps
        step = 0

        # Apply selected promo analyses
        for promo, var in self.promo_vars.items():
            if var.get():
                df = self.promo_functions[promo](df)
                step += 1
                self.progress['value'] = (step / total_steps) * 100
                self.master.update_idletasks()

        # Additional analyses
        df = self.drop_shipping_lines(df)
        step += 1
        self.progress['value'] = (step / total_steps) * 100
        self.master.update_idletasks()

        df = self.analyze_fp_purchase(df)
        step += 1
        self.progress['value'] = (step / total_steps) * 100
        self.master.update_idletasks()

        df = self.drop_discount_lines(df)
        step += 1
        self.progress['value'] = (step / total_steps) * 100
        self.master.update_idletasks()

        df['Tier Group'] = df['Customer: Tags'].apply(lambda tags: self.get_tier_group(tags))
        step += 1
        self.progress['value'] = 100  # Ensure progress bar reaches 100%
        self.master.update_idletasks()

        # Fill in blank 'Line: Product Type' based on 'Line: Title'
        df.loc[(df['Line: Product Type'].isna()) & (df['Line: Title'].str.contains('Trouser', case=False, na=False)), 'Line: Product Type'] = 'Trousers'
        df.loc[(df['Line: Product Type'].isna()) & (df['Line: Title'].str.contains('Waistcoat', case=False, na=False)), 'Line: Product Type'] = 'Waistcoat'

        self.df = df  # Store the DataFrame for later use

        messagebox.showinfo("Analysis Result", "Analysis completed successfully!")

        self.second_check()
        self.create_pivot_chart()
        self.notebook.select(self.results_tab)  # Switch to results tab

        # Prompt to save the Excel file
        if messagebox.askyesno("Save Excel File", "Do you want to save the analysed data as an Excel file?"):
            self.save_excel_file()

    def second_check(self):
        df = self.df

        # Define the suit multibuy prices
        suit_multibuy_prices = [175, 200, 275, 350, 400, 425, 575, 700]

        # Apply Suit Multibuy tag
        mask = (df['Line: Title'].str.contains('Jacket|Trouser', case=False, na=False)) & \
               (df['Line: Total'].isin(suit_multibuy_prices))
        df.loc[mask, 'Promo Type'] = 'Suit Multibuy'

        self.df = df

    def save_excel_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                self.df.to_excel(writer, sheet_name='Sheet1', index=False)

                # Create pivot table
                pivot_table = pd.pivot_table(self.df, values=['Line: Quantity', 'Line: Total'],
                                             index=['Promo Type'], aggfunc='sum')
                pivot_table = pivot_table.sort_values('Line: Total', ascending=False)
                pivot_table = pivot_table[pivot_table.index != '']  # Remove blank rows

                # Reorder columns
                pivot_table = pivot_table[['Line: Total', 'Line: Quantity']]

                # Add total row
                total_row = pd.DataFrame({
                    'Line: Total': pivot_table['Line: Total'].sum(),
                    'Line: Quantity': pivot_table['Line: Quantity'].sum()
                }, index=['Total'])
                pivot_table = pd.concat([pivot_table, total_row])

                pivot_table.to_excel(writer, sheet_name='Promo Analysis', startrow=0, startcol=0)

                workbook = writer.book
                worksheet = workbook['Promo Analysis']

                # Formatting
                for col in ['A', 'B', 'C']:
                    worksheet.column_dimensions[col].width = 20

                for row in worksheet['A1:C1']:
                    for cell in row:
                        cell.font = Font(bold=True)

                # Rename columns
                worksheet['A1'] = 'Promo Type'
                worksheet['B1'] = 'Sales $'
                worksheet['C1'] = 'Quantity'

                # Format Sales $ as currency
                for cell in worksheet['B']:
                    cell.number_format = '$#,##0'

                # Create a more visually appealing and modern pie chart
                labels = pivot_table.index.to_numpy()[:-1]  # Exclude 'Total' from labels
                data = pivot_table['Line: Total'].to_numpy()[:-1]  # Exclude 'Total' from data

                fig, ax = plt.subplots(figsize=(8, 6))  # Adjust figure size if needed
                ax.pie(data, labels=labels, autopct='%1.0f%%', pctdistance=0.85)
                ax.set_title("Promo Analysis by Sales $")

                # Customize the chart appearance
                ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')  # Move legend outside the chart
                ax.axis('equal')

                # Save the chart as an image within the workbook
                chart_path = os.path.join(os.path.dirname(file_path), "chart.png")  # save chart in the same directory as the Excel file
                plt.savefig(chart_path, bbox_inches='tight')

                img = openpyxl.drawing.image.Image(chart_path)
                worksheet.add_image(img, 'E2')

                messagebox.showinfo("Success", f"Excel file saved to: {file_path}")

    def save_results(self):
        # Save the Excel file
        excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if excel_path:
            self.save_excel_file()

        # Save the chart as an image
        chart_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
        if chart_path:
            self.fig.savefig(chart_path)
            messagebox.showinfo("Success", f"Chart saved to: {chart_path}")

    def create_pivot_chart(self):
        df = self.df[self.df['Promo Type'] != '']
        df['Promo Type'] = df['Promo Type'].apply(lambda x: 'Multibuy' if 'Multibuy' in x else x)
        pivot_table = df.pivot_table(index='Promo Type', values=['Line: Quantity', 'Line: Total'], aggfunc='sum')
        pivot_table = pivot_table.sort_values(by='Line: Total', ascending=False)

        self.ax.clear()
        labels = pivot_table.index.to_numpy()
        data = pivot_table['Line: Total'].to_numpy()
        wedges, texts, autotexts = self.ax.pie(data, labels=labels, autopct='%1.0f%%', pctdistance=0.85)

        self.ax.set_title("Promo Distribution", fontsize=16)
        self.fig.tight_layout()
        self.canvas.draw()

        result_text = "Promo Distribution:\n\n"
        for promo, total in zip(pivot_table.index, pivot_table['Line: Total']):
            result_text += f"{promo}: ${total:.2f}\n"

        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, result_text)

    # Promo analysis functions (unchanged)
    def analyze_taf25(self, df):
        df['Discount Ratio'] = (df['Line: Discount per Item'].abs() / df['Line: Price']).round(2)
        mask = (df['Discount Ratio'] >= 0.24) & (df['Discount Ratio'] <= 0.26) & (df['Line: Variant Compare At Price'] > 0)
        df.loc[mask, 'Promo Type'] = "TAF25"
        df.drop(columns=['Discount Ratio'], inplace=True)
        return df

    def analyze_25_percent_off(self, df):
        df['Discount Ratio'] = (df['Line: Discount per Item'].abs() / df['Line: Price']).round(2)
        mask = (df['Discount Ratio'] >= 0.24) & (df['Discount Ratio'] <= 0.26) & (df['Line: Variant Compare At Price'] == 0)
        df.loc[mask, 'Promo Type'] = "25% Off Selected Styles"
        df.drop(columns=['Discount Ratio'], inplace=True)
        return df

    def analyze_gift_card(self, df):
        mask = df['Line: Title'] == 'Gift Card'
        df.loc[mask, 'Promo Type'] = "Gift Card"
        return df

    def analyze_md_purchase(self, df):
        mask = (df['Line: Type'] == 'Line Item') & (~df['Line: Variant Compare At Price'].isnull()) & (df['Line: Variant Compare At Price'] != 0)
        df.loc[mask, 'Promo Type'] = "MD Purchase"
        return df

    def analyze_promo_code(self, df):
        promo_code_ids = set()
        promo_code_mask = df['Line: Name'].isin(['UNIDAYS', 'UNIDAYS20'])
        non_line_item_mask = df['Line: Type'] != 'Line Item'
        promo_code_ids.update(df.loc[promo_code_mask & non_line_item_mask, 'ID'].tolist())
        df.loc[df['ID'].isin(promo_code_ids), 'Promo Type'] = "Promo Code"
        return df

    def analyze_50_50(self, df):
        mask_not_null = df['Line: Product Tags'].notnull()
        mask = mask_not_null & df['Line: Product Tags'].str.contains(r'\b5050Jul24\b', case=False, regex=True)
        df.loc[mask, 'Promo Type'] = "50% Off 50 Styles"
        return df

    def analyze_sublime_suits(self, df):
        mask = df['Line: Product Tags'].str.contains(r'\bautomatic:(\$399|\$599) Suits\b', case=False) & \
               df['Line: Discount per Item'].isin([-174.50, -199.50, -249.50])
        df.loc[mask, 'Promo Type'] = "$399 & $599 Suits"
        return df

    def analyze_chino_25_percent_off(self, df):
        df['Discount Ratio'] = (df['Line: Discount per Item'].abs() / df['Line: Price']).round(2)
        mask = (df['Discount Ratio'] >= 0.24) & (df['Discount Ratio'] <= 0.26) & df['Line: Product Type'].str.contains('Chino', case=False) & (df['Line: Variant Compare At Price'] == 0)
        df.loc[mask, 'Promo Type'] = "25% Off Chinos"
        df.drop(columns=['Discount Ratio'], inplace=True)
        return df

    def analyze_25_percent_off_coats(self, df):
        df['Discount Ratio'] = (df['Line: Discount per Item'].abs() / df['Line: Price']).round(2)
        mask = (df['Discount Ratio'] >= 0.24) & (df['Discount Ratio'] <= 0.26) & df['Line: Product Type'].str.contains('Outerwear', case=False) & (df['Line: Variant Compare At Price'] == 0)
        df.loc[mask, 'Promo Type'] = "25% Off Coats/Outerwear"
        df.drop(columns=['Discount Ratio'], inplace=True)
        return df

    def analyze_25_percent_off_winter_tailoring(self, df):
        df['Discount Ratio'] = (df['Line: Discount per Item'].abs() / df['Line: Price']).round(2)
        mask = (
                (df['Discount Ratio'] >= 0.24) &
                (df['Discount Ratio'] <= 0.26) &
                (df['Line: Product Tags'].str.contains(r'\b25OFFWINTERTAILORING\b', case=False, regex=True)) &
                (df['Line: Variant Compare At Price'] == 0)
        )
        df.loc[mask, 'Promo Type'] = "25% Off Tailoring"
        df.drop(columns=['Discount Ratio'], inplace=True)
        return df

    def analyze_40_percent_off_tailoring(self, df):
        df['Discount Ratio'] = (df['Line: Discount per Item'].abs() / df['Line: Price']).round(2)
        mask = (
                (df['Discount Ratio'] >= 0.39) &
                (df['Discount Ratio'] <= 0.41) &
                (df['Line: Product Tags'].str.contains(r'\b40_Off_Tailoring_May24\b', case=False, regex=True)) &
                (df['Line: Variant Compare At Price'] == 0)
        )
        df.loc[mask, 'Promo Type'] = "40% Off Tailoring"
        df.drop(columns=['Discount Ratio'], inplace=True)
        return df

    def analyze_25_percent_off_knits(self, df):
        df['Discount Ratio'] = (df['Line: Discount per Item'].abs() / df['Line: Price']).round(2)
        mask = (df['Discount Ratio'] >= 0.24) & (df['Discount Ratio'] <= 0.26) & df['Line: Product Type'].str.contains('Knitwear', case=False) & (df['Line: Variant Compare At Price'] == 0)
        df.loc[mask, 'Promo Type'] = "Knits Offer"
        df.drop(columns=['Discount Ratio'], inplace=True)
        return df

    def analyze_tee_multibuy(self, df):
        mask = df['Line: Title'].str.contains("Mattia", case=False) & (df['Line: Total'] % 40 == 0) & (df['Line: Total'] != 0)
        df.loc[mask, 'Promo Type'] = "Tee Multibuy"
        return df

    def analyze_shirts_multibuy(self, df):
        mask = df['Line: Product Type'].str.contains('Shirts', case=False) & (df['Line: Discount per Item'] == -30)
        df.loc[mask, 'Promo Type'] = "Shirts Multibuy"
        return df

    def analyze_chino_multibuy(self, df):
        mask = df['Line: Product Tags'].str.contains(r'\bdiscount:2_each_\$110\b', case=False, regex=True) & (df['Line: Total'] % 110 == 0)
        df.loc[mask, 'Promo Type'] = "Chino Multibuy"
        return df

    def analyze_linen_shirts_multibuy(self, df):
        mask = df['Line: Product Tags'].str.contains(r'\bdiscount:2_each_\$130\b', case=False, regex=True) & (df['Line: Total'] % 130 == 0)
        df.loc[mask, 'Promo Type'] = "Linen Shirts Multibuy"
        return df

    def analyze_polo_multibuy(self, df):
        mask = df['Line: Product Tags'].str.contains(r'\bdiscount:2_each_\$109\b', case=False, regex=True) & (df['Line: Total'] % 109.99 == 0)
        df.loc[mask, 'Promo Type'] = "Polo Multibuy"
        return df

    def analyze_casual_bottom_multibuy(self, df):
        df['Discount Ratio'] = (df['Line: Discount per Item'].abs() / df['Line: Price']).round(2)
        mask = (df['Discount Ratio'] >= 0.29) & (df['Discount Ratio'] <= 0.31) & df['Line: Product Type'].str.contains('Chino', case=False) & (df['Line: Variant Compare At Price'] == 0)
        df.loc[mask, 'Promo Type'] = "Casual Bottom Multibuy"
        df.drop(columns=['Discount Ratio'], inplace=True)
        return df

    def analyze_fp_purchase(self, df):
        df_filtered = df[df['Promo Type'] != 'Promo Code']
        fp_purchase_mask = df['Line: Variant Compare At Price'] == 0
        fp_purchase_mask &= (df['Line: Discount'] == 0) & (df['Line: Discount per Item'] == 0)
        df.loc[fp_purchase_mask, 'Promo Type'] = 'FP Purchase'
        return df

    def analyze_suit_multibuy(self, df):
        # Define the suit multibuy prices
        suit_multibuy_prices = [175, 200, 275, 350, 400, 425, 575, 700]

        # Apply Suit Multibuy tag
        mask = (df['Line: Title'].str.contains('Jacket|Trouser', case=False, na=False)) & \
               (df['Line: Total'].isin(suit_multibuy_prices))
        df.loc[mask, 'Promo Type'] = 'Suit Multibuy'
        return df

    def get_tier_group(self, tags):
        if pd.isna(tags):
            return 'Silver'
        if 'cx-tier-tier-1' in tags:
            return 'Silver'
        elif 'cx-tier-tier-2' in tags:
            return 'Gold'
        elif 'cx-tier-tier-3' in tags:
            return 'Platinum'
        else:
            return 'Silver'

    def drop_shipping_lines(self, df):
        return df[df['Line: Type'] != 'Shipping Line']

    def drop_discount_lines(self, df):
        return df[df['Line: Type'] != 'Discount']

if __name__ == "__main__":
    root = tk.Tk()
    app = PromoAnalysisTool(root)
    root.mainloop()