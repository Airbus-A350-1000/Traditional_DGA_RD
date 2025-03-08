import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
import matplotlib.image as mpimg
import base64
from io import BytesIO

class DGACalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("DGA Calculator")
        self.root.geometry("1000x700")
        self.root.resizable(True, True)
        
        # Initialize variables
        self.excel_file = None
        self.df = None
        self.result_df = None
        
        # Create notebook (tabs)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)
        
        # Create main tab
        self.main_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.main_tab, text="Main")
        
        # Create instructions tab
        self.instructions_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.instructions_tab, text="Instructions")
        
        # Set up the main tab
        self.setup_main_tab()
        
        # Set up the instructions tab
        self.setup_instructions_tab()
        
        # Base64 encoded image string of Duval Triangle
        self.duval_triangle_base64 = "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"
    
    def setup_main_tab(self):
        # Frame for file selection
        file_frame = ttk.LabelFrame(self.main_tab, text="File Selection")
        file_frame.pack(fill="x", padx=10, pady=10)
        
        # File selection button
        self.file_path_var = tk.StringVar()
        ttk.Label(file_frame, text="Excel File:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)
        
        # Frame for analysis options
        analysis_frame = ttk.LabelFrame(self.main_tab, text="Analysis Options")
        analysis_frame.pack(fill="x", padx=10, pady=10)
        
        # Analysis type selection
        self.analysis_type = tk.StringVar(value="Duval Triangle")
        ttk.Label(analysis_frame, text="Analysis Type:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        analysis_combo = ttk.Combobox(analysis_frame, textvariable=self.analysis_type, state="readonly")
        analysis_combo["values"] = ("Duval Triangle", "Rogers Ratio", "Both")
        analysis_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # Buttons frame
        button_frame = ttk.Frame(self.main_tab)
        button_frame.pack(fill="x", padx=10, pady=10)
        
        # Analyze button
        ttk.Button(button_frame, text="Analyze", command=self.perform_analysis).pack(side="left", padx=5)
        
        # Save results button
        ttk.Button(button_frame, text="Save Results", command=self.save_results).pack(side="left", padx=5)
        
        # Results section
        results_frame = ttk.LabelFrame(self.main_tab, text="Results")
        results_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create a frame for the plots
        self.plot_frame = ttk.Frame(results_frame)
        self.plot_frame.pack(side="left", fill="both", expand=True)
        
        # Create a frame for the results table
        table_frame = ttk.Frame(results_frame)
        table_frame.pack(side="right", fill="both", expand=True)
        
        # Create Treeview for results
        self.result_tree = ttk.Treeview(table_frame)
        scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.result_tree.yview)
        scroll_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        self.result_tree.pack(fill="both", expand=True)
    
    def setup_instructions_tab(self):
        # Create a text widget for instructions
        instructions_text = tk.Text(self.instructions_tab, wrap="word", padx=10, pady=10)
        instructions_text.pack(fill="both", expand=True)
        
        # Add a scrollbar
        scrollbar = ttk.Scrollbar(instructions_text, command=instructions_text.yview)
        instructions_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        
        # Instructions content
        instructions = """
DGA Calculator - Instructions

1. File Format Requirements:
   - The Excel file should contain columns with the following gas concentrations (in ppm):
     • H2 (Hydrogen)
     • CH4 (Methane)
     • C2H6 (Ethane)
     • C2H4 (Ethylene)
     • C2H2 (Acetylene)
     • CO (Carbon Monoxide)
     • CO2 (Carbon Dioxide)
   - Each row should represent a different transformer or sample
   - Optional: Include 'ID' or 'Transformer_ID' column for identification

2. Using the Application:
   a) Click "Browse" to select your Excel file
   b) Choose the analysis type:
      • Duval Triangle - Identifies thermal and electrical faults
      • Rogers Ratio - Identifies fault types based on gas ratios
      • Both - Performs both analyses
   c) Click "Analyze" to process the data
   d) View the results in the table and graphical display
   e) Click "Save Results" to update your Excel file with the analysis results

3. Understanding the Results:
   
   Duval Triangle Diagnoses:
   • PD: Partial Discharges
   • D1: Low Energy Discharges
   • D2: High Energy Discharges
   • T1: Thermal Fault <300°C
   • T2: Thermal Fault 300-700°C
   • T3: Thermal Fault >700°C
   • DT: Mix of Thermal and Electrical Faults

   Rogers Ratio Diagnoses:
   • Normal Deterioration
   • Partial Discharge
   • Slight Overheating <150°C
   • Overheating 150°C-200°C
   • Overheating 200°C-300°C
   • General Conductor Overheating
   • Winding Circulating Currents
   • Core and tank circulating currents, overheated joints
   • Flashover without power follow through
   • Arc with power follow through
   • Continuous sparking to floating potential
   • Partial discharge with tracking (note CO)

   

4. Troubleshooting:
   • Ensure all required gas concentrations are present in your file
   • Gas values should be numeric (no text or special characters)
   • If the application cannot process your file, check the format and try again
   • Missing values may cause inaccurate results - ensure data completeness
   
References:
Duval Triangle Reference Plot: [1] Selim Koroglu, “Selim Koroglu * -(OHFWULFDO6\VWHPV Regular paper A Case Study on Fault Detection in Power Transformers Using Dissolved Gas Analysis and Electrical Test Methods,” Journal of Electrical Systems, vol. 12, no. 3, Aug. 2016, Available: https://www.researchgate.net/publication/308901223_Selim_Koroglu_-OHFWULFDO6VWHPV_Regular_paper_A_Case_Study_on_Fault_Detection_in_Power_Transformers_Using_Dissolved_Gas_Analysis_and_Electrical_Test_Methods

Rogers Ratio Threshold Values and Classification: [2] N. A. Muhamad, B. T. Phung, T. R. Blackburn and K. X. Lai, "Comparative Study and Analysis of DGA Methods for Transformer Mineral Oil," 2007 IEEE Lausanne Power Tech, Lausanne, Switzerland, 2007, pp. 45-50, doi: 10.1109/PCT.2007.4538290. keywords: {Dissolved gas analysis;Oil insulation;Minerals;Petroleum;Gases;Power transformer insulation;Circuit faults;Hydrogen;Testing;Electrical fault detection},

Code was developed with assitance from Github Copilot.
For any questions or issues, please contact me at m.mohamed1@qatar.tamu.edu
        """
        
        instructions_text.insert("1.0", instructions)
        instructions_text.configure(state="disabled")  # Make text read-only
    
    def browse_file(self):
        filetypes = [("Excel files", "*.xlsx *.xls")]
        self.excel_file = filedialog.askopenfilename(filetypes=filetypes)
        if self.excel_file:
            self.file_path_var.set(self.excel_file)
            try:
                self.df = pd.read_excel(self.excel_file)
                # Check if required columns exist
                required_columns = ['H2', 'CH4', 'C2H6', 'C2H4', 'C2H2']
                missing_columns = [col for col in required_columns if col not in self.df.columns]
                
                if missing_columns:
                    messagebox.showerror("Missing Columns", 
                                         f"The following required columns are missing: {', '.join(missing_columns)}")
                    self.df = None
                else:
                    messagebox.showinfo("File Loaded", "Excel file loaded successfully.")
                    self.update_result_tree(self.df)
            except Exception as e:
                messagebox.showerror("Error", f"Error loading file: {str(e)}")
                self.df = None
    
    def update_result_tree(self, dataframe):
        # Clear existing data
        for i in self.result_tree.get_children():
            self.result_tree.delete(i)
        
        # Update columns
        self.result_tree["columns"] = list(dataframe.columns)
        self.result_tree["show"] = "headings"
        
        # Set column headings
        for column in dataframe.columns:
            self.result_tree.heading(column, text=column)
            self.result_tree.column(column, width=100)
        
        # Add data rows
        for i, row in dataframe.iterrows():
            values = [row[column] for column in dataframe.columns]
            self.result_tree.insert("", "end", values=values)
    
    def perform_analysis(self):
        if self.df is None:
            messagebox.showerror("Error", "Please load a valid Excel file first.")
            return
        
        analysis_choice = self.analysis_type.get()
        self.result_df = self.df.copy()
        
        # Perform the requested analysis
        if analysis_choice in ["Duval Triangle", "Both"]:
            self.result_df = self.perform_duval_analysis(self.result_df)
            self.plot_duval_triangle()
        
        if analysis_choice in ["Rogers Ratio", "Both"]:
            self.result_df = self.perform_rogers_analysis(self.result_df)
        
        # Update the results table
        self.update_result_tree(self.result_df)
        
    def perform_duval_analysis(self, df):
    # Function to calculate Duval Triangle diagnosis based on the correct boundaries
        def duval_diagnosis(ch4_ppm, c2h4_ppm, c2h2_ppm):
            # Check if we have valid input
            total = ch4_ppm + c2h4_ppm + c2h2_ppm
            if total == 0:
                return "Insufficient data"
            
            # Normalize to percentages
            ch4_pct = 100 * ch4_ppm / total
            c2h4_pct = 100 * c2h4_ppm / total
            c2h2_pct = 100 * c2h2_ppm / total  # tr
            
            # Duval Triangle boundaries - these match defintions found in the literature review
            # PD: Partial Discharges region
            if ch4_pct >= 98 and c2h2_pct <= 2:
                return "PD: Partial Discharges"
            
            # D1: Low Energy Discharges region
            elif c2h2_pct >= 13 and c2h4_pct <= 23:
                return "D1: Low Energy Discharges"
            
            # D2: High Energy Discharges region
            elif (c2h2_pct >= 13 and ch4_pct <= 40 and c2h4_pct >=23) or c2h2_pct >=30:
                return "D2: High Energy Discharges"
            
            # T1: Thermal Fault <300°C region
            elif c2h2_pct <= 4 and c2h4_pct < 20 and ch4_pct < 98:
                return "T1: Thermal Fault <300°C"
            
            # T2: Thermal Fault 300-700°C region
            elif c2h2_pct <= 4 and c2h4_pct <= 50 and c2h4_pct >= 20:
                return "T2: Thermal Fault 300-700°C"
            
            # T3: Thermal Fault >700°C region
            elif c2h2_pct <= 15 and c2h4_pct >= 50:
                return "T3: Thermal Fault >700°C"
            
            else:
                return "DT: Mix of Thermal and Electrical Faults"
        
        # Apply diagnosis to each row
        duval_results = []
        duval_coords = []
        
        for i, row in df.iterrows():
            ch4 = row['CH4']
            c2h4 = row['C2H4']
            c2h2 = row['C2H2']
            
            diagnosis = duval_diagnosis(ch4, c2h4, c2h2)
            duval_results.append(diagnosis)
            
            # Store coordinates for plotting
            total = ch4 + c2h4 + c2h2
            if total > 0:
                ch4_pct = 100 * ch4 / total
                c2h4_pct = 100 * c2h4 / total
                c2h2_pct = 100 * c2h2 / total
                duval_coords.append((ch4_pct, c2h4_pct, c2h2_pct))
            else:
                duval_coords.append((None, None, None))
        
        # Add results to dataframe
        df['Duval_Diagnosis'] = duval_results
        
        # Store coordinates for plotting
        self.duval_coords = duval_coords
        
        return df
    def perform_rogers_analysis(self, df):
        # Calculate the gas ratios
        def calculate_ratios(row):
            try:
                ch4_h2 = row['CH4'] / row['H2'] if row['H2'] > 0 else 0
                c2h6_ch4 = row['C2H6'] / row['CH4'] if row['CH4'] > 0 else 0
                c2h4_c2h6 = row['C2H4'] / row['C2H6'] if row['C2H6'] > 0 else 0
                c2h2_c2h4 = row['C2H2'] / row['C2H4'] if row['C2H4'] > 0 else 0
                
                return ch4_h2, c2h6_ch4, c2h4_c2h6, c2h2_c2h4
            except:
                return 0, 0, 0, 0
        
        # Rogers ratio diagnosis function based on literature review tables
        def rogers_diagnosis(ch4_h2, c2h6_ch4, c2h4_c2h6, c2h2_c2h4):
            # Codify the ratios according to Table 3
            def codify_ratio_i(ratio):  # CH4/H2
                if ratio < 0.1:
                    return 5
                elif 0.1 <= ratio < 1.0:
                    return 0
                elif 1.0 <= ratio < 3.0:
                    return 1
                else:  # ratio >= 3.0
                    return 2
                    
            def codify_ratio_j(ratio):  # C2H6/CH4
                if ratio < 1.0:
                    return 0
                else:  # ratio >= 1.0
                    return 1
                    
            def codify_ratio_k(ratio):  # C2H4/C2H6
                if ratio < 1.0:
                    return 0
                elif 1.0 <= ratio < 3.0:
                    return 1
                else:  # ratio >= 3.0
                    return 2
                    
            def codify_ratio_l(ratio):  # C2H2/C2H4
                if ratio < 0.5:
                    return 0
                elif 0.5 <= ratio < 3.0:
                    return 1
                else:  # ratio >= 3.0
                    return 2
            
            # Codify each ratio
            i = codify_ratio_i(ch4_h2)
            j = codify_ratio_j(c2h6_ch4)
            k = codify_ratio_k(c2h4_c2h6)
            l = codify_ratio_l(c2h2_c2h4)
            
            # Rogers diagnosis table based on Table 4
            code_to_diagnosis = {
                (0, 0, 0, 0): "Normal deterioration",
                (5, 0, 0, 0): "Partial discharge",
                (1, 0, 0, 0): "Slight overheating <150°C",
                (2, 0, 0, 0): "Slight overheating <150°C",
                (1, 1, 0, 0): "Overheating 150°C-200°C",
                (2, 1, 0, 0): "Overheating 150°C-200°C",
                (0, 1, 0, 0): "Overheating 200°C-300°C",
                (0, 0, 1, 0): "General conductor overheating",
                (1, 0, 1, 0): "Winding circulating currents",
                (1, 0, 2, 0): "Core and tank circulating currents, overheated joints",
                (0, 0, 0, 1): "Flashover without power follow through",
                (0, 0, 1, 1): "Arc with power follow through",
                (0, 0, 2, 1): "Arc with power follow through",
                (0, 0, 1, 2): "Arc with power follow through",
                (0, 0, 2, 2): "Continuous sparking to floating potential",
                (5, 0, 0, 1): "Partial discharge with tracking (note CO)",
                (5, 0, 0, 2): "Partial discharge with tracking (note CO)"
            }
            
            code = (i, j, k, l)
            
            # Return diagnosis or closest match
            if code in code_to_diagnosis:
                return code_to_diagnosis[code]
            else:
                # Find closest match based on similarity of codes
                min_distance = float('inf')
                closest_diagnosis = "Undefined Fault Pattern"
                
                for known_code, diagnosis in code_to_diagnosis.items():
                    distance = sum((a - b) ** 2 for a, b in zip(code, known_code))
                    if distance < min_distance:
                        min_distance = distance
                        closest_diagnosis = f"{diagnosis} (closest match)"
                
                return closest_diagnosis
        
        # Apply to each row
        rogers_results = []
        ratio_values = []
        
        for i, row in df.iterrows():
            ch4_h2, c2h6_ch4, c2h4_c2h6, c2h2_c2h4 = calculate_ratios(row)
            diagnosis = rogers_diagnosis(ch4_h2, c2h6_ch4, c2h4_c2h6, c2h2_c2h4)
            rogers_results.append(diagnosis)
            #ratio_values.append((ch4_h2, c2h6_ch4, c2h4_c2h6, c2h2_c2h4))
        
        # Add results to dataframe
        df['Rogers_Diagnosis'] = rogers_results
        #df['CH4/H2'] = [r[0] for r in ratio_values]
        #df['C2H6/CH4'] = [r[1] for r in ratio_values]
        #df['C2H4/C2H6'] = [r[2] for r in ratio_values]
        #df['C2H2/C2H4'] = [r[3] for r in ratio_values]
        #Uncomment to check ratio calculations if needed
        
        # Store ratio values
        self.rogers_ratios = ratio_values
        
        return df
    
    def plot_duval_triangle(self):
        # Clear existing plots
        for widget in self.plot_frame.winfo_children():
            widget.destroy()
        
        # Create a figure and axis
        fig, ax = plt.subplots(figsize=(6, 5))
        
        # Load the image from base64 string
        if self.duval_triangle_base64:
            try:
                image_data = base64.b64decode(self.duval_triangle_base64)
                img = mpimg.imread(BytesIO(image_data), format='png')
                ax.imshow(img)
            except Exception as e:
                ax.text(0.5, 0.5, "Failed to load Duval Triangle image.", ha='center', va='center')
                ax.axis('off')
        else:
            ax.text(0.5, 0.5, "Duval Triangle image not provided.", ha='center', va='center')
            ax.axis('off')
        
        # Define the triangle corners based on the actual image coordinates
        # These values are based on the embedded image of the Duval Triangle
        bottom_left = (41, 365)   # 100% C2H2 corner
        bottom_right = (380, 365) # 100% C2H4 corner
        top = (211, 25.2)         # 100% CH4 corner
        
        # Function to convert ternary coordinates to image coordinates
        def ternary_to_image(ch4_pct, c2h4_pct, c2h2_pct):
            # Ensure percentages sum to 100
            total = ch4_pct + c2h4_pct + c2h2_pct
            if total == 0:
                return None, None
                
            if total != 100:
                ch4_pct = 100 * ch4_pct / total
                c2h4_pct = 100 * c2h4_pct / total
                c2h2_pct = 100 * c2h2_pct / total
            
            # Calculate position using coordinates
            x = (top[0] * (ch4_pct/100)) + (bottom_right[0] * (c2h4_pct/100)) + (bottom_left[0] * (c2h2_pct/100))
            y = (top[1] * (ch4_pct/100)) + (bottom_right[1] * (c2h4_pct/100)) + (bottom_left[1] * (c2h2_pct/100))
            
            return x, y
        
        # Plot the data points
        for coords in self.duval_coords:
            if None not in coords:
                ch4_pct, c2h4_pct, c2h2_pct = coords
                x, y = ternary_to_image(ch4_pct, c2h4_pct, c2h2_pct)
                if x is not None and y is not None:
                    ax.plot(x, y, 'ro', markersize=3)

             
        # Add a title and disable axis
        ax.set_title("Duval Triangle Analysis")
        ax.axis('off')
        
        # Add the figure to the GUI
        canvas = FigureCanvasTkAgg(fig, master=self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
    
    def save_results(self):
        if self.result_df is None:
            messagebox.showerror("Error", "No analysis results to save.")
            return
        
        if self.excel_file:
            try:
                # Save to the same file path
                self.result_df.to_excel(self.excel_file, index=False)
                messagebox.showinfo("Success", "Results saved to the original Excel file.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save results: {str(e)}")
        else:
            messagebox.showerror("Error", "No source file selected.")

if __name__ == "__main__":
    root = tk.Tk()
    app = DGACalculator(root)
    root.mainloop()
