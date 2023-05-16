import tkinter as tk
from tkinter import filedialog, Text
from KPI_Controller import ExcelDataReader, Data_Manager, Emission_Factor  # Assuming KPI_Controller.py is in the same directory
from KPI_Controller import KPIs

class AppInterface:
    def __init__(self, root):
        self.env_data = []
        self.emission_factor = None
        self.root = root
        self.create_widgets()

    def create_widgets(self):
        # Create frames
        self.left_frame = tk.Frame(self.root)
        self.left_frame.pack(side=tk.LEFT)
        
        self.right_frame = tk.Frame(self.root)
        self.right_frame.pack(side=tk.RIGHT)

        self.emission_factor_label = tk.Label(self.right_frame, text="Emission Factors:")
        self.emission_factor_label.pack(pady=10)  # Add some padding for better visual separation

        self.choose_file_button = tk.Button(self.right_frame, text="Choose File", command=self.browse_file)
        self.choose_file_button.pack()

        self.data_tables_label = tk.Label(self.right_frame, text="Data Tables:")
        self.data_tables_label.pack(pady=10)  # Add some padding for better visual separation

        self.choose_files_button = tk.Button(self.right_frame, text="Choose Files", command=self.browse_files)
        self.choose_files_button.pack()

        self.file_list = tk.Listbox(self.right_frame)
        self.file_list.pack(pady=10)  # Add some padding for better visual separation

        self.calculate_button = tk.Button(self.right_frame, text="Calculate", command=self.calculate)
        self.calculate_button.pack()

        self.results = Text(self.left_frame)
        self.results.pack()

    def browse_file(self):
        self.emission_factor = filedialog.askopenfilename(initialdir="/", title="Select file")
        self.emission_factor_label['text'] = "Emission Factors: " + self.emission_factor

    def browse_files(self):
        self.env_data = filedialog.askopenfilenames(initialdir="/", title="Select files")
        self.file_list.delete(0, tk.END)  # Clear the listbox before adding new files
        for file_path in self.env_data:
            self.file_list.insert(tk.END, file_path)

    def calculate(self):
        if not self.emission_factor or not self.env_data:
            self.results.delete('1.0', tk.END)
            self.results.insert(tk.END, 'Please select Emission Factors and Data Tables files.')
            return

        kpi = KPIs(self.env_data, self.emission_factor)
        total_area = kpi.calculate_total_area()

        self.results.delete('1.0', tk.END)
        self.results.insert(tk.END, f"Total Area: {total_area:,.2f}")

        MS_NOX_Total, MS_SOX_Total, MS_PM_Total, MS_CO2_Total, MS_CH4_Total, MS_N2O_Total = kpi.calculate_emissions_from_mobile_fuel()
        self.results.insert(tk.END, f"\nA1 Air Emission from Mobile Source:")
        self.results.insert(tk.END, f"\nNOx emission: {MS_NOX_Total/1000:,.2f} KG")
        self.results.insert(tk.END, f"\nSOx emission: {MS_SOX_Total/1000:,.2f} KG")
        self.results.insert(tk.END, f"\nPM emission: {MS_PM_Total/1000:,.2f} KG")
        self.results.insert(tk.END, f"\nCO2 emission: {MS_CO2_Total/1000:,.2f} KG")
        self.results.insert(tk.END, f"\nCH4 emission: {MS_CH4_Total/1000:,.2f} KG")
        self.results.insert(tk.END, f"\nN2O emission: {MS_N2O_Total/1000:,.2f} KG")

        EC_CO2_Total, Energy_Total = kpi.calculate_emissions_from_energy_consumption()
        self.results.insert(tk.END, f"\nA2 Air Emission from Energy Consumption:")
        self.results.insert(tk.END, f"\nCO2 emission: {EC_CO2_Total/1000:,.2f} KG")
        self.results.insert(tk.END, f"\nTotal Energy: {Energy_Total/1000:,.2f} KG")

        Paper_Total = kpi.calculate_emissions_from_paper_consumption()
        self.results.insert(tk.END, f"\nB1 Emission from Paper Consumption:")
        self.results.insert(tk.END, f"\nTotal Paper: {Paper_Total/1000:,.2f} KG")

        Water_Total = kpi.calculate_emissions_from_water_consumption()
        self.results.insert(tk.END, f"\nB2 Emission from Water Consumption:")
        self.results.insert(tk.END, f"\nTotal Water: {Water_Total/1000:,.2f} m3")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppInterface(root)
    root.mainloop()

