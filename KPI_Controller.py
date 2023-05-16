from typing import Optional
import pandas as pd

language = "EN" # CN=Chinese / EN=English
nor_factor = "Gross floor area" # 可供出租面积(平方米) / Gross floor area
if language=="EN":
  mobile_fuel_consumption_col_name = 'Fuel Consumption'
  mobile_mileage_col_name = 'Distance Travelled during the period/km'
  mobile_fuel_id_col_name = 'Transportation License #'
  mobile_fuel_code_col_name = 'Types of Transportation'
  mobile_fuel_type_col_name = 'Fuel Types'
  energy_consumption_col_name = 'Energy Consumption'
  energy_location_col_name = 'Location'
  paper_usage_col_name = 'Usage (kg)'
  water_consumption_col_name = 'Water Consumption'
else:
  mobile_fuel_consumption_col_name = '燃料消耗值'
  mobile_mileage_col_name = '年內航行公里'
  mobile_fuel_id_col_name = '车牌号码'
  mobile_fuel_code_col_name = '运输工具分类'
  mobile_fuel_type_col_name = '燃料类别'
  energy_consumption_col_name = '能耗值'
  energy_location_col_name = '地區'
  paper_usage_col_name = '使用量(千克)'
  water_consumption_col_name = '水資源消耗'

class ExcelDataReader:
    def __init__(self, file_id):
        self.file_id = file_id

    def read_excel(self, sheet_name, header_row=0):  # Modify the header_row parameter
        data = pd.read_excel(self.file_id, sheet_name=sheet_name, header=header_row)  # Modify this line
        return data

    def get_data(self):
        data = self.read_excel(self.data_path)
        return data

class Data_Manager:
    def __init__(self, file_id, sheet_name):
        self.data_reader = ExcelDataReader(file_id)
        self.sheet_name = sheet_name
        self.sheet = self.data_reader.read_excel(sheet_name=self.sheet_name)
        self.populate_properties()

    def populate_properties(self):
        ncols = len(self.sheet.columns)
        for i in range(ncols):
            column_name = self.sheet.columns[i]
            if not self.sheet[column_name].isnull().all():
                values = self.sheet[column_name].dropna().tolist()[1:]
                setattr(self, str(column_name), values)  # Convert column_name to string

    def get_column_data(self, column_name):
        if hasattr(self, column_name):
            data = getattr(self, column_name)
            return [float(x) if isinstance(x, (int, float)) and str(x) != 'nan' else (x if str(x) != 'nan' and str(x) != '' else None) for x in data]
        else:
            raise AttributeError(f"{column_name} not found in the {self.sheet_name}")


    def calculate_total(self, column_name):
        if hasattr(self, column_name):
            return sum(getattr(self, column_name))
        else:
            return 0

    def get_emission_factor(self, emission_type, source_type, code):
      emission_code = emission_type + source_type + code
      return self.emission_factor.get_emission_factor(emission_code)
            
    def print_column_data(self, column_name):
        column_data = self.get_column_data(column_name)
        print(f"{column_name}: {column_data}")

class Emission_Factor:
    sheet_name = 'Emission_factors'

    def __init__(self, file_path: str):
        self.data_reader = pd.read_excel(file_path, sheet_name=self.sheet_name)
        self.data = self.data_reader.set_index('Code')

    def get_emission_factor(self, code: str) -> Optional[float]:
        try:
            return float(self.data.loc[code, 'Emission Factor'])
        except KeyError:
            return None
        
class Normalization_Factor(Data_Manager):
    def __init__(self, file_id):
        super().__init__(file_id, sheet_name='Normalization_Factor')

    def calculate_area(self): 
        return self.calculate_total(nor_factor)
    
class Mobile_Fuel(Data_Manager):
    def __init__(self, file_id, emission_factor_file_path):
        super().__init__(file_id, sheet_name='Mobile_Fuel')
        self.file_id = file_id
        self.emission_factor = Emission_Factor(emission_factor_file_path)

    def calculate_mobile_fuel_consumption(self):
        return self.calculate_total(mobile_fuel_consumption_col_name)

    def get_mobile_fuel_data(self):
        if hasattr(self, mobile_fuel_consumption_col_name):
            data = self.get_column_data(mobile_fuel_consumption_col_name)
            return [x if x is not None else 0 for x in data]
        else:
            raise AttributeError(f"{mobile_fuel_consumption_col_name} not found in the {self.sheet_name}")
    
    def get_mobile_mileage_data(self):
        if hasattr(self, mobile_mileage_col_name):
            data = self.get_column_data(mobile_mileage_col_name)
            return [x if x is not None else 0 for x in data]
        else:
            raise AttributeError(f"{mobile_mileage_col_name} not found in the {self.sheet_name}")

    def get_mobile_fuel_id_data(self):
        if hasattr(self, mobile_fuel_id_col_name):
            data = self.get_column_data(mobile_fuel_id_col_name)
            return [x if x is not None else "not found" for x in data]
        else:
            raise AttributeError(f"{mobile_fuel_id_col_name} not found in the {self.sheet_name}")
            
    def get_mobile_fuel_code_data(self):
        if hasattr(self, mobile_fuel_code_col_name):
            data = self.get_column_data(mobile_fuel_code_col_name)
            return [x if x is not None else "not found" for x in data]
        else:
            raise AttributeError(f"{mobile_fuel_code_col_name} not found in the {self.sheet_name}")

    def get_mobile_fuel_type_data(self):
        if hasattr(self, mobile_fuel_type_col_name):
            data = self.get_column_data(mobile_fuel_type_col_name)
            return [x if x is not None else "not found" for x in data]
        else:
            raise AttributeError(f"{mobile_fuel_type_col_name} not found in the {self.sheet_name}")

    def get_mobile_fuel_emission_factor(self, emission_type, code):
      return self.get_emission_factor(emission_type, "|Mobile Combustion Sources|", code)

    def get_mobile_fuel_emission_factors(self, emission_type, codes):
      return [self.get_mobile_fuel_emission_factor(emission_type, code) for code in codes]
    
class Energy_Consumption(Data_Manager):
    def __init__(self, file_id, emission_factor_file_path):
        super().__init__(file_id, sheet_name='Energy_Consumption')
        self.file_id = file_id
        self.emission_factor = Emission_Factor(emission_factor_file_path)
        
    def get_energy_data(self):
        if hasattr(self, energy_consumption_col_name):
            data = self.get_column_data(energy_consumption_col_name)
            return [x if x is not None else 0 for x in data]
        else:
            raise AttributeError(f"{energy_consumption_col_name} not found in the {self.sheet_name}")

    def get_location(self):
        if hasattr(self, energy_location_col_name):
            data = self.get_column_data(energy_location_col_name)
            return [x if x is not None else "not found" for x in data]
        else:
            raise AttributeError(f"{energy_location_col_name} not found in the {self.sheet_name}")

    def get_energy_consumption_emission_factor(self, emission_type, code):
      return self.get_emission_factor(emission_type, "|", code)

    def get_energy_consumption_emission_factors(self, emission_type, codes):
      return [self.get_energy_consumption_emission_factor(emission_type, code) for code in codes]
    
class Paper_Consumption(Data_Manager):
    def __init__(self, file_id, emission_factor_file_path):
        super().__init__(file_id, sheet_name='Paper_Usage')
        self.file_id = file_id
        self.emission_factor = Emission_Factor(emission_factor_file_path)
        
    def get_paper_data(self):
        if hasattr(self, paper_usage_col_name):
            data = self.get_column_data(paper_usage_col_name)
            return [x if x is not None else 0 for x in data]
        else:
            raise AttributeError(f"{paper_usage_col_name} not found in the {self.sheet_name}")

class Water_Consumption(Data_Manager):
    def __init__(self, file_id, emission_factor_file_path):
        super().__init__(file_id, sheet_name='Water_Consumption')
        self.file_id = file_id
        self.emission_factor = Emission_Factor(emission_factor_file_path)
        
    def get_paper_data(self):
        if hasattr(self, water_consumption_col_name):
            data = self.get_column_data(water_consumption_col_name)
            return [x if x is not None else 0 for x in data]
        else:
            raise AttributeError(f"{water_consumption_col_name} not found in the {self.sheet_name}")
        
class KPIs:
    def __init__(self, file_paths, emission_factor_file_path):
        self.file_paths = file_paths
        self.normalization_factors = [Normalization_Factor(file_path) for file_path in file_paths]
        self.mobile_fuel = [Mobile_Fuel(file_path, emission_factor_file_path) for file_path in file_paths]
        self.energy_consumption = [Energy_Consumption(file_path, emission_factor_file_path) for file_path in file_paths]
        self.paper_consumption = [Paper_Consumption(file_path, emission_factor_file_path) for file_path in file_paths]
        self.water_consumption = [Water_Consumption(file_path, emission_factor_file_path) for file_path in file_paths]

    def calculate_total_area(self):
        #print([nf.calculate_area() for nf in self.normalization_factors]) 
        total_area = sum([nf.calculate_area() for nf in self.normalization_factors])
        return total_area

    def calculate_total_mobile_fuel_consumption(self):
        total_mobile_fuel_consumption = sum([mf.calculate_mobile_fuel_consumption() for mf in self.mobile_fuel])
        return total_mobile_fuel_consumption

    def get_source_name(self, file_id):
        start = file_id.find("Environmental_") + len("Environmental_")
        end = file_id.find(".xlsx")
        return file_id[start:end]

    def calculate_emissions_from_mobile_fuel(self):        
        Fuel_KWH_Total = 0
        NOX_Total = 0
        SOX_Total = 0
        PM_Total = 0
        CO2_Total = 0
        CH4_Total = 0
        N2O_Total = 0
        for mf in self.mobile_fuel:
            try:
                source_name = self.get_source_name(mf.file_id)
                #print(f"{mf.sheet_name} - from {source_name}:")
                mobile_fuel_id_data = mf.get_mobile_fuel_id_data()
                mobile_fuel_code_data = mf.get_mobile_fuel_code_data()
                mobile_fuel_type_data = mf.get_mobile_fuel_type_data()
                mobile_fuel_combined_code = []
                for i in range(len(mobile_fuel_code_data)):
                    combined_element = mobile_fuel_code_data[i] + mobile_fuel_type_data[i]
                    mobile_fuel_combined_code.append(combined_element)
                #mobile_fuel_combined_code = [(code, fuel_type) for code, fuel_type in zip(mobile_fuel_code_data, mobile_fuel_type_data)]
                mobile_fuel_consumption_data = mf.get_mobile_fuel_data()
                mobile_fuel_mileage_data = mf.get_mobile_mileage_data()
                NOX_emission_factor = mf.get_mobile_fuel_emission_factors("NOx Emission", mobile_fuel_code_data)
                SOX_emission_factor = mf.get_mobile_fuel_emission_factors("SOx Emission", mobile_fuel_type_data)
                PM_emission_factor = mf.get_mobile_fuel_emission_factors("PM Emission", mobile_fuel_code_data)
                CO2_emission_factor = mf.get_mobile_fuel_emission_factors("CO2 Emission", mobile_fuel_type_data)
                CH4_emission_factor = mf.get_mobile_fuel_emission_factors("CH4 Emission", mobile_fuel_combined_code)
                N2O_emission_factor = mf.get_mobile_fuel_emission_factors("N2O Emission", mobile_fuel_combined_code)
                table = []
                Fuel_KWH_Subtotal = 0
                NOX_Subtotal = 0
                SOX_Subtotal = 0
                PM_Subtotal = 0
                CO2_Subtotal = 0
                CH4_Subtotal = 0
                N2O_Subtotal = 0
                for code, data, mileage, NOX_ef, SOX_ef, PM_ef, CO2_ef, CH4_ef, N2O_ef in zip(mobile_fuel_id_data, mobile_fuel_consumption_data, mobile_fuel_mileage_data, NOX_emission_factor, SOX_emission_factor, PM_emission_factor, CO2_emission_factor, CH4_emission_factor, N2O_emission_factor):                    
                    table.append([code, data, mileage, data*NOX_ef, data*SOX_ef, data*PM_ef, data*CO2_ef, data*CH4_ef*28, data*N2O_ef*256])
                    Fuel_KWH_Subtotal += data*9.11 #Net Calorific Value (kWh/L) = 9.11 for Unleaded Petrol
                    NOX_Subtotal += mileage*NOX_ef
                    SOX_Subtotal += data*SOX_ef
                    PM_Subtotal += mileage*PM_ef
                    CO2_Subtotal += data*CO2_ef
                    CH4_Subtotal += data*CH4_ef*28
                    N2O_Subtotal += data*N2O_ef*256
                #print(tabulate(table, headers=['ID', 'Consumption Data', 'Mileage', 'Emission(NOx)', 'Emission(SOx)', 'Emission(PM)', 'Emission(CO2)', 'Emission(CH4)', 'Emission(N2O)'], stralign='left'))
                print(f"Fuel usage in {source_name} is {Fuel_KWH_Subtotal} kWh")
                Fuel_KWH_Total += Fuel_KWH_Subtotal
                NOX_Total += NOX_Subtotal
                SOX_Total += SOX_Subtotal
                PM_Total += PM_Subtotal
                CO2_Total += CO2_Subtotal
                CH4_Total += CH4_Subtotal
                N2O_Total += N2O_Subtotal
                #print(f"NOx emission: {NOX_Subtotal:,.2f}")
                #print(f"SOx emission: {SOX_Subtotal:,.2f}")
                #print(f"PM emission: {PM_Subtotal:,.2f}")
                #print(f"CO2 emission: {CO2_Subtotal:,.2f}")
                #print(f"CH4 emission: {CH4_Subtotal:,.2f}")
                #print(f"N2O emission: {N2O_Subtotal:,.2f}")
            except Exception as e:
                print(f"Error occurred while processing {mf.sheet_name}")
                print(f"Error: {e}")
                print(traceback.format_exc())
        return NOX_Total, SOX_Total, PM_Total, CO2_Total, CH4_Total, N2O_Total


    def calculate_emissions_from_energy_consumption(self):
        CO2_Total = 0
        Energy_Total = 0
        for ec in self.energy_consumption:
            try:
                source_name = self.get_source_name(ec.file_id)
                #print(f"{ec.sheet_name} - from {source_name}:")
                energy_consumption_data = ec.get_energy_data()
                energy_consumption_location = ec.get_location()
                CO2_emission_factor = ec.get_energy_consumption_emission_factors("Purchased Electricity", energy_consumption_location)
                table = []
                CO2_Subtotal = 0
                Energy_subtotal = 0
                for data, location, CO2_ef in zip(energy_consumption_data, energy_consumption_location, CO2_emission_factor):                    
                  if data is not None and CO2_ef is not None:
                      table.append([data, location, data*CO2_ef])
                      Energy_subtotal += data
                      CO2_Subtotal += data*CO2_ef
                  else:
                      table.append([data, location, None])
                #print(tabulate(table, headers=['Consumption Data', 'Location', 'Emission(CO2)'], stralign='left'))
                print(f"Energy usage in {source_name} is {Energy_subtotal} kWh")
                CO2_Total += CO2_Subtotal
                Energy_Total += Energy_subtotal
                #print(f"CO2 emission: {CO2_Subtotal:,.2f}")
            except Exception as e:
                print(f"Error occurred while processing {ec.sheet_name}")
                print(f"Error: {e}")
                print(traceback.format_exc())
        return CO2_Total, Energy_Total

    def calculate_emissions_from_paper_consumption(self):
        Paper_Total = 0
        for pc in self.paper_consumption:
            try:
                source_name = self.get_source_name(pc.file_id)
                paper_consumption_data = pc.get_paper_data()
                table = []
                Paper_Subtotal = 0
                for data in paper_consumption_data:                    
                    if data is not None:
                        table.append([data])
                        Paper_Subtotal += data/1000 #data is in gram
                    else:
                        table.append([None])
                Paper_Total += Paper_Subtotal
                print(source_name, Paper_Subtotal/1000)
            except Exception as e:
                print(f"Error occurred while processing {pc.sheet_name}")
                print(f"Error: {e}")
                print(traceback.format_exc())
        return Paper_Total

    def calculate_emissions_from_water_consumption(self):
        Water_Total = 0
        for wc in self.water_consumption:
            try:
                source_name = self.get_source_name(wc.file_id)
                water_consumption_data = wc.get_paper_data()
                table = []
                Water_Subtotal = 0
                for data in water_consumption_data:                    
                    if data is not None:
                        table.append([data])
                        Water_Subtotal += data #data is in m3
                    else:
                        table.append([None])
                Water_Total += Water_Subtotal
                print(source_name, Water_Subtotal)
            except Exception as e:
                print(f"Error occurred while processing {wc.sheet_name}")
                print(f"Error: {e}")
                print(traceback.format_exc())
        return Water_Total
    
