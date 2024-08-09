#!/usr/bin/env python3

import pandas as pd
from datetime import timedelta
from tkinter import Tk, filedialog, messagebox

class ReactantConsumption:
    def __init__(self, stock_data, supply_data):
        self.stock_data = stock_data
        self.supply_data = supply_data
        self.consumption_summary = pd.DataFrame()
        self.zero_stock_day_fc_sup_pln_cons = {}
        self.zero_stock_day_fc_sup_actl_cons = {}
        self.zero_stock_day_actl_sup_pln_cons = {}
        self.zero_stock_day_actl_sup_actl_cons = {}

    # Main fucntion to calculate consumption
    def calculate_daily_consumption(self):
        data = []
        for reactant, attributes in self.stock_data.iterrows():
            reactant_name = attributes['reactant_name']
            unit_of_measure = attributes['unit_of_measure']
            date = attributes['stock_update_date']
            stock_supply = attributes['stock_amount']
            daily_planned_consumption = attributes['consumption_rate_plan']
            daily_actual_consumption = attributes['consumption_rate_actual']
            threshold_days = attributes['threshold_days']
            accounting_period = attributes['accounting_period']
            
            '''
            Using data frame (converted into dictionary) as source for further calculations:
            stock_supply: the amount of supply allocated in the stock at the start of the calculation (stock_update_date)
            daily_planned_consumption: the amount of planned daily consumption
            daily_actual_consumption: the amount of actual daily consumption
            threshold_days: the amount of days needed to perform new purchase and delivery. using as a thershold to start new purchase before supply is depleted
            accounting_period: the amount of days to perform all calculations
            '''
            
            # Filter supply (forecast + actual) data for the current reactant
            reactant_supply_data = self.supply_data[self.supply_data['reactant_code'] == reactant]

            '''
            Variable name decoding:
            stock_fc_sup_pln_cons: stock amount + forecast supply - planned consumption
            stock_fc_sup_actl_cons: stock amount + forecast supply - actual consumption
            stock_actl_sup_pln_cons: stock amount + actual supply - planned consumption
            stock_actl_sup_actl_cons: stock amount + actual supply - actual consumption
            '''

            # Initialize stock supply with various options
            stock_fc_sup_pln_cons = stock_supply
            stock_fc_sup_actl_cons = stock_supply
            stock_actl_sup_pln_cons = stock_supply
            stock_actl_sup_actl_cons = stock_supply
            
            zero_stock_day_fc_sup_pln_cons = None
            zero_stock_day_fc_sup_actl_cons = None
            zero_stock_day_actl_sup_pln_cons = None
            zero_stock_day_actl_sup_actl_cons = None

            # Calculate daily consumption for each day in the accounting period
            for day in range(accounting_period):
                # Iterate over each day and check for forecast and actual supply and update stock
                current_date = date + timedelta(days = day)
                forecast_supply_daily = reactant_supply_data[reactant_supply_data['date_of_delivery_forecast'] == current_date]['quantity'].sum()
                actual_supply_daily = reactant_supply_data[reactant_supply_data['date_of_delivery_actual'] == current_date]['quantity'].sum()

                stock_fc_sup_pln_cons += forecast_supply_daily
                stock_fc_sup_actl_cons += forecast_supply_daily
                stock_actl_sup_pln_cons += actual_supply_daily
                stock_actl_sup_actl_cons += actual_supply_daily

                '''
                Variable name decoding:
                rmn_stock_fc_sup_pln_cons: remaining stock amount + forecast supply - planned consumption 
                rmn_stock_fc_sup_actl_cons: remaining stock amount + forecast supply - actual consumption
                rmn_stock_actl_sup_pln_cons: remaining stock amount + actual supply - planned consumption
                rmn_stock_actl_sup_actl_cons: remaining stock amount + actual supply - actual consumption
                '''

                # Calculate remaining stock amount after planned and actual consumption
                # Incorporate check: if current stock after daily consumption falls below zero, return 0. Stock amount can not fall to negative values
                rmn_stock_fc_sup_pln_cons = max(0, stock_fc_sup_pln_cons - daily_planned_consumption)
                rmn_stock_fc_sup_actl_cons = max(0, stock_fc_sup_actl_cons - daily_actual_consumption)
                rmn_stock_actl_sup_pln_cons = max(0, stock_actl_sup_pln_cons - daily_planned_consumption)
                rmn_stock_actl_sup_actl_cons = max(0, stock_actl_sup_actl_cons - daily_actual_consumption)
                
                # Determine if new supply needs to be contracted based on contracting period days
                need_new_supply_fc_sup_pln_cons = rmn_stock_fc_sup_pln_cons <= threshold_days * daily_planned_consumption
                need_new_supply_fc_sup_actl_cons = rmn_stock_fc_sup_actl_cons <= threshold_days * daily_actual_consumption
                need_new_supply_actl_sup_pln_cons = rmn_stock_actl_sup_pln_cons <= threshold_days * daily_planned_consumption
                need_new_supply_actl_sup_actl_cons = rmn_stock_actl_sup_actl_cons <= threshold_days * daily_actual_consumption

                # Check if stock amount becomes zero and return day number
                if zero_stock_day_fc_sup_pln_cons is None and rmn_stock_fc_sup_pln_cons == 0:
                    zero_stock_day_fc_sup_pln_cons = current_date
                if zero_stock_day_fc_sup_actl_cons is None and rmn_stock_fc_sup_actl_cons == 0:
                    zero_stock_day_fc_sup_actl_cons = current_date
                if zero_stock_day_actl_sup_pln_cons is None and rmn_stock_actl_sup_pln_cons == 0:
                    zero_stock_day_actl_sup_pln_cons = current_date                    
                if zero_stock_day_actl_sup_actl_cons is None and rmn_stock_actl_sup_actl_cons == 0:
                    zero_stock_day_actl_sup_actl_cons = current_date

                data.append({
                    'Reactant code': reactant,
                    'Reactant name': reactant_name,
                    'Unit of measure': unit_of_measure,
                    'Day': current_date,
                    'Stock and future supply using planned consumption': rmn_stock_fc_sup_pln_cons,
                    'Stock and future supply using actual consumption': rmn_stock_fc_sup_actl_cons,
                    'Stock and actual supply using planned consumption': rmn_stock_actl_sup_pln_cons,
                    'Stock and actual supply using actual consumption': rmn_stock_actl_sup_actl_cons,
                    'Forecast supply': forecast_supply_daily,
                    'Actual supply': actual_supply_daily,
                    'Daily planned consumption': daily_planned_consumption,
                    'Daily actual consumption': daily_actual_consumption,
                    'Day when new contracting should start using future supply and planned consumption': need_new_supply_fc_sup_pln_cons,
                    'Day when new contracting should start using future supply and actual consumption': need_new_supply_fc_sup_actl_cons,
                    'Day when new contracting should start using actual supply and planned consumption': need_new_supply_actl_sup_pln_cons,
                    'Day when new contracting should start using actual supply and actual consumption': need_new_supply_actl_sup_actl_cons
                })

                # Update the stock data for the next day iteration
                stock_fc_sup_pln_cons = rmn_stock_fc_sup_pln_cons
                stock_fc_sup_actl_cons = rmn_stock_fc_sup_actl_cons
                stock_actl_sup_pln_cons = rmn_stock_actl_sup_pln_cons
                stock_actl_sup_actl_cons = rmn_stock_actl_sup_actl_cons

            # Store the first zero stock day signals
            self.zero_stock_day_fc_sup_pln_cons[reactant] = zero_stock_day_fc_sup_pln_cons
            self.zero_stock_day_fc_sup_actl_cons[reactant] = zero_stock_day_fc_sup_actl_cons
            self.zero_stock_day_actl_sup_pln_cons[reactant] = zero_stock_day_actl_sup_pln_cons
            self.zero_stock_day_actl_sup_actl_cons[reactant] = zero_stock_day_actl_sup_actl_cons
        
        self.consumption_summary = pd.DataFrame(data)

    # Fucntion to return data frame resulting all the calculations
    def get_consumption_summary(self):
        return self.consumption_summary

    # Function to return days of zero stock based on different options (planned/actual consumtion, includinf future supply or only actual supply)
    def get_zero_stock_days(self):
        return self.zero_stock_day_fc_sup_pln_cons, self.zero_stock_day_fc_sup_actl_cons, self.zero_stock_day_actl_sup_pln_cons, self.zero_stock_day_actl_sup_actl_cons

    def save_consumption_summary(self, file_path):
        # Save the consumption summary data frame to excel file
        self.consumption_summary.to_excel(file_path, index=False)
        print(f"Reactant consumption data frame saved to {file_path}")

def main():
    # Initialize Tkinter root
    root = Tk()
    root.withdraw()  # Hide the main window
    
    # Ask the user to select the 'stock_data' excel file
    stock_data_path = filedialog.askopenfilename(title='Select the `stock_data` Excel file', filetypes=[('Excel files', '*.xlsx')])
    if not stock_data_path:
        messagebox.showerror('Error', 'No file selected for material data.')
        return
    
    # Ask the user to select the 'supply_data' excel file
    supply_data_path = filedialog.askopenfilename(title='Select the `supply_data` Excel file', filetypes=[('Excel files', '*.xlsx')])
    if not supply_data_path:
        messagebox.showerror('Error', 'No file selected for supply data.')
        return
    
    # Load the data from the provided excel files
    stock_data = pd.read_excel(stock_data_path, index_col='reactant_code')
    supply_data = pd.read_excel(supply_data_path)
    
    # Convert date columns in the provided data frames
    stock_data['stock_update_date'] = pd.to_datetime(stock_data['stock_update_date'])
    supply_data['date_of_delivery_forecast'] = pd.to_datetime(supply_data['date_of_delivery_forecast'])
    supply_data['date_of_delivery_actual'] = pd.to_datetime(supply_data['date_of_delivery_actual'])
    
    # Create an instance of the class with the stock data and supply data
    reactant_consumption = ReactantConsumption(stock_data, supply_data)
    
    # Calculate the daily consumption
    reactant_consumption.calculate_daily_consumption()
    
    # Get the results as a data frame
    reactant_df = reactant_consumption.get_consumption_summary()
    
    # Get the first day when stock equals zero based on different options (planned/actual consumtion, including future supply or only actual supply)
    zero_stock_day_fc_sup_pln_cons, zero_stock_day_fc_sup_actl_cons, zero_stock_day_actl_sup_pln_cons, zero_stock_day_actl_sup_actl_cons = reactant_consumption.get_zero_stock_days()
    
    # Ask the user to specify the save location for the excel file
    save_path = filedialog.asksaveasfilename(title='Save Reactant consumption excel file', defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')])
    if save_path:
        reactant_consumption.save_consumption_summary(save_path)
    else:
        messagebox.showerror('Error', 'No save location selected.')


if __name__ == '__main__':
    main()