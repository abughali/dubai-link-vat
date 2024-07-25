import streamlit as st
import pandas as pd
from xlsxwriter.workbook import Workbook
import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth
from common import utils

        
############################## RAW IMPORTED ###############################################################################
def create_raw_imported(df, workbook):
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    float_format = workbook.add_format({'num_format': '#,##0.00'})

    header_format = workbook.add_format({'bg_color': '#D9E3C0', 'bold': False, 'border': 0})
    total_format = workbook.add_format({'num_format': '#,##0.00','bg_color': '#ffe6e6', 'bold': True})

    worksheet = workbook.add_worksheet('RAW IMPORTED')
    worksheet.set_tab_color('black')
    # Write the column headers
    for col_num, column in enumerate(df.columns):
        worksheet.write(0, col_num, column, header_format)

    # Write data from DataFrame to worksheet
    for row_num, (index, row) in enumerate(df.iterrows(), start=1):
        for col_num, value in enumerate(row):
            if isinstance(value, pd.Timestamp):
               format = date_format
            elif isinstance(value, float):
               format = float_format
            else:
                format = None
            worksheet.write(row_num, col_num, value, format)

    # Calculate the last data row dynamically
    rows = len(df) + 1

    # Add totals for column N
    worksheet.write_formula(rows, 13, f'=SUM(N2:N{rows})', total_format)

    # Add totals for column O
    worksheet.write_formula(rows, 14, f'=SUM(O2:O{rows})', total_format)

    worksheet.autofit()

############################## TOTAL CONVERTED ###############################################################################

def create_total_converted(df, workbook):
    worksheet = workbook.add_worksheet('TOTAL CONVERTED')
    worksheet.set_tab_color('black')

    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    float_format = workbook.add_format({'num_format': '#,##0.00'})

    header_format = workbook.add_format({'bg_color': '#D9E3C0', 'bold': False, 'border': 0})
    total_format = workbook.add_format({'num_format': '#,##0.00','bg_color': '#ffe6e6', 'bold': True})

    suppliers_df = utils.load_rules('suppliers.csv')
    areas_df = utils.load_rules('areas.csv')


    # Flag to indicate if any rows were not found
    supplier_not_found = []
    area_not_found = []

    # Loop through each row in the transactional data DataFrame
    for index, row in df.iterrows():
        supplier = row['Supplier name']
        area = row['Area name']
        
        # Check if the supplier exists in the lookup DataFrame
        if supplier in suppliers_df['Supplier Name'].values:
            # If the supplier is found, get the corresponding 'Service Type' from the lookup DataFrame
            service_type = suppliers_df.loc[suppliers_df['Supplier Name'] == supplier, 'Service Type'].iloc[0]
            
            # Add the 'Service Type' to the transactional DataFrame
            df.at[index, 'Service Type'] = service_type
        else:
            supplier_not_found.append(supplier)

        # Check if the supplier exists in the lookup DataFrame
        if area in areas_df['Area'].values:
            # If the supplier is found, get the corresponding 'Service Type' from the lookup DataFrame
            emirate = areas_df.loc[areas_df['Area'] == area, 'Emirate'].iloc[0]
            
            # Add the 'Service Type' to the transactional DataFrame
            df.at[index, 'Emirate'] = emirate
            df.at[index, 'Country'] = 'UAE'
        else:
            df.at[index, 'Emirate'] = 'NA'
            df.at[index, 'Country'] = 'ROW'
            area_not_found.append(area)

    # Check if any rows were not found and inform the user accordingly
    if supplier_not_found:
        st.error("Undefined supplier(s) found: " + ", ".join(set(supplier_not_found)))
        st.stop()

    if area_not_found:
        st.warning("These areas will be considered ROW: " + ", ".join(map(str, set(area_not_found))))


    df_copy = df.copy() 
                
    column_order = ['Country',
                    'Area name',
                    'Booking code',
                    'No. of nights',
                    'Start date',
                    'End date',
                    'Supplier name',
                    'Description',
                    'Product group',
                    'Product Type',
                    'Service Type',
                    'Final base sales in base currency', 
                    'Final base cost in base currency']

    # Reorder columns in the merged DataFrame
    df = df[column_order]

    # Write the column headers
    for col_num, column in enumerate(df.columns):
        worksheet.write(0, col_num, column, header_format)

    # Write data from DataFrame to worksheet
    for row_num, (index, row) in enumerate(df.iterrows(), start=1):
        for col_num, value in enumerate(row):
            if isinstance(value, pd.Timestamp):
               worksheet.write_datetime(row_num, col_num, value, date_format)
            elif isinstance(value, float):
               worksheet.write_number(row_num, col_num, value, float_format)
            else:
                worksheet.write(row_num, col_num, value)

    # Calculate the last data row dynamically
    rows = len(df) + 1

    # Add totals for column L
    worksheet.write_formula(rows, 11, f'=SUM(L2:L{rows})', total_format)

    # Add totals for column M
    worksheet.write_formula(rows, 12, f'=SUM(M2:M{rows})', total_format)

    (max_row, max_col) = df.shape
    
    worksheet.autofilter(0, 0, max_row, max_col - 1)

    worksheet.autofit()

    return df_copy

############################## HR TAX ###############################################################################

def create_hr_tax_sheet(df, workbook):
    worksheet = workbook.add_worksheet('HR TAX')
    worksheet.set_tab_color('red')

    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    float_format = workbook.add_format({'num_format': '#,##0.00'})
    format_for_zeros = workbook.add_format({'num_format': '-', 'align': 'right'})
    

    header_format = workbook.add_format({'bg_color': '#D9E3C0', 'bold': False, 'border': 0})
    header_format2 = workbook.add_format({'bg_color': '#ffff00', 'bold': False, 'border': 0})
    total_format = workbook.add_format({'num_format': '#,##0.00','bg_color': '#ffe6e6', 'bold': True})
    
    df = df.loc[(df['Country'] == 'UAE') & (df['Service Type'] == 'Hotel Reservation') & (df['Product Type'] != 'Adjustment')]

    suppliers_df = utils.load_rules('suppliers.csv')
    vat_df = utils.load_rules('vat_setup.csv')


    # Loop through each row in the transactional data DataFrame
    for index, row in df.iterrows():
        supplier = row['Supplier name']
        emirate = row['Emirate']
        
        tax_included = suppliers_df.loc[suppliers_df['Supplier Name'] == supplier, 'Taxes Included'].iloc[0]
        bd_amt = vat_df.loc[vat_df['Emirate'] == emirate, 'Basic Division'].iloc[0]
        sc_pct = vat_df.loc[vat_df['Emirate'] == emirate, 'Service Charge'].iloc[0]
        mf_pct = vat_df.loc[vat_df['Emirate'] == emirate, 'Municipality Fee'].iloc[0]
        vat_pct = vat_df.loc[vat_df['Emirate'] == emirate, 'VAT Percentage'].iloc[0]


        if tax_included:
            df.at[index, 'Basic'] = 0
            df.at[index, 'Service Charge'] = 0
            df.at[index, 'Municipality Fee'] = 0
            df.at[index, 'VAT Paid'] = 0
            df.at[index, 'Taxable value input'] = 0
        else:
            df.at[index, 'Basic'] = row['Final base cost in base currency']/bd_amt
            df.at[index, 'Service Charge'] = df.at[index, 'Basic']*(sc_pct/100)
            df.at[index, 'Municipality Fee'] = df.at[index, 'Basic']*(mf_pct/100)
            df.at[index, 'VAT Paid'] = (df.at[index, 'Basic'] + df.at[index, 'Service Charge'])*(vat_pct/100)
            df.at[index, 'Taxable value input'] = df.at[index, 'VAT Paid']/(vat_pct/100)

        df.at[index, 'Total VAT'] = row['Final base sales in base currency']/(100+vat_pct)*vat_pct
        df.at[index, 'Taxable value output'] = df.at[index, 'Total VAT']/(vat_pct/100)
        df.at[index, 'Net VAT payable'] = df.at[index, 'Total VAT']-df.at[index, 'VAT Paid']

    column_order = ['Country',
                    'Emirate',
                    'Area name',
                    'Booking code',
                    'No. of nights',
                    'Start date',
                    'End date',
                    'Supplier name',
                    'Description',
                    'Product group',
                    'Product Type',
                    'Service Type',
                    'Final base sales in base currency', 
                    'Final base cost in base currency',
                    'Basic',
                    'Service Charge',
                    'VAT Paid',
                    'Taxable value input',
                    'Taxable value output',
                    'Total VAT',
                    'Net VAT payable']

    # Reorder columns in the merged DataFrame
    df = df[column_order]
    df.sort_values(by=['Emirate', 'Supplier name'], ascending=[True, True], inplace=True)


    # Define the number of columns to apply the different format
    n_last_columns = 7  # Change this value as needed

    # Write the column headers
    for col_num, column in enumerate(df.columns):
        if col_num >= len(df.columns) - n_last_columns:
            # Apply total_format to the last N columns
            worksheet.write(0, col_num, column, header_format2)
        else:
            # Apply header_format to other columns
            worksheet.write(0, col_num, column, header_format)

    # Write data from DataFrame to worksheet
    for row_num, (index, row) in enumerate(df.iterrows(), start=1):
        for col_num, value in enumerate(row):
            if isinstance(value, pd.Timestamp):
               worksheet.write_datetime(row_num, col_num, value, date_format)
            elif isinstance(value, float):
                if value == 0:  # Check if the value is zero
                    worksheet.write(row_num, col_num, value, format_for_zeros)  # Write dash for zero values
                else:
                    worksheet.write_number(row_num, col_num, value, float_format)
            else:
                worksheet.write(row_num, col_num, value)

    # Calculate the last data row dynamically
    rows = len(df) + 1

    worksheet.write_formula(rows, 12, f'=SUM(M2:M{rows})', total_format)
    worksheet.write_formula(rows, 13, f'=SUM(N2:N{rows})', total_format)
    worksheet.write_formula(rows, 14, f'=SUM(O2:O{rows})', total_format)
    worksheet.write_formula(rows, 15, f'=SUM(P2:P{rows})', total_format)
    worksheet.write_formula(rows, 16, f'=SUM(Q2:Q{rows})', total_format)
    worksheet.write_formula(rows, 17, f'=SUM(R2:R{rows})', total_format)
    worksheet.write_formula(rows, 18, f'=SUM(S2:S{rows})', total_format)
    worksheet.write_formula(rows, 19, f'=SUM(T2:T{rows})', total_format)
    worksheet.write_formula(rows, 20, f'=SUM(U2:U{rows})', total_format)

    (max_row, max_col) = df.shape
    
    worksheet.autofilter(0, 0, max_row, max_col - 1)

    worksheet.autofit()


############################## HR ZERO ###############################################################################

def create_hr_zero_sheet(df, workbook):
    worksheet = workbook.add_worksheet('HR ZERO')
    worksheet.set_tab_color('#3E552A')

    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    float_format = workbook.add_format({'num_format': '#,##0.00'})

    header_format = workbook.add_format({'bg_color': '#D9E3C0', 'bold': False, 'border': 0})
    total_format = workbook.add_format({'num_format': '#,##0.00','bg_color': '#ffe6e6', 'bold': True})
    
    df = df.loc[(df['Country'] == 'ROW') & (df['Service Type'] == 'Hotel Reservation') & (df['Product Type'] != 'Adjustment')]

    column_order = ['Country',
                    'Emirate',
                    'Area name',
                    'Booking code',
                    'No. of nights',
                    'Start date',
                    'End date',
                    'Supplier name',
                    'Description',
                    'Product group',
                    'Product Type',
                    'Service Type',
                    'Final base sales in base currency', 
                    'Final base cost in base currency']

    # Reorder columns in the merged DataFrame
    df = df[column_order]

    # Write the column headers
    for col_num, column in enumerate(df.columns):
        worksheet.write(0, col_num, column, header_format)

    # Write data from DataFrame to worksheet
    for row_num, (index, row) in enumerate(df.iterrows(), start=1):
        for col_num, value in enumerate(row):
            if isinstance(value, pd.Timestamp):
               worksheet.write_datetime(row_num, col_num, value, date_format)
            elif isinstance(value, float):
               worksheet.write_number(row_num, col_num, value, float_format)
            else:
                worksheet.write(row_num, col_num, value)

    # Calculate the last data row dynamically
    rows = len(df) + 1

    # Add totals for column M
    worksheet.write_formula(rows, 12, f'=SUM(M2:M{rows})', total_format)

    # Add totals for column N
    worksheet.write_formula(rows, 13, f'=SUM(N2:N{rows})', total_format)

    (max_row, max_col) = df.shape
    
    worksheet.autofilter(0, 0, max_row, max_col - 1)

    worksheet.autofit()

############################## EX TAX ###############################################################################

def create_ex_tax_sheet(df, workbook):
    worksheet = workbook.add_worksheet('EX TAX')
    worksheet.set_tab_color('red')

    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    float_format = workbook.add_format({'num_format': '#,##0.00'})
    format_for_zeros = workbook.add_format({'num_format': '-', 'align': 'right'})
    

    header_format = workbook.add_format({'bg_color': '#D9E3C0', 'bold': False, 'border': 0})
    header_format2 = workbook.add_format({'bg_color': '#ffff00', 'bold': False, 'border': 0})
    total_format = workbook.add_format({'num_format': '#,##0.00','bg_color': '#ffe6e6', 'bold': True})
    
    df = df.loc[(df['Country'] == 'UAE') & (df['Service Type'] == 'Excursion') & (df['Product Type'] != 'Adjustment')]

    suppliers_df = utils.load_rules('suppliers.csv')
    vat_df = utils.load_rules('vat_setup.csv')


    # Loop through each row in the transactional data DataFrame
    for index, row in df.iterrows():
        supplier = row['Supplier name']
        emirate = row['Emirate']
        
        tax_included = suppliers_df.loc[suppliers_df['Supplier Name'] == supplier, 'Taxes Included'].iloc[0]
        vat_pct = vat_df.loc[vat_df['Emirate'] == emirate, 'VAT Percentage'].iloc[0]


        if tax_included:
            df.at[index, 'VAT Paid'] = 0
            df.at[index, 'Taxable value input'] = 0
        else:
            df.at[index, 'VAT Paid'] = row['Final base cost in base currency']/(100+vat_pct)*vat_pct
            df.at[index, 'Taxable value input'] = df.at[index, 'VAT Paid']/(vat_pct/100)

        df.at[index, 'Profit'] = row['Final base sales in base currency']-row['Final base cost in base currency']
        df.at[index, 'VAT Output'] = row['Final base sales in base currency']/(100+vat_pct)*vat_pct
        df.at[index, 'Taxable value output'] = df.at[index, 'VAT Output']/(vat_pct/100)
        df.at[index, 'Net VAT payable'] = df.at[index, 'VAT Output']-df.at[index, 'VAT Paid']

    column_order = ['Country',
                    'Emirate',
                    'Area name',
                    'Booking code',
                    'No. of nights',
                    'Start date',
                    'End date',
                    'Supplier name',
                    'Description',
                    'Product group',
                    'Product Type',
                    'Service Type',
                    'Final base sales in base currency', 
                    'Final base cost in base currency',
                    'Profit',
                    'VAT Paid',
                    'Taxable value input',
                    'VAT Output',
                    'Taxable value output',
                    'Net VAT payable']

    # Reorder columns in the merged DataFrame
    df = df[column_order]
    df.rename(columns={'Final base sales in base currency': 'Final Sale', 'Final base cost in base currency': 'Final Cost'}, inplace=True)

    df.sort_values(by=['Emirate', 'Supplier name'], ascending=[True, True], inplace=True)


    # Define the number of columns to apply the different format
    n_last_columns = 6  # Change this value as needed

    # Write the column headers
    for col_num, column in enumerate(df.columns):
        if col_num >= len(df.columns) - n_last_columns:
            # Apply total_format to the last N columns
            worksheet.write(0, col_num, column, header_format2)
        else:
            # Apply header_format to other columns
            worksheet.write(0, col_num, column, header_format)

    # Write data from DataFrame to worksheet
    for row_num, (index, row) in enumerate(df.iterrows(), start=1):
        for col_num, value in enumerate(row):
            if isinstance(value, pd.Timestamp):
               worksheet.write_datetime(row_num, col_num, value, date_format)
            elif isinstance(value, float):
                if value == 0:  # Check if the value is zero
                    worksheet.write(row_num, col_num, value, format_for_zeros)  # Write dash for zero values
                else:
                    worksheet.write_number(row_num, col_num, value, float_format)
            else:
                worksheet.write(row_num, col_num, value)

    # Calculate the last data row dynamically
    rows = len(df) + 1

    worksheet.write_formula(rows, 12, f'=SUM(M2:M{rows})', total_format)
    worksheet.write_formula(rows, 13, f'=SUM(N2:N{rows})', total_format)
    worksheet.write_formula(rows, 14, f'=SUM(O2:O{rows})', total_format)
    worksheet.write_formula(rows, 15, f'=SUM(P2:P{rows})', total_format)
    worksheet.write_formula(rows, 16, f'=SUM(Q2:Q{rows})', total_format)
    worksheet.write_formula(rows, 17, f'=SUM(R2:R{rows})', total_format)
    worksheet.write_formula(rows, 18, f'=SUM(S2:S{rows})', total_format)
    worksheet.write_formula(rows, 19, f'=SUM(T2:T{rows})', total_format)

    (max_row, max_col) = df.shape
    
    worksheet.autofilter(0, 0, max_row, max_col - 1)

    worksheet.autofit()

############################## EX ZERO ###############################################################################

def create_excursion_zero_sheet(df, workbook):
    worksheet = workbook.add_worksheet('EX ZERO')
    worksheet.set_tab_color('#3E552A')

    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    float_format = workbook.add_format({'num_format': '#,##0.00'})

    header_format = workbook.add_format({'bg_color': '#D9E3C0', 'bold': False, 'border': 0})
    total_format = workbook.add_format({'num_format': '#,##0.00','bg_color': '#ffe6e6', 'bold': True})
    
    df = df.loc[(df['Country'] == 'ROW') & (df['Service Type'] == 'Excursion') & (df['Product Type'] != 'Adjustment')]

    column_order = ['Country',
                    'Emirate',
                    'Area name',
                    'Booking code',
                    'No. of nights',
                    'Start date',
                    'End date',
                    'Supplier name',
                    'Description',
                    'Product group',
                    'Product Type',
                    'Service Type',
                    'Final base sales in base currency', 
                    'Final base cost in base currency']

    # Reorder columns in the merged DataFrame
    df = df[column_order]

    # Write the column headers
    for col_num, column in enumerate(df.columns):
        worksheet.write(0, col_num, column, header_format)

    # Write data from DataFrame to worksheet
    for row_num, (index, row) in enumerate(df.iterrows(), start=1):
        for col_num, value in enumerate(row):
            if isinstance(value, pd.Timestamp):
               worksheet.write_datetime(row_num, col_num, value, date_format)
            elif isinstance(value, float):
               worksheet.write_number(row_num, col_num, value, float_format)
            else:
                worksheet.write(row_num, col_num, value)

    # Calculate the last data row dynamically
    rows = len(df) + 1

    # Add totals for column M
    worksheet.write_formula(rows, 12, f'=SUM(M2:M{rows})', total_format)

    # Add totals for column N
    worksheet.write_formula(rows, 13, f'=SUM(N2:N{rows})', total_format)

    (max_row, max_col) = df.shape
    
    worksheet.autofilter(0, 0, max_row, max_col - 1)

    worksheet.autofit()

############################## AIR TICKET ###############################################################################

def create_air_ticket_sheet(df, workbook):
    worksheet = workbook.add_worksheet('AIR TICKET')
    worksheet.set_tab_color('#6A9AD0')

    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    float_format = workbook.add_format({'num_format': '#,##0.00'})

    header_format = workbook.add_format({'bg_color': '#D9E3C0', 'bold': False, 'border': 0})
    total_format = workbook.add_format({'num_format': '#,##0.00','bg_color': '#ffe6e6', 'bold': True})
    
    df = df.loc[df['Service Type'] == 'Air Ticket']

    column_order = ['Country',
                    'Emirate',
                    'Area name',
                    'Booking code',
                    'No. of nights',
                    'Start date',
                    'End date',
                    'Supplier name',
                    'Description',
                    'Product group',
                    'Product Type',
                    'Service Type',
                    'Final base sales in base currency', 
                    'Final base cost in base currency']

    # Reorder columns in the merged DataFrame
    df = df[column_order]

    # Write the column headers
    for col_num, column in enumerate(df.columns):
        worksheet.write(0, col_num, column, header_format)

    # Write data from DataFrame to worksheet
    for row_num, (index, row) in enumerate(df.iterrows(), start=1):
        for col_num, value in enumerate(row):
            if isinstance(value, pd.Timestamp):
               worksheet.write_datetime(row_num, col_num, value, date_format)
            elif isinstance(value, float):
               worksheet.write_number(row_num, col_num, value, float_format)
            else:
                worksheet.write(row_num, col_num, value)

    # Calculate the last data row dynamically
    rows = len(df) + 1

    # Add totals for column M
    worksheet.write_formula(rows, 12, f'=SUM(M2:M{rows})', total_format)

    # Add totals for column N
    worksheet.write_formula(rows, 13, f'=SUM(N2:N{rows})', total_format)

    (max_row, max_col) = df.shape
    
    worksheet.autofilter(0, 0, max_row, max_col - 1)

    worksheet.autofit()

############################## VISA ###############################################################################

def create_visa_sheet(df, workbook):
    worksheet = workbook.add_worksheet('VISA')
    worksheet.set_tab_color('red')

    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    float_format = workbook.add_format({'num_format': '#,##0.00'})
    format_for_zeros = workbook.add_format({'num_format': '-', 'align': 'right'})
    

    header_format = workbook.add_format({'bg_color': '#D9E3C0', 'bold': False, 'border': 0})
    header_format2 = workbook.add_format({'bg_color': '#ffff00', 'bold': False, 'border': 0})
    total_format = workbook.add_format({'num_format': '#,##0.00','bg_color': '#ffe6e6', 'bold': True})
    
    df = df.loc[(df['Country'] == 'UAE') & (df['Service Type'] == 'Visa') & (df['Product Type'] != 'Adjustment')]

    suppliers_df = utils.load_rules('suppliers.csv')
    vat_df = utils.load_rules('vat_setup.csv')


    # Loop through each row in the transactional data DataFrame
    for index, row in df.iterrows():
        supplier = row['Supplier name']
        emirate = row['Emirate']
        
        tax_included = suppliers_df.loc[suppliers_df['Supplier Name'] == supplier, 'Taxes Included'].iloc[0]
        vat_pct = vat_df.loc[vat_df['Emirate'] == emirate, 'VAT Percentage'].iloc[0]


        df.at[index, 'Basic Charges'] = 0
        df.at[index, 'Service Charges'] = 0
        df.at[index, 'Naqoodi Charges'] = 0
        df.at[index, 'VAT Paid'] = 0
        df.at[index, 'Reconciled'] = 0
        df.at[index, 'Taxable Value Input'] = 0
        df.at[index, 'VAT Output'] = 0
        df.at[index, 'Taxable Value Output'] = 0
        df.at[index, 'VAT payable'] = df.at[index, 'VAT Output']-df.at[index, 'VAT Paid']


    column_order = ['Country',
                    'Emirate',
                    'Area name',
                    'Booking code',
                    'No. of nights',
                    'Start date',
                    'End date',
                    'Supplier name',
                    'Description',
                    'Product group',
                    'Product Type',
                    'Service Type',
                    'Final base sales in base currency', 
                    'Final base cost in base currency',
                    'Basic Charges',
                    'Service Charges',
                    'Naqoodi Charges',
                    'VAT Paid',
                    'Reconciled',
                    'Taxable Value Input',
                    'VAT Output',
                    'Taxable Value Output',
                    'VAT payable']

    # Reorder columns in the merged DataFrame
    df = df[column_order]

    df.sort_values(by=['Emirate', 'Supplier name'], ascending=[True, True], inplace=True)


    # Define the number of columns to apply the different format
    n_last_columns = 9  # Change this value as needed

    # Write the column headers
    for col_num, column in enumerate(df.columns):
        if col_num >= len(df.columns) - n_last_columns:
            # Apply total_format to the last N columns
            worksheet.write(0, col_num, column, header_format2)
        else:
            # Apply header_format to other columns
            worksheet.write(0, col_num, column, header_format)

    # Write data from DataFrame to worksheet
    for row_num, (index, row) in enumerate(df.iterrows(), start=1):
        for col_num, value in enumerate(row):
            if isinstance(value, pd.Timestamp):
               worksheet.write_datetime(row_num, col_num, value, date_format)
            elif isinstance(value, float):
                if value == 0:  # Check if the value is zero
                    worksheet.write(row_num, col_num, value, format_for_zeros)  # Write dash for zero values
                else:
                    worksheet.write_number(row_num, col_num, value, float_format)
            else:
                worksheet.write(row_num, col_num, value)

    # Calculate the last data row dynamically
    rows = len(df) + 1

    worksheet.write_formula(rows, 12, f'=SUM(M2:M{rows})', total_format)
    worksheet.write_formula(rows, 13, f'=SUM(N2:N{rows})', total_format)
    worksheet.write_formula(rows, 14, f'=SUM(O2:O{rows})', total_format)
    worksheet.write_formula(rows, 15, f'=SUM(P2:P{rows})', total_format)
    worksheet.write_formula(rows, 16, f'=SUM(Q2:Q{rows})', total_format)
    worksheet.write_formula(rows, 17, f'=SUM(R2:R{rows})', total_format)
    worksheet.write_formula(rows, 18, f'=SUM(S2:S{rows})', total_format)
    worksheet.write_formula(rows, 19, f'=SUM(T2:T{rows})', total_format)
    worksheet.write_formula(rows, 20, f'=SUM(U2:U{rows})', total_format)
    worksheet.write_formula(rows, 21, f'=SUM(V2:V{rows})', total_format)
    worksheet.write_formula(rows, 22, f'=SUM(W2:W{rows})', total_format)


    (max_row, max_col) = df.shape
    
    worksheet.autofilter(0, 0, max_row, max_col - 1)

    worksheet.autofit()

############################## OTHERS ###############################################################################

def create_others_sheet(df, workbook):
    worksheet = workbook.add_worksheet('OTHER NA')
    worksheet.set_tab_color('#475468')

    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    float_format = workbook.add_format({'num_format': '#,##0.00'})

    header_format = workbook.add_format({'bg_color': '#D9E3C0', 'bold': False, 'border': 0})
    total_format = workbook.add_format({'num_format': '#,##0.00','bg_color': '#ffe6e6', 'bold': True})
    
    df = df.loc[df['Service Type'] == 'Other']

    column_order = ['Country',
                    'Emirate',
                    'Area name',
                    'Booking code',
                    'No. of nights',
                    'Start date',
                    'End date',
                    'Supplier name',
                    'Description',
                    'Product group',
                    'Product Type',
                    'Service Type',
                    'Final base sales in base currency', 
                    'Final base cost in base currency']

    # Reorder columns in the merged DataFrame
    df = df[column_order]

    # Write the column headers
    for col_num, column in enumerate(df.columns):
        worksheet.write(0, col_num, column, header_format)

    # Write data from DataFrame to worksheet
    for row_num, (index, row) in enumerate(df.iterrows(), start=1):
        for col_num, value in enumerate(row):
            if isinstance(value, pd.Timestamp):
               worksheet.write_datetime(row_num, col_num, value, date_format)
            elif isinstance(value, float):
               worksheet.write_number(row_num, col_num, value, float_format)
            else:
                worksheet.write(row_num, col_num, value)

    # Calculate the last data row dynamically
    rows = len(df) + 1

    # Add totals for column M
    worksheet.write_formula(rows, 12, f'=SUM(M2:M{rows})', total_format)

    # Add totals for column N
    worksheet.write_formula(rows, 13, f'=SUM(N2:N{rows})', total_format)

    (max_row, max_col) = df.shape
    
    worksheet.autofilter(0, 0, max_row, max_col - 1)

    worksheet.autofit()

############################## GENERATE REPORT ###############################################################################


def generate_report(df):
    report_name = 'processed_data.xlsx'
    workbook = Workbook(report_name, {'nan_inf_to_errors': True, 'default_date_format': 'yyyy-mm-dd'})
    create_raw_imported(df, workbook)
    df_all = create_total_converted(df, workbook)
    create_hr_tax_sheet(df_all, workbook)
    create_hr_zero_sheet(df_all, workbook)
    create_ex_tax_sheet(df_all, workbook)    
    create_excursion_zero_sheet(df_all, workbook)
    create_air_ticket_sheet(df_all, workbook)
    create_visa_sheet(df_all, workbook)
    create_others_sheet(df_all, workbook)
    workbook.close()
    return report_name

def add_suffix(number):
    if 10 <= number % 100 <= 20:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(number % 10, 'th')
    return str(number) + suffix

def main_app():
    pd.options.mode.copy_on_write = True
    title = 'VAT Report Generator'
    st.markdown(f"<h1 style='font-size:24px;'>{title}</h1>", unsafe_allow_html=True)

    file = st.file_uploader("Upload an Excel file", type=['xlsx'])
    if file:
        df = pd.read_excel(file)
        q_name = pd.to_datetime(max(df['Start date'])) + pd.tseries.offsets.QuarterEnd(0)
        quarter = q_name.quarter
        formatted_date = q_name.strftime("%d %b %Y")
        report_name = 'VAT ' + add_suffix(quarter) + ' QTR ' + formatted_date.upper() + '.xlsx'
        if st.button('Generate Report'): 
            excel_path = generate_report(df)
            with open(excel_path, "rb") as file:
                st.download_button(label=f'📥 Download {report_name} Report', data=file, file_name=report_name, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


utils.hide_home_page()
utils.add_logo()

with open('./common/config.yaml') as file:
    config = yaml.load(file, Loader=SafeLoader)

authenticator = stauth.Authenticate(
    config['credentials'],
    config['cookie']['name'],
    config['cookie']['key'],
    config['cookie']['expiry_days'],
    config['pre-authorized']
)

authenticator.login(max_login_attempts=5, clear_on_submit=True)

if st.session_state["authentication_status"]:
    authenticator.logout(location='sidebar')
    main_app()
elif st.session_state["authentication_status"] is False:
    st.error('Username/password is incorrect')

