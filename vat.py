import streamlit as st
import pandas as pd
from io import BytesIO
from st_aggrid import AgGrid, GridOptionsBuilder
from streamlit_option_menu import option_menu
from xlsxwriter.workbook import Workbook
import requests
import xml.etree.ElementTree as ET
from datetime import datetime

def load_rules(filename):
    try:
        return pd.read_csv(filename)
    except FileNotFoundError:
        if 'suppliers.csv' in filename:
            return pd.DataFrame(columns=['Supplier Name', 'Service Type', 'Taxes Included'])
        elif 'areas.csv' in filename:
            return pd.DataFrame(columns=['Area', 'Emirate'])
        elif 'services.csv' in filename:
            return pd.DataFrame(columns=['Service Type', 'VAT Exempt'])
        elif 'vat_setup.csv' in filename:
            return pd.DataFrame(columns=['Emirate', 'Basic Division', 'Service Charge', 'Municipality Fee', 'VAT Percentage'])
        else:
            return pd.DataFrame()  # Default empty DataFrame if no specific columns match

def save_rules(rules_df, filename):
    rules_df.to_csv(filename, index=False)

def load_service_types(filename):
    try:
        df = pd.read_csv(filename)
        return df['Service Type'].dropna().unique().tolist()
    except FileNotFoundError:
        return []

def load_emirates(filename):
    try:
        df = pd.read_csv(filename)
        return df['Emirate'].dropna().unique().tolist()
    except FileNotFoundError:
        return []

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

def area_state_country_editor(filename):
    title = "Areas Mapping"
    st.markdown(f"<h3 style='font-size:24px;'>{title}</h3>", unsafe_allow_html=True)
    rules_df = load_rules(filename)
    if rules_df.empty:
        rules_df = pd.DataFrame(columns=['Area', 'Emirate'])

    # Ag-Grid implementation
    grid_options = GridOptionsBuilder.from_dataframe(rules_df)
    grid_options.configure_side_bar()
    grid_options.configure_selection('single')
    AgGrid(rules_df, grid_options=grid_options.build(), 
           update_mode='MODEL_CHANGED', 
           fit_columns_on_grid_load=True, 
           enable_enterprise_modules=False,
           pagination=True,
           paginationPageSize=10)

    emirates = load_emirates('vat_setup.csv')

    # Add or update rule
    with st.form(key='form_area_state_country', clear_on_submit=True):
        new_area = st.text_input("Area")
        new_state = st.selectbox("Emirate", options=emirates)
        submitted = st.form_submit_button("💾 Save")
        if submitted:
            new_row = pd.DataFrame([{'Area': new_area, 'Emirate': new_state}])
            rules_df = pd.concat([rules_df, new_row]).drop_duplicates(subset=['Area', 'Emirate'], keep='last')
            save_rules(rules_df, filename)
            st.success("Saved!")
            st.rerun()



def product_supplier_editor(filename):
    title = "Product and Supplier Mapping"
    st.markdown(f"<h3 style='font-size:24px;'>{title}</h3>", unsafe_allow_html=True)
    rules_df = load_rules(filename)
    if rules_df.empty:
        rules_df = pd.DataFrame(columns=['Supplier Name', 'Service Type', 'Taxes Included'])

    # Load service types for dropdown
    service_types = load_service_types('services.csv')

    # Ag-Grid implementation
    grid_options = GridOptionsBuilder.from_dataframe(rules_df)
    grid_options.configure_pagination(paginationAutoPageSize=True)
    grid_options.configure_side_bar()
    grid_options.configure_selection('single')
    AgGrid(rules_df, grid_options=grid_options.build(), 
           update_mode='MODEL_CHANGED', 
           fit_columns_on_grid_load=True, 
           enable_enterprise_modules=False,
           pagination=True,
           paginationPageSize=10)
    
    # Add or update rule
    with st.form(key='form_product_supplier', clear_on_submit=True):
        new_supplier = st.text_input("Supplier Name")
        new_service_type = st.selectbox("Service Type", options=service_types)
        new_tax = st.checkbox("Taxes Included")
    
        submitted = st.form_submit_button("💾 Save")
        if submitted:
            new_row = pd.DataFrame([{'Supplier Name': new_supplier, 'Service Type': new_service_type, 'Taxes Included': new_tax}])
            rules_df = pd.concat([rules_df, new_row]).drop_duplicates(subset=['Supplier'], keep='last')
            save_rules(rules_df, filename)
            st.success("Saved!")
            st.rerun()


def service_type_editor(filename):
    title = "Service Type Management"
    st.markdown(f"<h3 style='font-size:24px;'>{title}</h3>", unsafe_allow_html=True)

    rules_df = load_rules(filename)
    if rules_df.empty:
        rules_df = pd.DataFrame(columns=['Service Type', 'VAT Exempt'])

    # Ag-Grid implementation
    grid_options = GridOptionsBuilder.from_dataframe(rules_df)
    grid_options.configure_pagination(paginationAutoPageSize=True)
    grid_options.configure_side_bar()
    grid_options.configure_selection('single')
    AgGrid(rules_df, grid_options=grid_options.build(), update_mode='MODEL_CHANGED', fit_columns_on_grid_load=True)

    # Add or update rule
    with st.form(key='form_service_type', clear_on_submit=True):
        new_service_type = st.text_input("Service Type")
        new_vat_exempt = st.checkbox("VAT Exempt")
        submitted = st.form_submit_button("💾 Save")
        if submitted:
            new_row = pd.DataFrame([{'Service Type': new_service_type, 'VAT Exempt': new_vat_exempt}])
            rules_df = pd.concat([rules_df, new_row]).drop_duplicates(subset=['Service Type'], keep='last')
            save_rules(rules_df, filename)
            st.success("Saved!")
            st.rerun()

def vat_setup_editor(filename):
    title = "VAT Setup"
    st.markdown(f"<h3 style='font-size:24px;'>{title}</h3>", unsafe_allow_html=True)
    rules_df = load_rules(filename)
    if rules_df.empty:
        rules_df = pd.DataFrame(columns=['Emirate', 'Basic Division', 'Service Charge', 'Municipality Fee', 'VAT Percentage'])

    # Ag-Grid implementation
    grid_options = GridOptionsBuilder.from_dataframe(rules_df)
    grid_options.configure_pagination(paginationAutoPageSize=True)
    grid_options.configure_side_bar()
    grid_options.configure_selection('single')
    AgGrid(rules_df, grid_options=grid_options.build(), update_mode='MODEL_CHANGED', fit_columns_on_grid_load=True)


    # Add or update rule
    with st.form(key='form_vat_setup', clear_on_submit=True):
        new_emirate = st.text_input("Emirate")
        new_bd = st.number_input("Basic Division", format="%.3f")
        new_sc = st.number_input("Service Charge", format="%.2f")
        new_mf = st.number_input("Municipality Fee", format="%.2f")
        new_vat = st.number_input("VAT Percentage", format="%.2f")

        submitted = st.form_submit_button("💾 Save")
        if submitted:
            new_row = pd.DataFrame([{'Emirate': new_emirate, 'Basic Division': new_bd, 'Service Charge': new_sc, 'Municipality Fee': new_mf, 'VAT Percentage': new_vat}])
            rules_df = pd.concat([rules_df, new_row]).drop_duplicates(subset=['Emirate'], keep='last')
            save_rules(rules_df, filename)
            st.success("Saved!")
            st.rerun()

       
 
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

    # Add totals for column J
    worksheet.write_formula(rows, 9, f'=SUM(J2:J{rows})', total_format)

    # Add totals for column K
    worksheet.write_formula(rows, 10, f'=SUM(K2:K{rows})', total_format)

    worksheet.autofit()

############################## TOTAL CONVERTED ###############################################################################

def create_total_converted(df, workbook):
    worksheet = workbook.add_worksheet('TOTAL CONVERTED')
    worksheet.set_tab_color('black')

    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    float_format = workbook.add_format({'num_format': '#,##0.00'})

    header_format = workbook.add_format({'bg_color': '#D9E3C0', 'bold': False, 'border': 0})
    total_format = workbook.add_format({'num_format': '#,##0.00','bg_color': '#ffe6e6', 'bold': True})

    suppliers_df = load_rules('suppliers.csv')
    areas_df = load_rules('areas.csv')


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
        st.warning("These areas will be considered ROW: " + ", ".join(set(area_not_found)))

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

    suppliers_df = load_rules('suppliers.csv')
    vat_df = load_rules('vat_setup.csv')


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

    suppliers_df = load_rules('suppliers.csv')
    vat_df = load_rules('vat_setup.csv')


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

    suppliers_df = load_rules('suppliers.csv')
    vat_df = load_rules('vat_setup.csv')


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
#########################################################################################################
supplierList = []

def fetch_and_populate_suppliers():
    """Fetch suppliers from the API and populate the supplier list."""
    global supplierList
    
    # URL and headers
    url = "https://www.gte.travel/wsExportacion/wssuppliers.asmx/getSupplierList"
    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }
    
    # Data payload
    data = {
        "user": st.secrets["gte_user"],
        "password": st.secrets["gte_password"],   
        "SupplierId": "",
        "ExportMode": "",
        "creationDateFrom": "",
        "creationDateTo": ""
    }
    
    # Category mapping
    category_mapping = {
        "1": "Cross Sell Hotel",
        "2": "Dynamic Hotel",
        "3": "Excursions",
        "4": "Extranet Hotel",
        "5": "Static Hotel",
        "6": "Tickets",
        "7": "Visa",
        "8": "XML Hotel",
        "9": "Offline Hotel"
    }
    
    response = requests.post(url, headers=headers, data=data)

    if response.status_code == 200:        
        # Parse the XML response
        root = ET.fromstring(response.text)
        
        # Extract supplier information
        for supplier in root.findall(".//Supplier"):
            category_element = supplier.find(".//Category")
            category_id = category_element.get('Id') if category_element is not None else None
            category_name = category_mapping.get(category_id, "Others")
            
            supplier_data = {
                "Supplier Id": supplier.get("Id"),
                "Category Id": category_id,
                "Category Name": category_name
            }
            supplierList.append(supplier_data)
        
    else:
        st.error(f"Failed to fetch suppliers. Status code: {response.status_code}")

def get_category_name(supplier_id):
    for supplier in supplierList:
        if supplier["Supplier Id"] == supplier_id:
            return supplier["Category Name"]
    return "Supplier ID not found"

# Function to remove time from datetime string
def format_date(date_str):
    if date_str:
        return datetime.strptime(date_str, "%Y-%m-%dT%H:%M:%S").strftime("%Y-%m-%d")
    return ""

def fetch_invoices(invoice_date_from, invoice_date_to):
    # URL and headers
    url = "https://www.gte.travel/wsExportacion/wsinvoices.asmx/GetInvoices"
    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }

    # Data payload
    data = {
        "user": st.secrets["gte_user"],
        "password": st.secrets["gte_password"],      
        "InvoiceSeries": "",
        "InvoiceNumberFrom": "",
        "InvoiceNumberTo": "",
        "InvoiceDateFrom": invoice_date_from,
        "InvoiceDateTo": invoice_date_to,
        "InvoiceIdNumberFrom": "",
        "InvoiceIdNumberTo": "",
        "BeginTravelDate": "",
        "EndTravelDate": "",
        "customerId": "",
        "ExportMode": "",
        "channel": "",
        "IncludeRelatedInvoice": "",
        "locator": ""
    }

    # Make the POST request
    response = requests.post(url, headers=headers, data=data)

    if response.status_code == 200:
        root = ET.fromstring(response.text)

        invoice_count = 0
        invoice_item_count = 0
        invoices = []

        for invoice in root.findall(".//Invoice"):
            invoice_count += 1
            invoice_number = invoice.get("InvoiceNumber")
            invoice_date = format_date(invoice.get("InvoiceDate"))
            due_date = format_date(invoice.get("DueDate"))
            currency = invoice.get("Currency")
            customer_name = invoice.find(".//CustomerName").text if invoice.find(".//CustomerName") is not None else ""
            operation_rate_elem = invoice.find(".//OperationRate")
            exchange_rate = float(operation_rate_elem.text) if operation_rate_elem is not None and operation_rate_elem.text else 1.0

            for line in invoice.findall(".//Line"):
                invoice_item_count += 1
                service = line.find(".//Service").text if line.find(".//Service") is not None else ""
                pax_name = invoice.find(".//Passenger/name").text if invoice.find(".//Passenger/name") is not None else ""
                pax_surname = invoice.find(".//Passenger/surname").text if invoice.find(".//Passenger/surname") is not None else ""
                begin_travel_date = format_date(line.get("BeginTravelDate"))
                end_travel_date = format_date(line.get("EndTravelDate"))
                service_date = begin_travel_date
                # Extract SupplierId from Cost element
                cost_elem = line.find(".//Cost")
                supplier_id = cost_elem.get("SupplierId") if cost_elem is not None else ""
                
                # Conditionally include passenger name if available
                if pax_name and pax_surname:
                    item_description = f"Name :- {pax_name} {pax_surname}\n{service}\nTravel Date {begin_travel_date} - {end_travel_date}"
                else:
                    item_description = f"{service}\nTravel Date {begin_travel_date} - {end_travel_date}"

                # Convert amounts to AED if not already in AED
                if currency != "AED":
                    item_amount = round(float(line.get("NetLineAmount")) * exchange_rate, 2)
                    taxes = round(float(line.get("Taxes")) * exchange_rate, 2)
                else:
                    item_amount = float(line.get("NetLineAmount"))
                    taxes = float(line.get("Taxes"))

                line_data = {
                    "Invoice No": invoice_number,
                    "InvoiceDate": invoice_date,
                    "DueDate": due_date,
                    "Service Date": service_date,
                    "Currency": "AED",
                    "CustomerName": customer_name,
                    "Memo": line.get("BookingCode"),
                    "Item Amount": item_amount,
                    "Taxes": taxes,
                    "Item Description": item_description,
                    "Tax Code": "5% VAT" if taxes > 0 else "EX Exempt",
                    "Service": get_category_name(supplier_id)
                }


                invoices.append(line_data)
        
        # Create a DataFrame
        df = pd.DataFrame(invoices)
        
        return invoice_count, invoice_item_count, df
    else:
        st.error(f"Failed to fetch invoices. Status code: {response.status_code}")
        return 0, 0, pd.DataFrame()

def save_csv_files(df, start_date_str, end_date_str):
    # Exclude rows with negative amounts
    df_positive = df[df["Item Amount"] >= 0]
    
    # Group by invoice number
    grouped = df_positive.groupby("Invoice No")
    
    # Create chunks
    chunk_size = 1000
    current_chunk = []
    chunks = []
    current_size = 0
    
    for name, group in grouped:
        group_size = len(group)
        if current_size + group_size > chunk_size:
            chunks.append(pd.concat(current_chunk))
            current_chunk = []
            current_size = 0
        current_chunk.append(group)
        current_size += group_size
    
    if current_chunk:
        chunks.append(pd.concat(current_chunk))
    
    csv_files = []
    for idx, chunk in enumerate(chunks):
        file_name = f'invoices_{start_date_str}_{end_date_str}_part{idx+1}.csv'
        chunk_csv = chunk.to_csv(index=False)
        csv_files.append((file_name, chunk_csv))
    
    return csv_files

def save_credit_memo_files(df, start_date_str, end_date_str):
    credit_memo_df = df[df["Item Amount"] < 0]
    if not credit_memo_df.empty:
        credit_memo_file_name = f'credit_memo_{start_date_str}_{end_date_str}.csv'
        credit_memo_csv = credit_memo_df.to_csv(index=False)
        return credit_memo_file_name, credit_memo_csv
    return None, None

def qb_invoices():
    pd.options.mode.copy_on_write = True
    title = 'Juniper Invoices Generator'
    st.markdown(f"<h1 style='font-size:24px;'>{title}</h1>", unsafe_allow_html=True)

    invoice_date_from = st.date_input("Invoice Date From")
    invoice_date_to = st.date_input("Invoice Date To")

    if "csv_files" not in st.session_state:
        st.session_state.csv_files = []
    if "credit_memo_file" not in st.session_state:
        st.session_state.credit_memo_file = None
    if "invoice_count" not in st.session_state:
        st.session_state.invoice_count = 0
    if "invoice_item_count" not in st.session_state:
        st.session_state.invoice_item_count = 0
    if "df" not in st.session_state:
        st.session_state.df = pd.DataFrame()

    # Validation checks
    if invoice_date_from and invoice_date_to:
        same_month = invoice_date_from.month == invoice_date_to.month
        if not same_month:
            st.error("The period must be within the same month.")
        else:
            if st.button('Fetch Invoices'):
                with st.spinner('Fetching invoices...'):
                    invoice_date_from_str = invoice_date_from.strftime("%Y%m%d")
                    invoice_date_to_str = invoice_date_to.strftime("%Y%m%d")
                    fetch_and_populate_suppliers()
                    invoice_count, invoice_item_count, df = fetch_invoices(invoice_date_from_str, invoice_date_to_str)

                    if not df.empty:
                        df.index = range(1, len(df) + 1)
                        st.session_state.csv_files = save_csv_files(df, invoice_date_from_str, invoice_date_to_str)
                        st.session_state.invoice_count = invoice_count
                        st.session_state.invoice_item_count = invoice_item_count
                        st.session_state.df = df

                        # Save credit memo files
                        credit_memo_file_name, credit_memo_csv = save_credit_memo_files(df, invoice_date_from_str, invoice_date_to_str)
                        if credit_memo_csv:
                            st.session_state.credit_memo_file = (credit_memo_file_name, credit_memo_csv)
                        else:
                            st.session_state.credit_memo_file = None

    if st.session_state.invoice_count:
        st.write(f"Number of invoices: {st.session_state.invoice_count}")
    if st.session_state.invoice_item_count:
        st.write(f"Number of invoice items: {st.session_state.invoice_item_count}")
    if not st.session_state.df.empty:
        st.write(st.session_state.df)
    
    if st.session_state.csv_files:
        for file_name, csv in st.session_state.csv_files:
            st.download_button(
                label=f"📥 Download {file_name}",
                data=csv,
                file_name=file_name,
                mime='text/csv',
            )
    
    if st.session_state.credit_memo_file:
        file_name, csv = st.session_state.credit_memo_file
        st.download_button(
            label=f"📥 Download {file_name}",
            data=csv,
            file_name=file_name,
            mime='text/csv',
        )
##################################################################################################################################

def fetch_bills(invoice_date_from, invoice_date_to):
    # URL and headers
    url = "https://www.gte.travel/wsExportacion/wsinvoices.asmx/GetInvoices"
    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }

    # Data payload
    data = {
        "user": st.secrets["gte_user"],
        "password": st.secrets["gte_password"],      
        "InvoiceSeries": "",
        "InvoiceNumberFrom": "",
        "InvoiceNumberTo": "",
        "InvoiceDateFrom": invoice_date_from,
        "InvoiceDateTo": invoice_date_to,
        "InvoiceIdNumberFrom": "",
        "InvoiceIdNumberTo": "",
        "BeginTravelDate": "",
        "EndTravelDate": "",
        "customerId": "",
        "ExportMode": "",
        "channel": "",
        "IncludeRelatedInvoice": "",
        "locator": ""
    }

    # Make the POST request
    response = requests.post(url, headers=headers, data=data)

    if response.status_code == 200:
        root = ET.fromstring(response.text)

        invoice_count = 0
        invoice_item_count = 0
        invoices = []

        for invoice in root.findall(".//Invoice"):
            invoice_count += 1
            invoice_number = invoice.get("InvoiceNumber")
            invoice_date = format_date(invoice.get("InvoiceDate"))
            due_date = format_date(invoice.get("DueDate"))
            #customer_name = invoice.find(".//CustomerName").text if invoice.find(".//CustomerName") is not None else ""
            tax_elem = invoice.find(".//tax")
            tax_value = float(tax_elem.get("value")) if tax_elem is not None and tax_elem.get("value") else 0.0

            for line in invoice.findall(".//Line"):
                invoice_item_count += 1
                # Extract SupplierId from Cost element
                cost_elem = line.find(".//Cost")
                supplier_id = cost_elem.get("SupplierId") if cost_elem is not None else ""
                supplier_name = line.find(".//SupplierName").text
                service = line.find(".//ArticleOfCost").text
                begin_travel_date = format_date(line.get("BeginTravelDate"))
                end_travel_date = format_date(line.get("EndTravelDate"))
                # Extract SupplierId from Cost element
                cost_elem = line.find(".//Cost")
                supplier_id = cost_elem.get("SupplierId") if cost_elem is not None else ""
                
                item_description = f"{service}\nTravel Date {begin_travel_date} - {end_travel_date}"
                exchange_rate = float(cost_elem.get("ExchangeRate"))
                currency = cost_elem.get("Currency")


                # Convert amounts to AED if not already in AED
                if currency != "AED":
                    item_amount = round(float(cost_elem.get("TotalAmount")) / exchange_rate, 2)
                else:
                    item_amount = float(cost_elem.get("TotalAmount"))

                taxes = round((tax_value / 100) * item_amount,2)

                line_data = {
                    "Bill No": invoice_number,
                    "Bill Date": invoice_date,
                    "DueDate": due_date,
                    "Currency": "AED",
                    "Supplier": supplier_name,
                    "Memo": line.get("BookingCode"),
                    "Line Amount": item_amount,
                    "Line Tax Amount": taxes,
                    "Line Description": item_description,
                    "Line Tax Code": "5% VAT" if taxes > 0 else "EX Exempt",
                    "Account": get_category_name(supplier_id)
                }


                invoices.append(line_data)
        
        # Create a DataFrame
        df = pd.DataFrame(invoices)
        
        return invoice_count, invoice_item_count, df
    else:
        st.error(f"Failed to fetch invoices. Status code: {response.status_code}")
        return 0, 0, pd.DataFrame()

def bill_save_csv_files(df, start_date_str, end_date_str):
    # Exclude rows with negative amounts
    df_positive = df[df["Line Amount"] >= 0]
    
    # Group by invoice number
    grouped = df_positive.groupby("Bill No")
    
    # Create chunks
    chunk_size = 1000
    current_chunk = []
    chunks = []
    current_size = 0
    
    for name, group in grouped:
        group_size = len(group)
        if current_size + group_size > chunk_size:
            chunks.append(pd.concat(current_chunk))
            current_chunk = []
            current_size = 0
        current_chunk.append(group)
        current_size += group_size
    
    if current_chunk:
        chunks.append(pd.concat(current_chunk))
    
    csv_files = []
    for idx, chunk in enumerate(chunks):
        file_name = f'bills_{start_date_str}_{end_date_str}_part{idx+1}.csv'
        chunk_csv = chunk.to_csv(index=False)
        csv_files.append((file_name, chunk_csv))
    
    return csv_files

def bill_save_credit_memo_files(df, start_date_str, end_date_str):
    credit_memo_df = df[df["Line Amount"] < 0]
    if not credit_memo_df.empty:
        credit_memo_file_name = f'supplier_credit_{start_date_str}_{end_date_str}.csv'
        credit_memo_csv = credit_memo_df.to_csv(index=False)
        return credit_memo_file_name, credit_memo_csv
    return None, None

def qb_bills():
    pd.options.mode.copy_on_write = True
    title = 'Juniper Bills Generator'
    st.markdown(f"<h1 style='font-size:24px;'>{title}</h1>", unsafe_allow_html=True)

    invoice_date_from = st.date_input("Bill Date From")
    invoice_date_to = st.date_input("Bill Date To")

    if "bill_csv_files" not in st.session_state:
        st.session_state.bill_csv_files = []
    if "bill_credit_memo_file" not in st.session_state:
        st.session_state.bill_credit_memo_file = None
    if "bill_invoice_count" not in st.session_state:
        st.session_state.bill_invoice_count = 0
    if "bill_invoice_item_count" not in st.session_state:
        st.session_state.bill_invoice_item_count = 0
    if "bill_df" not in st.session_state:
        st.session_state.bill_df = pd.DataFrame()

    # Validation checks
    if invoice_date_from and invoice_date_to:
        same_month = invoice_date_from.month == invoice_date_to.month
        if not same_month:
            st.error("The period must be within the same month.")
        else:
            if st.button('Fetch Bills'):
                with st.spinner('Fetching bills...'):
                    invoice_date_from_str = invoice_date_from.strftime("%Y%m%d")
                    invoice_date_to_str = invoice_date_to.strftime("%Y%m%d")
                    fetch_and_populate_suppliers()
                    invoice_count, invoice_item_count, bill_df = fetch_bills(invoice_date_from_str, invoice_date_to_str)

                    if not bill_df.empty:
                        bill_df.index = range(1, len(bill_df) + 1)
                        st.session_state.bill_csv_files = bill_save_csv_files(bill_df, invoice_date_from_str, invoice_date_to_str)
                        st.session_state.bill_invoice_count = invoice_count
                        st.session_state.bill_invoice_item_count = invoice_item_count
                        st.session_state.bill_df = bill_df

                        # Save credit memo files
                        credit_memo_file_name, credit_memo_csv = bill_save_credit_memo_files(bill_df, invoice_date_from_str, invoice_date_to_str)
                        if credit_memo_csv:
                            st.session_state.bill_credit_memo_file = (credit_memo_file_name, credit_memo_csv)
                        else:
                            st.session_state.bill_credit_memo_file = None

    if st.session_state.bill_invoice_count:
        st.write(f"Number of bills: {st.session_state.bill_invoice_count}")
    if st.session_state.bill_invoice_item_count:
        st.write(f"Number of bill items: {st.session_state.bill_invoice_item_count}")
    if not st.session_state.bill_df.empty:
        st.write(st.session_state.bill_df)
    
    if st.session_state.bill_csv_files:
        for file_name, csv in st.session_state.bill_csv_files:
            st.download_button(
                label=f"📥 Download {file_name}",
                data=csv,
                file_name=file_name,
                mime='text/csv',
            )
    
    if st.session_state.bill_credit_memo_file:
        file_name, csv = st.session_state.bill_credit_memo_file
        st.download_button(
            label=f"📥 Download {file_name}",
            data=csv,
            file_name=file_name,
            mime='text/csv',
        )

def streamlit_menu():
        with st.sidebar:
            selected = option_menu(
                menu_title="Main Menu",  # required
                options=["Juniper Invoices", "Juniper Bills", "Cities","Services", "Suppliers", "Fees Setup", "VAT Report Generator"],  # required
                icons=["house", "receipt", "pin-map", "buildings", "gear", "calculator", "receipt"],  # optional
                menu_icon="cast",  # optional
                default_index=0,  # optional
            )
        return selected


def main():
    st.sidebar.markdown("""
    <style>
        .st-emotion-cache-dvne4q {
        margin-top: -75px;
        }
    </style>
    """, unsafe_allow_html=True)    
    st.sidebar.image("https://res.cloudinary.com/blackrock/image/upload/v1674116709/dubailink/dl-logo-icon_psmtsg.png", width=150)

    app_mode = streamlit_menu()

    if app_mode == "VAT Report Generator":
        main_app()
    elif app_mode == "Cities":
        area_state_country_editor('areas.csv')
    elif app_mode == "Suppliers":
        product_supplier_editor('suppliers.csv')
    elif app_mode == "Services":
        service_type_editor('services.csv')   
    elif app_mode == "Fees Setup":
        vat_setup_editor('vat_setup.csv')
    elif app_mode == "Juniper Invoices":
        qb_invoices()   
    elif app_mode == "Juniper Bills":
        qb_bills()

if __name__ == "__main__":
    main()
