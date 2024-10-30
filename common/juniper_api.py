import streamlit as st
import pandas as pd
import requests
import xml.etree.ElementTree as ET
from datetime import datetime
from ratelimit import limits, sleep_and_retry
from concurrent.futures import ThreadPoolExecutor as PoolExecutor

# Define the global supplier list
supplierList = []

def fetch_and_populate_suppliers():

    global supplierList
    supplierList.clear()  # Clear the list to fetch fresh data

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
        "3": "Static Hotel",
        "4": "Extranet Hotel",
        "5": "XML Hotel",
        "6": "Tickets",
        "7": "Offline Hotel",
        "8": "Excursions",
        "9": "Visa"
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
    global supplierList
    for supplier in supplierList:
        if supplier["Supplier Id"] == supplier_id:
            return supplier["Category Name"]
    return "Supplier ID not found"

# Function to remove time from datetime string
def format_date(date_str):
    if date_str:
        return datetime.strptime(date_str, "%Y-%m-%dT%H:%M:%S").strftime("%Y-%m-%d")
    return ""

def get_account_manager(customer_id):
    # URL and headers
    url = "https://www.gte.travel/wsExportacion/wsCustomers.asmx/getCustomerList"
    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }

    # Data payload
    data = {
        "user": st.secrets["gte_user"],
        "password": st.secrets["gte_password"],      
        "customerType": "",
        "creationDateFrom": "",
        "creationDateTo": "",
        "id": customer_id,
        "BranchType": "",
        "ExportMode": "",
        "LastModifiedDateFrom": "",
        "LastModifiedDateTo": "",
        "LastModifiedTimeFrom": "",
        "LastModifiedTimeTo": "",
        "AmountBaseCurrency": ""
    }

    try:
        response = requests.post(url, headers=headers, data=data)
        response.raise_for_status()

        content = response.text
        root = ET.fromstring(content)

        account_manager = root.find('.//AccountManager')

        if account_manager is not None:
            return account_manager.text.strip()
        else:
            print(f"Warning: No Account Manager found for {customer_id}")
            return ""

    except requests.RequestException as e:
        print(f"Error fetching data for {customer_id}: {e}")
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
                    # Extracting Customer Id
            customer = invoice.find(".//Customer")
            customer_id = customer.get("Id") if customer is not None else ""
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
                    "Service": get_category_name(supplier_id),
                    "Account Manager": get_account_manager(customer_id)
                }

                invoices.append(line_data)
        
        # Create a DataFrame
        df = pd.DataFrame(invoices)
        return invoice_count, invoice_item_count, df
    else:
        st.error(f"Failed to fetch invoices. Status code: {response.status_code}")
        return 0, 0, pd.DataFrame()


@sleep_and_retry
@limits(calls=1000, period=1)
def get_booking_details(booking_code):
        url = 'https://www.gte.travel/wsExportacion/wsbookings.asmx/getBookings'
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        data = {
            "user": st.secrets["gte_user"],
            "password": st.secrets["gte_password"],   
            'BookingCode': booking_code,
            'BookingDateFrom': '',
            'BookingDateTo': '',
            'BookingTimeFrom': '',
            'BookingTimeTo': '',
            'BeginTravelDate': '',
            'EndTravelDate': '',
            'LastModifiedDateFrom': '',
            'LastModifiedDateTo': '',
            'LastModifiedTimeFrom': '',
            'LastModifiedTimeTo': '',
            'Status': '',
            'id': '',
            'ExportMode': '',
            'channel': '',
            'ModuleType': '',
            'IdBooking': '',
            'AgencyRef': '',
            'BeginTravelDateFrom': '',
            'BeginTravelDateTo': '',
            'EndTravelDateFrom': '',
            'EndTravelDateTo': '',
            'PackageBookings': '',
            'BlockedBookings': ''    }

        try:
            response = requests.post(url, headers=headers, data=data)
            response.raise_for_status()  # Check for HTTP errors

            content = response.text
            root = ET.fromstring(content)
            booking = root.find('.//Booking')

            if booking is None:
                print(f"Warning: No booking details found for {booking_code}")
                return []

            results = []
            status = booking.get('Status')
            for line in booking.findall('.//Line'):
                id_book_line = line.get('IdBookLine')
                cost_amount = line.findtext('.//CostAmountToBeInvoiced')
                cost_amount = float(cost_amount) if cost_amount is not None else 0.0
                comm_amount = line.findtext('ComissionAmount')
                comm_amount = float(comm_amount) if comm_amount is not None else 0.0                
                total_cost_taxes = sum(float(tax.findtext('totalcost', default='0.0')) for tax in line.findall('.//Tax'))
                results.append([booking_code, id_book_line, cost_amount - comm_amount, total_cost_taxes, status])

            return results
        except requests.RequestException as e:
            print(f"Error fetching data for {booking_code}: {e}")
            return []

def fetch_booking_details_concurrently(booking_codes, max_workers=1000):
    results = []
    with PoolExecutor(max_workers=max_workers) as executor:
        for result in executor.map(get_booking_details, booking_codes):
            results.extend(result)
    return results

# Function to remove time from datetime string
def format_date(date_str):
    if date_str:
        return datetime.strptime(date_str, "%Y-%m-%dT%H:%M:%S").strftime("%Y-%m-%d")
    return ""

# Function for currency conversion
def currency_converter(amount, cost_rate, sell_rate):
    first_conversion = amount * cost_rate
    final_conversion = first_conversion * sell_rate
    return round(final_conversion, 2)

# Fetch bills function
def get_bill_details(invoice_date_from, invoice_date_to):
    url = "https://www.gte.travel/wsExportacion/wsinvoices.asmx/GetInvoices"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
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

    response = requests.post(url, headers=headers, data=data)

    if response.status_code == 200:
        root = ET.fromstring(response.text)
        invoices = []

        for invoice in root.findall(".//Invoice"):
            invoice_number = invoice.get("InvoiceNumber")
            invoice_date = format_date(invoice.get("InvoiceDate"))
            due_date = format_date(invoice.get("DueDate"))
            operation_rate_elem = invoice.find(".//OperationRate")
            sell_exchange_rate = float(operation_rate_elem.text) if operation_rate_elem is not None and operation_rate_elem.text else 1.0
            for line in invoice.findall(".//Line"):
                booking_code = line.get("BookingCode")
                id_book_line = line.get("IdBookingLine")
                supplier_name = line.find(".//SupplierName").text
                service = line.find(".//ArticleOfCost").text
                begin_travel_date = format_date(line.get("BeginTravelDate"))
                end_travel_date = format_date(line.get("EndTravelDate"))
                cost_elem = line.find(".//Cost")
                supplier_id = cost_elem.get("SupplierId") if cost_elem is not None else ""
                invoice_line_amount = float(line.get("TotalLineAmount"))
                supplier_cost = float(cost_elem.get("TotalAmount"))


                item_description = f"{service}\nTravel Date {begin_travel_date} - {end_travel_date}"
                cost_exchange_rate = float(cost_elem.get("ExchangeRate"))
                
                if (invoice_line_amount != 0):# and (supplier_cost != 0):
                    invoices.append({
                                "Bill No": invoice_number,
                                "Bill Date": invoice_date,
                                "DueDate": due_date,
                                "Currency": "AED",
                                "Supplier": supplier_name,
                                "Memo": booking_code,
                                "IdBookLine": id_book_line,
                                "Line Description": item_description,
                                "SellExchangeRate": sell_exchange_rate,
                                "CostExchangeRate": cost_exchange_rate,
                                "Account": get_category_name(supplier_id)
                            })

        df = pd.DataFrame(invoices)
        return df        
    else:
        st.error(f"Failed to fetch invoices. Status code: {response.status_code}")
        return pd.DataFrame()

def add_suffix_to_duplicate_bills(df):
    """
    Adds suffixes to duplicate Bill No where there are multiple suppliers.
    """
    # Identify duplicate 'Bill No' with different suppliers
    duplicate_bills = df.groupby('Bill No').filter(lambda x: x['Supplier'].nunique() > 1)

    # Apply suffix to duplicates
    for bill_no in duplicate_bills['Bill No'].unique():
        suppliers = df[df['Bill No'] == bill_no]['Supplier'].unique()
        for idx, supplier in enumerate(suppliers):
            if idx == 0:
                continue  # Skip the first supplier to retain the original Bill No
            suffix = f"-{idx}"
            df.loc[(df['Bill No'] == bill_no) & (df['Supplier'] == supplier), 'Bill No'] += suffix

    return df

def fetch_bills(invoice_date_from, invoice_date_to):

    bills = get_bill_details(invoice_date_from, invoice_date_to)
    booking_codes = bills["Memo"].unique().tolist()
    booking_details = fetch_booking_details_concurrently(booking_codes)
    booking_details_df = pd.DataFrame(booking_details, columns=["Memo", "IdBookLine", "Line Amount", "Line Tax Amount", "Status"])
    merged_df = pd.merge(bills, booking_details_df, on=["Memo", "IdBookLine"], how="inner")
    merged_df['Line Amount'] = merged_df.apply(lambda row: currency_converter(row['Line Amount'], row['CostExchangeRate'], row['SellExchangeRate']), axis=1)
    merged_df['Line Tax Amount'] = merged_df.apply(lambda row: currency_converter(row['Line Tax Amount'], row['CostExchangeRate'], row['SellExchangeRate']), axis=1)
    merged_df['Line Tax Code'] = merged_df['Line Tax Amount'].apply(lambda x: "5% VAT" if x > 0 else "EX Exempt")
    filtered_df = merged_df[merged_df['Line Amount'] != 0]
    filtered_df = filtered_df[['Bill No','Bill Date','DueDate','Currency','Supplier','Memo','Line Amount','Line Tax Amount','Line Description', 'Line Tax Code','Account']]
    filtered_df = filtered_df.sort_values(by='Bill No')

    bill_line_count = filtered_df['Bill No'].count()
    bill_count = filtered_df['Bill No'].nunique()
    filtered_df = add_suffix_to_duplicate_bills(filtered_df)

    return bill_count, bill_line_count, filtered_df