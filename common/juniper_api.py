import streamlit as st
import pandas as pd
import requests
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta


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
    

def get_booking_details(begin_travel_date_from, begin_travel_date_to):

    date_from = datetime.strptime(begin_travel_date_from, '%Y%m%d')
    date_from_prev = date_from - timedelta(days=1)
    begin_travel_date_from = date_from_prev.strftime('%Y%m%d')

    url = 'https://www.gte.travel/wsExportacion/wsbookings.asmx/getBookings'
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {
        'user': st.secrets["gte_user"],
        'password': st.secrets["gte_password"],    
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
        'BookingCode': '',
        'Status': '',
        'id': '',
        'ExportMode': '',
        'channel': '',
        'ModuleType': '',
        'IdBooking': '',
        'AgencyRef': '',
        'BeginTravelDateFrom': begin_travel_date_from,
        'BeginTravelDateTo': begin_travel_date_to,
        'EndTravelDateFrom': '',
        'EndTravelDateTo': '',
        'PackageBookings': '',
        'BlockedBookings': ''
    }

    response = requests.post(url, headers=headers, data=data)
    
    if response.status_code != 200:
        print(f"Error: Unable to fetch data, status code: {response.status_code}")
        return []

    # Parse the XML data
    root = ET.fromstring(response.content)

    # Find all Booking elements
    bookings = root.findall('.//Booking')
    
    results = []

    # Loop through each Booking element found
    for booking in bookings:
        booking_code = booking.get('BookingCode')
        status = booking.get('Status')
        
        # Loop through each Line element within the current Booking
        for line in booking.findall('.//Line'):
            id_book_line = line.get('IdBookLine')

            # Extract CostAmountToBeInvoiced
            cost_amount = line.findtext('.//CostAmountToBeInvoiced')
            cost_amount = float(cost_amount) if cost_amount is not None else 0.0

            # Sum all totalcost values for CostTaxes within the current Line
            total_cost_taxes = sum(float(tax.findtext('totalcost', default='0.0')) for tax in line.findall('.//Tax'))

            results.append([booking_code, id_book_line, cost_amount, total_cost_taxes, status])

    return results


def get_booking_details_by_code(booking_code):
    
    url = 'https://www.gte.travel/wsExportacion/wsbookings.asmx/getBookings'
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {
        "user": st.secrets["gte_user"],
        "password": st.secrets["gte_password"],    
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
        'BookingCode': booking_code,
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
        'BlockedBookings': ''
    }

    response = requests.post(url, headers=headers, data=data)
    
    if response.status_code != 200:
        print(f"Error: Unable to fetch data, status code: {response.status_code}")
        return []

    # Parse the XML data
    root = ET.fromstring(response.content)

    # Find the Booking element with the specified BookingCode
    booking = root.find(f'.//Booking')
    
    if booking is None:
        return []

    results = []

    status = booking.get('Status')

    # Loop through each Line element within the found Booking
    for line in booking.findall('.//Line'):
        id_book_line = line.get('IdBookLine')

        # Extract CostAmountToBeInvoiced
        cost_amount = line.findtext('.//CostAmountToBeInvoiced')
        cost_amount = float(cost_amount) if cost_amount is not None else 0.0

        # Sum all totalcost values for CostTaxes within the current Line
        total_cost_taxes = sum(float(tax.findtext('totalcost', default='0.0')) for tax in line.findall('.//Tax'))

        results.append([booking_code, id_book_line, cost_amount, total_cost_taxes, status])

    return results

def lookup_booking_and_line(details, booking_code, id_book_line):
    # Find and return the details for the given booking_code and id_book_line
    for detail in details:
        if detail[0] == booking_code and detail[1] == id_book_line:
            return detail
    # If not found, return None
    return None

def currency_converter(amount, cost_rate, sell_rate):
    first_conversion = amount * cost_rate
    final_conversion = first_conversion * sell_rate
    return round(final_conversion, 2)


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

        seq = 1
        invoices = []

        details = get_booking_details(invoice_date_from, invoice_date_to)


        for invoice in root.findall(".//Invoice"):
            invoice_number = invoice.get("InvoiceNumber")
            invoice_date = format_date(invoice.get("InvoiceDate"))
            due_date = format_date(invoice.get("DueDate"))
            operation_rate_elem = invoice.find(".//OperationRate")
            sell_exchange_rate = float(operation_rate_elem.text) if operation_rate_elem is not None and operation_rate_elem.text else 1.0

            for line in invoice.findall(".//Line"):
                
                cost_elem = line.find(".//Cost")
                # Extract SupplierId from Cost element
                supplier_id = cost_elem.get("SupplierId") if cost_elem is not None else ""
                supplier_name = line.find(".//SupplierName").text
                service = line.find(".//ArticleOfCost").text
                begin_travel_date = format_date(line.get("BeginTravelDate"))
                end_travel_date = format_date(line.get("EndTravelDate"))
                
                item_description = f"{service}\nTravel Date {begin_travel_date} - {end_travel_date}"
                cost_exchange_rate = float(cost_elem.get("ExchangeRate"))
                supplier_cost = float(cost_elem.get("TotalAmount"))
                currency = cost_elem.get("Currency")

                booking_code = line.get("BookingCode")
                id_book_line = line.get("IdBookingLine")

                # Lookup specific booking code and line ID
                
                lookup_result = lookup_booking_and_line(details, booking_code, id_book_line)

                if lookup_result:
                    cost, tax, status = lookup_result[2], lookup_result[3], lookup_result[4]
                else:
                    # If booking code not found, fetch details by booking code
                    print(f"Booking code {booking_code} not found in initial results, fetching details by booking code.")
                    additional_details = get_booking_details_by_code(booking_code)
                    additional_lookup_result = lookup_booking_and_line(additional_details, booking_code, id_book_line)

                    if additional_lookup_result:
                        cost, tax, status = additional_lookup_result[2], additional_lookup_result[3], additional_lookup_result[4]
                    else:
                        print(f"No details found for booking code {booking_code} and line ID {id_book_line}")
                

                item_amount = currency_converter(cost, cost_exchange_rate, sell_exchange_rate)
                taxes = currency_converter(tax, cost_exchange_rate, sell_exchange_rate)

                print(f"{seq:<6}{invoice_number:<8}{booking_code:<8}{id_book_line:<8}{supplier_cost:<10.2f}{cost:<10.2f}{tax:<10.2f}{item_amount:<12.2f}{taxes:<10.2f}{currency:<6}{status:<6}")

                seq +=1


                line_data = {
                    "Bill No": invoice_number,
                    "Bill Date": invoice_date,
                    "DueDate": due_date,
                    "Currency": "AED",
                    "Supplier": supplier_name,
                    "Memo": booking_code,
                    "Line Amount": item_amount,
                    "Line Tax Amount": taxes,
                    "Line Description": item_description,
                    "Line Tax Code": "5% VAT" if taxes > 0 else "EX Exempt",
                    "Account": get_category_name(supplier_id)
                }
                

                if item_amount != 0:
                    invoices.append(line_data)
                    
        
        # Create a DataFrame
        df = pd.DataFrame(invoices)
        invoice_item_count = df['Bill No'].count()
        invoice_count = df['Bill No'].nunique()
        
        return invoice_count, invoice_item_count, df
    else:
        st.error(f"Failed to fetch invoices. Status code: {response.status_code}")
        return 0, 0, pd.DataFrame()
