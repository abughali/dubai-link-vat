import streamlit as st
import pandas as pd
import streamlit_authenticator as stauth
import yaml
from yaml.loader import SafeLoader
from common import juniper_api, utils

def bill_save_csv_files(df, start_date_str, end_date_str):
    # Exclude rows with negative amounts
    df_positive = df[df["Line Amount"] >= 0]
    
    # Group by invoice number
    grouped = df_positive.groupby("Bill No")
    
    # Create chunks
    chunk_size = 10000
    current_chunk = []
    chunks = []
    current_size = 0
    
    for _, group in grouped:
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
    credit_memo_df = df[df["Line Amount"] < 0].copy()

    credit_memo_df["Line Amount"] = credit_memo_df["Line Amount"].abs()
    credit_memo_df["Line Tax Amount"] = credit_memo_df["Line Tax Amount"].abs()

    if not credit_memo_df.empty:
        credit_memo_file_name = f'vendor_credit_{start_date_str}_{end_date_str}.csv'
        credit_memo_csv = credit_memo_df.to_csv(index=False)
        return credit_memo_file_name, credit_memo_csv

    return None, None

def qb_bills():
    pd.options.mode.copy_on_write = True
    title = 'Juniper Bills Generator'
    st.markdown(f"<h1 style='font-size:24px;'>{title}</h1>", unsafe_allow_html=True)

    invoice_date_from = st.date_input("Bill Date From")
    invoice_date_to = st.date_input("Bill Date To", value=invoice_date_from)

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
                    juniper_api.fetch_and_populate_suppliers()
                    invoice_count, invoice_item_count, bill_df = juniper_api.fetch_bills(invoice_date_from_str, invoice_date_to_str)

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
    qb_bills()
elif st.session_state["authentication_status"] is False:
    st.error('Username/password is incorrect')