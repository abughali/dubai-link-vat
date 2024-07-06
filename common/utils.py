import pandas as pd
import streamlit as st

def load_emirates(filename):
    try:
        df = pd.read_csv(filename)
        return df['Emirate'].dropna().unique().tolist()
    except FileNotFoundError:
        return []
    
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

def load_service_types(filename):
    try:
        df = pd.read_csv(filename)
        return df['Service Type'].dropna().unique().tolist()
    except FileNotFoundError:
        return []
    
def save_rules(rules_df, filename):
    rules_df.to_csv(filename, index=False)

def hide_home_page():

    styling = f"""
                    div[data-testid=\"stSidebarNav\"] li:nth-child({1}) {{
                        display: none;
                    }}
                """

    styling = f"""
        <style>
            {styling}
        </style>
    """

    st.write(
        styling,
        unsafe_allow_html=True,
    )

def add_logo():
    st.markdown(
        """
        <style>
            [data-testid="stSidebar"] {
                background-image: url(https://i.postimg.cc/mZc99z5t/dl.png);
                background-repeat: no-repeat;
                padding-top: 120px;
                background-position: 20px 20px;
            }
            .block-container
            {
                padding-top: 1rem;
                padding-bottom: 0rem;
                margin-top: 1rem;
            }            
        </style>
        """,
        unsafe_allow_html=True,
    )