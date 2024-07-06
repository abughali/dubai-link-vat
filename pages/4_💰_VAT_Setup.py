import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder
import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth
from common import utils


def vat_setup_editor(filename):
    title = "VAT Setup"
    st.markdown(f"<h3 style='font-size:24px;'>{title}</h3>", unsafe_allow_html=True)
    rules_df = utils.load_rules(filename)
    if rules_df.empty:
        rules_df = pd.DataFrame(columns=['Emirate', 'Basic Division', 'Service Charge', 'Municipality Fee', 'VAT Percentage'])

    # Ag-Grid implementation
    grid_options = GridOptionsBuilder.from_dataframe(rules_df)
    grid_options.configure_pagination(paginationAutoPageSize=True)
    grid_options.configure_side_bar(False, False)
    grid_options.configure_selection('single')
    AgGrid(rules_df, gridOptions=grid_options.build(), update_mode='MODEL_CHANGED', fit_columns_on_grid_load=True)


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
            utils.save_rules(rules_df, filename)
            st.success("Saved!")
            st.rerun()

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
    vat_setup_editor('vat_setup.csv')
elif st.session_state["authentication_status"] is False:
    st.error('Username/password is incorrect')
