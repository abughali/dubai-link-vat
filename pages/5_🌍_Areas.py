import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder
import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth
from common import utils


def area_state_country_editor(filename):
    title = "Areas Mapping"
    st.markdown(f"<h3 style='font-size:24px;'>{title}</h3>", unsafe_allow_html=True)
    rules_df = utils.load_rules(filename)
    if rules_df.empty:
        rules_df = pd.DataFrame(columns=['Area', 'Emirate'])

    # Ag-Grid implementation
    grid_options = GridOptionsBuilder.from_dataframe(rules_df)
    grid_options.configure_pagination(paginationAutoPageSize=True)
    grid_options.configure_side_bar(False, False)
    grid_options.configure_selection('single')
    AgGrid(rules_df, gridOptions=grid_options.build(), update_mode='MODEL_CHANGED', fit_columns_on_grid_load=True)

    emirates = utils.load_emirates('vat_setup.csv')

    # Add or update rule
    with st.form(key='form_area_state_country', clear_on_submit=True):
        new_area = st.text_input("Area")
        new_state = st.selectbox("Emirate", options=emirates)
        submitted = st.form_submit_button("💾 Save")
        if submitted:
            new_row = pd.DataFrame([{'Area': new_area, 'Emirate': new_state}])
            rules_df = pd.concat([rules_df, new_row]).drop_duplicates(subset=['Area', 'Emirate'], keep='last')
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
    area_state_country_editor('areas.csv')
elif st.session_state["authentication_status"] is False:
    st.error('Username/password is incorrect')