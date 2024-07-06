import streamlit as st
import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth
from common import utils


def main():
 
    utils.add_logo()

    left_co, cent_co, last_co = st.columns(3)
    with cent_co:
        st.image("./images/dl.png", use_column_width=False)
    

 #   app_mode = streamlit_menu()

#    if app_mode == "VAT Report Generator":
#        main_app()
#    elif app_mode == "Cities":
#        area_state_country_editor('areas.csv')
#    elif app_mode == "Suppliers":
#        product_supplier_editor('suppliers.csv')
#    elif app_mode == "Services":
#        service_type_editor('services.csv')   
#    elif app_mode == "Fees Setup":
#        vat_setup_editor('vat_setup.csv')
#    elif app_mode == "Juniper Invoices":
#        qb_invoices()   
#    elif app_mode == "Juniper Bills":
#        qb_bills()


main()
utils.hide_home_page()


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
    st.write(f'### Welcome, {st.session_state["name"]}!')
elif st.session_state["authentication_status"] is False:
    st.error('Username/password is incorrect')