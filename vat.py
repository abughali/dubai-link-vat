import streamlit as st
import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth
from common import utils


def main():
 
    utils.add_logo()

    _, cent_co, _ = st.columns(3)
    with cent_co:
        st.image("./images/dl.png", use_column_width=False)
    
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