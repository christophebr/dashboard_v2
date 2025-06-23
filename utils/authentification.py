import streamlit as st
import streamlit_authenticator as stauth
import bcrypt
from pathlib import Path
import pickle

def authenticate_user():
    file_path = Path(__file__).parent.parent / 'hashed_pw.pkl'
    with file_path.open('rb') as file:
        hashed_passwords = pickle.load(file)

    credentials = {
        'usernames': {
            'cbri': {'name': 'Christophe Brichet', 'password': hashed_passwords['cbri']},
            'mpec': {'name': 'Miguel Pecqueux', 'password': hashed_passwords['mpec']},
            'elap': {'name': 'Eric Laporte', 'password': hashed_passwords['elap']},
            'pgou': {'name': 'Pierre Goupillon', 'password': hashed_passwords['pgou']},
            'osai': {'name': 'Olivier Sainte-Rose', 'password': hashed_passwords['osai']},
            'mhum': {'name': 'Mourad Humblot', 'password': hashed_passwords['mhum']},
            'fsau': {'name': 'Frederic Sauvan', 'password': hashed_passwords['fsau']},
            'akes': {'name': 'Archimede Kessi', 'password': hashed_passwords['akes']},
            'dlau': {'name': 'Dorian Lau', 'password': hashed_passwords['dlau']},
            'jdel': {'name': 'Julie Delplanque', 'password': hashed_passwords['jdel']},
        }
    }

    authenticator = stauth.Authenticate(credentials, 'dashboard_support', 'support', cookie_expiry_days=2)
    name, authentication_status, username = authenticator.login()

    if authentication_status == False:
        st.error('Username/password is incorrect')
    elif authentication_status == None:
        st.warning('Please enter your username and password')
    else:
        if 'name' not in st.session_state:
            st.session_state['name'] = name
        st.sidebar.title(f"Welcome, {st.session_state['name']}")
        authenticator.logout('Logout', 'sidebar')
