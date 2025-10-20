# main.py
import streamlit as st
from app.ui import run_app

st.set_page_config(page_title="Bin Label Generator", layout="wide")
run_app()
