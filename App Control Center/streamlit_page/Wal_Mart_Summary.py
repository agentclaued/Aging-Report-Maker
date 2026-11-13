"""Streamlit page that shows the WalMartSummary app inside an existing Streamlit app.

Copy this file into the `pages/` folder of the Streamlit app on 192.168.2.14 (next to its main script),
then restart/refresh Streamlit. The WalMartSummary server keeps running on its own PC.
"""
import socket
from urllib.parse import urlparse

import streamlit as st
import streamlit.components.v1 as components

APP_URL = "http://192.168.2.123:8040"   # HOST:PORT from app_runtime.env of the WalMartSummary PC
FRAME_HEIGHT = 1000

st.set_page_config(page_title="Wal Mart Summary", layout="wide")


def app_is_up(url: str) -> bool:
    parsed = urlparse(url)
    try:
        with socket.create_connection((parsed.hostname, parsed.port or 80), timeout=2):
            return True
    except OSError:
        return False


if app_is_up(APP_URL):
    st.link_button("Open in a new tab", APP_URL)
    components.iframe(APP_URL, height=FRAME_HEIGHT, scrolling=True)
else:
    st.error(f"The Wal Mart Summary app is not running at {APP_URL}.")
    st.info("On the Wal Mart PC run '11 1 ON - Start in background.bat' (App Control Center), then press Rerun.")
    st.button("Rerun")
