"""
Mock streamlit before any test module imports app.py.
pytest loads conftest.py first, so sys.modules is patched before the import.
"""
import sys
from unittest.mock import MagicMock

import pytest


# ── Minimal st.session_state stand-in ────────────────────────────────────────

class SessionState:
    def __init__(self):
        object.__setattr__(self, "_d", {})

    def __getattr__(self, key):
        return object.__getattribute__(self, "_d").get(key)

    def __setattr__(self, key, value):
        object.__getattribute__(self, "_d")[key] = value

    def __contains__(self, key):
        return key in object.__getattribute__(self, "_d")

    def __delattr__(self, key):
        object.__getattribute__(self, "_d").pop(key, None)

    def __setitem__(self, key, value):
        object.__getattribute__(self, "_d")[key] = value

    def __getitem__(self, key):
        return object.__getattribute__(self, "_d")[key]

    def pop(self, key, *args):
        return object.__getattribute__(self, "_d").pop(key, *args)

    def reset(self):
        object.__getattribute__(self, "_d").clear()


_session_state = SessionState()

_st = MagicMock()
_st.session_state = _session_state

# Prevent widget return values from being truthy, which would cause
# conditional UI blocks to run at import time and hit real I/O.
_st.file_uploader.return_value = None
_st.button.return_value = False
_st.checkbox.return_value = False
_st.multiselect.return_value = []
_st.text_area.return_value = ""
_st.text_input.return_value = ""
_st.selectbox.return_value = None
_st.radio.return_value = "batch"
_st.toggle.return_value = True
_st.slider.return_value = 500
_st.number_input.return_value = 5

sys.modules["streamlit"] = _st
sys.modules["st_copy"] = MagicMock()


# ── Session state reset fixture (runs before every test) ─────────────────────

@pytest.fixture(autouse=True)
def fresh_state():
    _session_state.reset()
    _session_state.all_groups = None
    _session_state.missing_ids = []
    _session_state.merge_selected = set()
    _session_state.df = None
    yield
    _session_state.reset()
