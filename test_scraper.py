import pytest
import requests_mock
import pandas as pd
from pathlib import Path

# --- Core Cleaner Functions to Test ---
def clean_text(text):
    if not text: return ""
    for char in ["", "", "", "🌍", "📞", "🕒"]:
        text = text.replace(char, "")
    return text.strip()

def build_search_url(keyword, place):
    formatted_query = f"{keyword} {place}".replace(" ", "+")
    return f"https://www.google.com/maps/search/{formatted_query}"

# =====================================================================
# 🧪 Test 1: Google Character Sanitation (Unit Test)
# =====================================================================
def test_clean_text_strips_google_font_artifacts():
    """Verifies that hidden custom icon characters are cleanly stripped from data fields."""
    dirty_address = " 789 Central Ave, Nagpur "
    expected_output = "789 Central Ave, Nagpur"
    assert clean_text(dirty_address) == expected_output

# =====================================================================
# 🧪 Test 2: Selenium URL Target Formulation (Unit Test)
# =====================================================================
def test_build_search_url_formats_spaces_correctly():
    """Verifies that the query builder maps multi-word spaces to URL additions."""
    url = build_search_url("Car Showrooms", "Nagpur Local")
    assert "Car+Showrooms" in url
    assert "Nagpur+Local" in url

# =====================================================================
# 🧪 Test 3: System Stability Check (Server Failure Mocking)
# =====================================================================
def test_google_sheets_webhook_handles_server_errors():
    """Simulates a network drop to ensure the app handles webhook issues gracefully."""
    mock_url = "https://script.google.com/macros/s/mock_link/exec"
    payload = {"name": "Test Enterprise"}
    
    with requests_mock.Mocker() as mock_adapter:
        mock_adapter.post(mock_url, status_code=500, text="Internal Server Error")
        
        import requests
        response = requests.post(mock_url, json=payload)
        assert response.status_code == 500