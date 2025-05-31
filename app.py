# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
import os
import json
import datetime
from PIL import Image  # For a potential logo
import time  # For retry delay and balloon visibility
import plotly.express as px  # Added for better charts

# --- Google API Imports ---
from google.oauth2 import service_account  # For Service Account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload, MediaIoBaseDownload  # Added MediaIoBaseDownload

# --- Gemini AI Import ---
import google.generativeai as genai

# --- Custom Module Imports ---
from dar_processor import preprocess_pdf_text  # Assuming dar_processor.py is in the same directory
from models import FlattenedAuditData, DARHeaderSchema, AuditParaSchema, ParsedDARReport

# --- Streamlit Option Menu for better navigation ---
from streamlit_option_menu import option_menu

# --- Configuration ---

YOUR_GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
 # !!! REPLACE WITH YOUR ACTUAL GEMINI API KEY !!!

SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
CREDENTIALS_FILE = 'credentials.json'

# Google Drive Master Configuration
MASTER_DRIVE_FOLDER_NAME = "e-MCM_Root_DAR_App"  # Master folder on Google Drive
MCM_PERIODS_FILENAME_ON_DRIVE = "mcm_periods_config.json"  # Config file on Google Drive

# --- User Credentials ---
USER_CREDENTIALS = {
    "planning_officer": "pco_password",
    **{f"audit_group{i}": f"ag{i}_audit" for i in range(1, 31)}
}
USER_ROLES = {
    "planning_officer": "PCO",
    **{f"audit_group{i}": "AuditGroup" for i in range(1, 31)}
}
AUDIT_GROUP_NUMBERS = {
    f"audit_group{i}": i for i in range(1, 31)
}


# --- Custom CSS Styling ---
def load_custom_css():
    st.markdown("""
    <style>
        /* --- Global Styles --- */
        body {
            font-family: 'Roboto', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #eef2f7;
            color: #4A4A4A;
            line-height: 1.6;
        }
        .stApp {
             background: linear-gradient(135deg, #f0f7ff 0%, #cfe7fa 100%);
        }

        /* --- Titles and Headers --- */
        .page-main-title {
            font-size: 3em;
            color: #1A237E;
            text-align: center;
            padding: 30px 0 10px 0;
            font-weight: 700;
            letter-spacing: 1.5px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
        }
        .page-app-subtitle {
            font-size: 1.3em;
            color: #3F51B5;
            text-align: center;
            margin-top: -5px;
            margin-bottom: 30px;
            font-weight: 400;
        }
        .app-description {
            font-size: 1.0em;
            color: #455A64;
            text-align: center;
            margin-bottom: 25px;
            padding: 0 20px;
            max-width: 700px;
            margin-left: auto;
            margin-right: auto;
        }
        .sub-header {
            font-size: 1.6em;
            color: #2779bd;
            border-bottom: 3px solid #5dade2;
            padding-bottom: 12px;
            margin-top: 35px;
            margin-bottom: 25px;
            font-weight: 600;
        }
        .card h3 {
            margin-top: 0;
            color: #1abc9c;
            font-size: 1.3em;
            font-weight: 600;
        }
         .card h4 {
            color: #2980b9;
            font-size: 1.1em;
            margin-top: 15px;
            margin-bottom: 8px;
        }

        /* --- Cards --- */
        .card {
            background-color: #ffffff;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 6px 12px rgba(0,0,0,0.08);
            margin-bottom: 25px;
            border-left: 6px solid #5dade2;
        }

        /* --- Streamlit Widgets Styling --- */
        .stButton>button {
            border-radius: 25px;
            background-image: linear-gradient(to right, #1abc9c 0%, #16a085 100%);
            color: white;
            padding: 12px 24px;
            font-weight: bold;
            border: none;
            transition: all 0.3s ease;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .stButton>button:hover {
            background-image: linear-gradient(to right, #16a085 0%, #1abc9c 100%);
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
        }
        .stButton>button[kind="secondary"] {
            background-image: linear-gradient(to right, #e74c3c 0%, #c0392b 100%);
        }
        .stButton>button[kind="secondary"]:hover {
            background-image: linear-gradient(to right, #c0392b 0%, #e74c3c 100%);
        }
        .stButton>button:disabled {
            background-image: none;
            background-color: #bdc3c7;
            color: #7f8c8d;
            box-shadow: none;
            transform: none;
        }
        .stTextInput>div>div>input, .stSelectbox>div>div>div, .stDateInput>div>div>input, .stNumberInput>div>div>input {
            border-radius: 8px;
            border: 1px solid #ced4da;
            padding: 10px;
        }
        .stTextInput>div>div>input:focus, .stSelectbox>div>div>div:focus-within, .stNumberInput>div>div>input:focus {
            border-color: #5dade2;
            box-shadow: 0 0 0 0.2rem rgba(93, 173, 226, 0.25);
        }
        .stFileUploader>div>div>button {
            border-radius: 25px;
            background-image: linear-gradient(to right, #5dade2 0%, #2980b9 100%);
            color: white;
            padding: 10px 18px;
        }
        .stFileUploader>div>div>button:hover {
            background-image: linear-gradient(to right, #2980b9 0%, #5dade2 100%);
        }

        /* --- Login Page Specific --- */
        .login-form-container {
            max-width: 500px;
            margin: 20px auto;
            padding: 30px;
            background-color: #ffffff;
            border-radius: 15px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
        }
        .login-form-container .stButton>button {
            background-image: linear-gradient(to right, #34495e 0%, #2c3e50 100%);
        }
        .login-form-container .stButton>button:hover {
            background-image: linear-gradient(to right, #2c3e50 0%, #34495e 100%);
        }
        .login-header-text {
            text-align: center;
            color: #1a5276;
            font-weight: 600;
            font-size: 1.8em;
            margin-bottom: 25px;
        }
        .login-logo { /* MODIFIED */
            display: block;
            margin-left: auto;
            margin-right: auto;
            max-width: 35px; /* Reduced size */
            margin-bottom: 15px;
            /* border-radius: 50%; REMOVED for no oval shape */
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        /* --- Sidebar Styling --- */
        .css-1d391kg {
            background-color: #ffffff;
            padding: 15px !important;
        }
        .sidebar .stButton>button {
             background-image: linear-gradient(to right, #e74c3c 0%, #c0392b 100%);
        }
        .sidebar .stButton>button:hover {
             background-image: linear-gradient(to right, #c0392b 0%, #e74c3c 100%);
        }
        .sidebar .stMarkdown > div > p > strong {
            color: #2c3e50;
        }

        /* --- Option Menu Customization --- */
        div[data-testid="stOptionMenu"] > ul {
            background-color: #ffffff;
            border-radius: 25px;
            padding: 8px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        }
        div[data-testid="stOptionMenu"] > ul > li > button {
            border-radius: 20px;
            margin: 0 5px !important;
            border: none !important;
            transition: all 0.3s ease;
        }
        div[data-testid="stOptionMenu"] > ul > li > button.selected {
            background-image: linear-gradient(to right, #1abc9c 0%, #16a085 100%);
            color: white;
            font-weight: bold;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        div[data-testid="stOptionMenu"] > ul > li > button:hover:not(.selected) {
            background-color: #e0e0e0;
            color: #333;
        }

        /* --- Links --- */
        a {
            color: #3498db;
            text-decoration: none;
            font-weight: 500;
        }
        a:hover {
            text-decoration: underline;
            color: #2980b9;
        }

        /* --- Info/Warning/Error Boxes --- */
        .stAlert {
            border-radius: 8px;
            padding: 15px;
            border-left-width: 5px;
        }
        .stAlert[data-baseweb="notification"][role="alert"] > div:nth-child(2) {
             font-size: 1.0em;
        }
        .stAlert[data-testid="stNotification"] {
            box-shadow: 0 2px 10px rgba(0,0,0,0.07);
        }
        .stAlert[data-baseweb="notification"][kind="info"] { border-left-color: #3498db; }
        .stAlert[data-baseweb="notification"][kind="success"] { border-left-color: #2ecc71; }
        .stAlert[data-baseweb="notification"][kind="warning"] { border-left-color: #f39c12; }
        .stAlert[data-baseweb="notification"][kind="error"] { border-left-color: #e74c3c; }

    </style>
    """, unsafe_allow_html=True)


# --- Google API Authentication and Service Initialization ---
# def get_google_services():
#     creds = None
#     if not os.path.exists(CREDENTIALS_FILE):
#         st.error(f"Service account credentials file ('{CREDENTIALS_FILE}') not found.")
#         return None, None
#     try:
#         creds = service_account.Credentials.from_service_account_file(
#             CREDENTIALS_FILE, scopes=SCOPES)
#     except Exception as e:
#         st.error(f"Failed to load service account credentials: {e}")
#         return None, None
#     if not creds: return None, None
#     try:
#         drive_service = build('drive', 'v3', credentials=creds)
#         sheets_service = build('sheets', 'v4', credentials=creds)
#         return drive_service, sheets_service
#     except HttpError as error:
#         st.error(f"An error occurred initializing Google services: {error}")
#         return None, None
#     except Exception as e:
#         st.error(f"An unexpected error with Google services: {e}")
#         return None, None

def get_google_services():
    creds = None
    # if not os.path.exists(CREDENTIALS_FILE): # Remove this check
    #     st.error(f"Service account credentials file ('{CREDENTIALS_FILE}') not found.")
    #     return None, None
    try:
        # Load credentials from Streamlit secrets
        creds_dict = st.secrets["google_credentials"]
        creds = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=SCOPES
        )
    except KeyError:
        st.error("Google credentials not found in Streamlit secrets. Ensure 'google_credentials' are set.")
        return None, None
    except Exception as e:
        st.error(f"Failed to load service account credentials from secrets: {e}")
        return None, None

    if not creds: return None, None # Should be caught by exceptions above mostly

    try:
        drive_service = build('drive', 'v3', credentials=creds)
        sheets_service = build('sheets', 'v4', credentials=creds)
        return drive_service, sheets_service
    except HttpError as error:
        st.error(f"An error occurred initializing Google services: {error}")
        return None, None
    except Exception as e:
        st.error(f"An unexpected error with Google services: {e}")
        return None, None
# --- Google Drive Helper Functions ---
def find_drive_item_by_name(drive_service, name, mime_type=None, parent_id=None):
    query = f"name = '{name}' and trashed = false"
    if mime_type:
        query += f" and mimeType = '{mime_type}'"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    # else: # If no parent_id, searches in locations accessible to service account, including root
    # query += " and 'root' in parents" # Be cautious with this if service account has broad access

    try:
        response = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        items = response.get('files', [])
        if items:
            return items[0].get('id')
    except HttpError as error:
        st.warning(f"Error searching for '{name}' in Drive: {error}. This might be okay if the item is to be created.")
    except Exception as e:
        st.warning(f"Unexpected error searching for '{name}' in Drive: {e}")
    return None


def set_public_read_permission(drive_service, file_id):
    try:
        permission = {'type': 'anyone', 'role': 'reader'}
        drive_service.permissions().create(fileId=file_id, body=permission).execute()
    except HttpError as error:
        st.warning(f"Could not set public read permission for file ID {file_id}: {error}.")
    except Exception as e:
        st.warning(f"Unexpected error setting public permission for file ID {file_id}: {e}")


def create_drive_folder(drive_service, folder_name, parent_id=None):  # MODIFIED
    try:
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        if parent_id:
            file_metadata['parents'] = [parent_id]

        folder = drive_service.files().create(body=file_metadata, fields='id, webViewLink').execute()
        folder_id = folder.get('id')
        if folder_id:
            set_public_read_permission(drive_service, folder_id)  # Consider if needed for master folder too
        return folder_id, folder.get('webViewLink')
    except HttpError as error:
        st.error(f"An error occurred creating Drive folder '{folder_name}': {error}")
        return None, None
    except Exception as e:
        st.error(f"Unexpected error creating Drive folder '{folder_name}': {e}")
        return None, None


# --- MCM Period Management (using Google Drive for storage) ---
def initialize_drive_structure(drive_service):
    master_id = st.session_state.get('master_drive_folder_id')
    if not master_id:
        master_id = find_drive_item_by_name(drive_service, MASTER_DRIVE_FOLDER_NAME,
                                            'application/vnd.google-apps.folder')
        if not master_id:
            st.info(f"Master folder '{MASTER_DRIVE_FOLDER_NAME}' not found on Drive, attempting to create it...")
            master_id, _ = create_drive_folder(drive_service, MASTER_DRIVE_FOLDER_NAME,
                                               parent_id=None)  # Create in root
            if master_id:
                st.success(f"Master folder '{MASTER_DRIVE_FOLDER_NAME}' created successfully.")
            else:
                st.error(f"Fatal: Failed to create master folder '{MASTER_DRIVE_FOLDER_NAME}'. Cannot proceed.")
                return False  # Critical failure
        st.session_state.master_drive_folder_id = master_id

    if not st.session_state.master_drive_folder_id:
        st.error("Master Drive folder ID could not be established. Cannot proceed.")
        return False

    mcm_file_id = st.session_state.get('mcm_periods_drive_file_id')
    if not mcm_file_id:  # Check if file ID for mcm_periods.json is already known
        mcm_file_id = find_drive_item_by_name(drive_service, MCM_PERIODS_FILENAME_ON_DRIVE,
                                              parent_id=st.session_state.master_drive_folder_id)
        if mcm_file_id:
            st.session_state.mcm_periods_drive_file_id = mcm_file_id
        # else:
        # st.info(f"'{MCM_PERIODS_FILENAME_ON_DRIVE}' not found in master folder. Will be created if MCM periods are saved.")
        # No need to explicitly set to None here, get will return None
    return True


def load_mcm_periods(drive_service):
    mcm_periods_file_id = st.session_state.get('mcm_periods_drive_file_id')
    if not mcm_periods_file_id:  # If ID isn't in session, try to find it (e.g., on first load after master folder init)
        if st.session_state.get('master_drive_folder_id'):
            mcm_periods_file_id = find_drive_item_by_name(drive_service, MCM_PERIODS_FILENAME_ON_DRIVE,
                                                          parent_id=st.session_state.master_drive_folder_id)
            st.session_state.mcm_periods_drive_file_id = mcm_periods_file_id  # Store if found
        else:  # Master folder not even known
            return {}

    if mcm_periods_file_id:
        try:
            request = drive_service.files().get_media(fileId=mcm_periods_file_id)
            fh = BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
            fh.seek(0)
            return json.load(fh)
        except HttpError as error:
            if error.resp.status == 404:  # File not found
                # st.warning(f"'{MCM_PERIODS_FILENAME_ON_DRIVE}' (ID: {mcm_periods_file_id}) not found on Drive. Returning empty periods.")
                st.session_state.mcm_periods_drive_file_id = None  # Reset ID as it's invalid
            else:
                st.error(f"Error loading '{MCM_PERIODS_FILENAME_ON_DRIVE}' from Drive: {error}")
            return {}
        except json.JSONDecodeError:
            st.error(f"Error decoding JSON from '{MCM_PERIODS_FILENAME_ON_DRIVE}'. File might be corrupted.")
            return {}
        except Exception as e:
            st.error(f"Unexpected error loading '{MCM_PERIODS_FILENAME_ON_DRIVE}': {e}")
            return {}
    return {}


def save_mcm_periods(drive_service, periods_data):
    master_folder_id = st.session_state.get('master_drive_folder_id')
    if not master_folder_id:
        st.error("Master Drive folder ID not set. Cannot save MCM periods configuration to Drive.")
        return False

    mcm_periods_file_id = st.session_state.get('mcm_periods_drive_file_id')
    file_content = json.dumps(periods_data, indent=4).encode('utf-8')
    fh = BytesIO(file_content)
    media_body = MediaIoBaseUpload(fh, mimetype='application/json', resumable=True)

    try:
        if mcm_periods_file_id:
            file_metadata_update = {'name': MCM_PERIODS_FILENAME_ON_DRIVE}  # Ensure name is correct
            updated_file = drive_service.files().update(
                fileId=mcm_periods_file_id,
                body=file_metadata_update,  # Send empty body if only media is updated, or metadata if name changes
                media_body=media_body,
                fields='id, name'  # Request ID and name back
            ).execute()
            # print(f"Updated '{MCM_PERIODS_FILENAME_ON_DRIVE}' with ID: {updated_file.get('id')}")
        else:
            file_metadata_create = {'name': MCM_PERIODS_FILENAME_ON_DRIVE, 'parents': [master_folder_id]}
            new_file = drive_service.files().create(
                body=file_metadata_create,
                media_body=media_body,
                fields='id, name'  # Request ID and name back
            ).execute()
            st.session_state.mcm_periods_drive_file_id = new_file.get('id')
            # print(f"Created '{MCM_PERIODS_FILENAME_ON_DRIVE}' with ID: {new_file.get('id')} in master folder.")
        return True
    except HttpError as error:
        st.error(f"Error saving '{MCM_PERIODS_FILENAME_ON_DRIVE}' to Drive: {error}")
        return False
    except Exception as e:
        st.error(f"Unexpected error saving '{MCM_PERIODS_FILENAME_ON_DRIVE}': {e}")
        return False


def upload_to_drive(drive_service, file_content_or_path, folder_id, filename_on_drive):
    try:
        file_metadata = {'name': filename_on_drive, 'parents': [folder_id]}
        media_body = None

        if isinstance(file_content_or_path, str) and os.path.exists(file_content_or_path):
            media_body = MediaFileUpload(file_content_or_path, mimetype='application/pdf', resumable=True)
        elif isinstance(file_content_or_path, bytes):
            fh = BytesIO(file_content_or_path)
            media_body = MediaIoBaseUpload(fh, mimetype='application/pdf', resumable=True)
        elif isinstance(file_content_or_path, BytesIO):
            file_content_or_path.seek(0)  # Ensure stream is at the beginning
            media_body = MediaIoBaseUpload(file_content_or_path, mimetype='application/pdf', resumable=True)
        else:
            st.error(f"Unsupported file content type for Google Drive upload: {type(file_content_or_path)}")
            return None, None

        if media_body is None:
            st.error("Media body for upload could not be prepared.")
            return None, None

        request = drive_service.files().create(
            body=file_metadata,
            media_body=media_body,
            fields='id, webViewLink'
        )
        file = request.execute()
        file_id = file.get('id')
        if file_id:
            set_public_read_permission(drive_service, file_id)
        return file_id, file.get('webViewLink')
    except HttpError as error:
        st.error(f"An API error occurred uploading to Drive: {error}")
        return None, None
    except Exception as e:
        st.error(f"An unexpected error in upload_to_drive: {e}")
        return None, None


# --- Google Sheets Functions ---
def create_spreadsheet(sheets_service, drive_service, title, parent_folder_id=None):  # Added parent_folder_id
    try:
        spreadsheet_body = {'properties': {'title': title}}
        spreadsheet = sheets_service.spreadsheets().create(body=spreadsheet_body,
                                                           fields='spreadsheetId,spreadsheetUrl').execute()
        spreadsheet_id = spreadsheet.get('spreadsheetId')

        if spreadsheet_id and drive_service:
            set_public_read_permission(drive_service, spreadsheet_id)
            # Move spreadsheet to the specified parent folder (e.g., master MCM folder)
            if parent_folder_id:
                file = drive_service.files().get(fileId=spreadsheet_id, fields='parents').execute()
                previous_parents = ",".join(file.get('parents'))
                drive_service.files().update(fileId=spreadsheet_id,
                                             addParents=parent_folder_id,
                                             removeParents=previous_parents,  # Remove from root if it was there
                                             fields='id, parents').execute()
        return spreadsheet_id, spreadsheet.get('spreadsheetUrl')
    except HttpError as error:
        st.error(f"An error occurred creating Spreadsheet: {error}")
        return None, None
    except Exception as e:
        st.error(f"An unexpected error occurred creating Spreadsheet: {e}")
        return None, None


def append_to_spreadsheet(sheets_service, spreadsheet_id, values_to_append):
    try:
        body = {'values': values_to_append}
        sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', '')
        first_sheet_title = sheets[0].get("properties", {}).get("title", "Sheet1")

        range_to_check = f"{first_sheet_title}!A1:L1"  # Assuming L is the last column
        result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,
                                                            range=range_to_check).execute()
        header_row_values = result.get('values', [])

        if not header_row_values:  # If sheet is empty or header not present
            header_values_list = [[  # Must be list of lists for append
                "Audit Group Number", "GSTIN", "Trade Name", "Category",
                "Total Amount Detected (Overall Rs)", "Total Amount Recovered (Overall Rs)",
                "Audit Para Number", "Audit Para Heading",
                "Revenue Involved (Lakhs Rs)", "Revenue Recovered (Lakhs Rs)",
                "DAR PDF URL", "Record Created Date"
            ]]
            sheets_service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=f"{first_sheet_title}!A1",  # Start at A1 for header
                valueInputOption='USER_ENTERED',
                body={'values': header_values_list}
            ).execute()

        append_result = sheets_service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{first_sheet_title}!A1",  # Append after existing data (or header)
            valueInputOption='USER_ENTERED',
            body=body  # body already contains {'values': values_to_append}
        ).execute()
        return append_result
    except HttpError as error:
        st.error(f"An error occurred appending to Spreadsheet: {error}")
        return None
    except Exception as e:
        st.error(f"Unexpected error appending to Spreadsheet: {e}")
        return None


def read_from_spreadsheet(sheets_service, spreadsheet_id, sheet_name="Sheet1"):
    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=sheet_name  # Read the whole sheet
        ).execute()
        values = result.get('values', [])
        if not values:
            return pd.DataFrame()  # No data
        else:
            expected_cols = [
                "Audit Group Number", "GSTIN", "Trade Name", "Category",
                "Total Amount Detected (Overall Rs)", "Total Amount Recovered (Overall Rs)",
                "Audit Para Number", "Audit Para Heading",
                "Revenue Involved (Lakhs Rs)", "Revenue Recovered (Lakhs Rs)",
                "DAR PDF URL", "Record Created Date"
            ]
            # Check if the first row is the header
            if values and values[0] == expected_cols:
                return pd.DataFrame(values[1:], columns=values[0])
            else:
                # st.warning(f"Spreadsheet '{sheet_name}' headers do not match expected or are missing. Attempting to load data without specific headers.")
                # Try to load with first row as header if it's not matching, or no header if it's clearly data
                if values:
                    # A simple heuristic: if the first row has fewer columns than expected, or contains numbers, might be data
                    if len(values[0]) < len(expected_cols) // 2 or any(isinstance(c, (int, float)) for c in values[0]):
                        return pd.DataFrame(values)  # Load without specific headers
                    else:
                        return pd.DataFrame(values[1:], columns=values[0])  # Assume first row is a different header
                else:  # Should be caught by 'if not values:'
                    return pd.DataFrame()
    except HttpError as error:
        st.error(f"An error occurred reading from Spreadsheet: {error}")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Unexpected error reading from Spreadsheet: {e}")
        return pd.DataFrame()


def delete_spreadsheet_rows(sheets_service, spreadsheet_id, sheet_id_gid, row_indices_to_delete):
    """Deletes specified rows from a sheet. Row indices are 0-based sheet data indices (after header)."""
    if not row_indices_to_delete:
        return True

    requests = []
    # Sort indices in descending order to avoid shifting issues during batch deletion
    # The row_indices_to_delete should be 0-based relative to the data (excluding header)
    # So, if deleting the first data row, it's index 0, which is sheet row 2 (1-indexed)
    # For API (0-indexed), this is startIndex = 1
    for data_row_index in sorted(row_indices_to_delete, reverse=True):
        sheet_row_start_index = data_row_index + 1  # +1 because Sheets API is 0-indexed for rows, and we skip header
        requests.append({
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id_gid,
                    "dimension": "ROWS",
                    "startIndex": sheet_row_start_index,
                    "endIndex": sheet_row_start_index + 1
                }
            }
        })

    if requests:
        try:
            body = {'requests': requests}
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body=body).execute()
            return True
        except HttpError as error:
            st.error(f"An error occurred deleting rows from Spreadsheet: {error}")
            return False
        except Exception as e:
            st.error(f"Unexpected error deleting rows: {e}")
            return False
    return True


# --- Gemini Data Extraction with Retry ---
def get_structured_data_with_gemini(api_key: str, text_content: str, max_retries=2) -> ParsedDARReport:
    if not api_key or api_key == "YOUR_API_KEY_HERE":  # Check actual key, not placeholder
        return ParsedDARReport(parsing_errors="Gemini API Key not configured in app.py.")
    if text_content.startswith("Error processing PDF with pdfplumber:") or \
            text_content.startswith("Error in preprocess_pdf_text_"):  # Catch preprocessing errors
        return ParsedDARReport(parsing_errors=text_content)

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash-latest')

    prompt = f"""
    You are an expert GST audit report analyst. Based on the following FULL text from a Departmental Audit Report (DAR),
    where all text from all pages, including tables, is provided, extract the specified information
    and structure it as a JSON object. Focus on identifying narrative sections for audit para details,
    even if they are intermingled with tabular data. Notes like "[INFO: ...]" in the text are for context only.

    The JSON object should follow this structure precisely:
    {{
      "header": {{
        "audit_group_number": "integer or null (e.g., if 'Group-VI' or 'Gr 6', extract 6; must be between 1 and 30)",
        "gstin": "string or null",
        "trade_name": "string or null",
        "category": "string ('Large', 'Medium', 'Small') or null",
        "total_amount_detected_overall_rs": "float or null (numeric value in Rupees)",
        "total_amount_recovered_overall_rs": "float or null (numeric value in Rupees)"
      }},
      "audit_paras": [
        {{
          "audit_para_number": "integer or null (primary number from para heading, e.g., for 'Para-1...' use 1; must be between 1 and 50)",
          "audit_para_heading": "string or null (the descriptive title of the para)",
          "revenue_involved_lakhs_rs": "float or null (numeric value in Lakhs of Rupees, e.g., Rs. 50,000 becomes 0.5)",
          "revenue_recovered_lakhs_rs": "float or null (numeric value in Lakhs of Rupees)"
        }}
      ],
      "parsing_errors": "string or null (any notes about parsing issues, or if extraction is incomplete)"
    }}

    Key Instructions:
    1.  Header Information: Extract `audit_group_number` (as integer 1-30, e.g., 'Group-VI' becomes 6), `gstin`, `trade_name`, `category`, `total_amount_detected_overall_rs`, `total_amount_recovered_overall_rs`.
    2.  Audit Paras: Identify each distinct para. Extract `audit_para_number` (as integer 1-50), `audit_para_heading`, `revenue_involved_lakhs_rs` (converted to Lakhs), `revenue_recovered_lakhs_rs` (converted to Lakhs).
    3.  Use null for missing values. Monetary values as float.
    4.  If no audit paras found, `audit_paras` should be an empty list [].

    DAR Text Content:
    --- START OF DAR TEXT ---
    {text_content}
    --- END OF DAR TEXT ---

    Provide ONLY the JSON object as your response. Do not include any explanatory text before or after the JSON.
    """

    attempt = 0
    last_exception = None
    while attempt <= max_retries:
        attempt += 1
        # print(f"\n--- Calling Gemini (Attempt {attempt}/{max_retries + 1}) ---")
        try:
            response = model.generate_content(prompt)
            cleaned_response_text = response.text.strip()
            if cleaned_response_text.startswith("```json"):
                cleaned_response_text = cleaned_response_text[7:]
            elif cleaned_response_text.startswith("`json"):
                cleaned_response_text = cleaned_response_text[6:]
            if cleaned_response_text.endswith("```"): cleaned_response_text = cleaned_response_text[:-3]

            if not cleaned_response_text:
                error_message = f"Gemini returned an empty response on attempt {attempt}."
                # print(error_message)
                last_exception = ValueError(error_message)
                if attempt > max_retries: return ParsedDARReport(parsing_errors=error_message)
                time.sleep(1 + attempt);
                continue

            json_data = json.loads(cleaned_response_text)
            if "header" not in json_data or "audit_paras" not in json_data:
                error_message = f"Gemini response (Attempt {attempt}) missing 'header' or 'audit_paras' key. Response: {cleaned_response_text[:500]}"
                # print(error_message)
                last_exception = ValueError(error_message)
                if attempt > max_retries: return ParsedDARReport(parsing_errors=error_message)
                time.sleep(1 + attempt);
                continue

            parsed_report = ParsedDARReport(**json_data)
            # print(f"Gemini call (Attempt {attempt}) successful. Paras found: {len(parsed_report.audit_paras)}")
            return parsed_report
        except json.JSONDecodeError as e:
            raw_response_text = locals().get('response',
                                             {}).text if 'response' in locals() else "No response text captured"
            error_message = f"Gemini output (Attempt {attempt}) was not valid JSON: {e}. Response: '{raw_response_text[:1000]}...'"
            # print(error_message)
            last_exception = e
            if attempt > max_retries: return ParsedDARReport(parsing_errors=error_message)
            time.sleep(attempt * 2)
        except Exception as e:
            raw_response_text = locals().get('response',
                                             {}).text if 'response' in locals() else "No response text captured"
            error_message = f"Error (Attempt {attempt}) during Gemini/Pydantic: {type(e).__name__} - {e}. Response: {raw_response_text[:500]}"
            # print(error_message)
            last_exception = e
            if attempt > max_retries: return ParsedDARReport(parsing_errors=error_message)
            time.sleep(attempt * 2)
    return ParsedDARReport(
        parsing_errors=f"Gemini call failed after {max_retries + 1} attempts. Last error: {last_exception}")


# --- Streamlit App UI and Logic ---
st.set_page_config(layout="wide", page_title="e-MCM App - GST Audit 1")
load_custom_css()

# --- Session State Initialization ---
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'username' not in st.session_state: st.session_state.username = ""
if 'role' not in st.session_state: st.session_state.role = ""
if 'audit_group_no' not in st.session_state: st.session_state.audit_group_no = None
if 'ag_current_extracted_data' not in st.session_state: st.session_state.ag_current_extracted_data = []
if 'ag_pdf_drive_url' not in st.session_state: st.session_state.ag_pdf_drive_url = None
if 'ag_validation_errors' not in st.session_state: st.session_state.ag_validation_errors = []
if 'ag_editor_data' not in st.session_state: st.session_state.ag_editor_data = pd.DataFrame()
if 'ag_current_mcm_key' not in st.session_state: st.session_state.ag_current_mcm_key = None
if 'ag_current_uploaded_file_name' not in st.session_state: st.session_state.ag_current_uploaded_file_name = None
# For Drive structure
if 'master_drive_folder_id' not in st.session_state: st.session_state.master_drive_folder_id = None
if 'mcm_periods_drive_file_id' not in st.session_state: st.session_state.mcm_periods_drive_file_id = None
if 'drive_structure_initialized' not in st.session_state: st.session_state.drive_structure_initialized = False


def login_page():
    st.markdown("<div class='page-main-title'>e-MCM App</div>", unsafe_allow_html=True)
    st.markdown("<h2 class='page-app-subtitle'>GST Audit 1 Commissionerate</h2>", unsafe_allow_html=True)

    #st.markdown("<div class='login-form-container'>", unsafe_allow_html=True)  # Container for login elements
    import base64  # Moved import here as it's only used here

    def get_image_base64_str(img_path):
        try:
            with open(img_path, "rb") as img_file:
                return base64.b64encode(img_file.read()).decode('utf-8')
        except FileNotFoundError:
            st.error(f"Logo image not found at path: {img_path}. Ensure 'logo.png' is present.")
            return None
        except Exception as e:
            st.error(f"Error reading image file {img_path}: {e}")
            return None

    image_path = "logo.png"
    base64_image = get_image_base64_str(image_path)
    if base64_image:
        image_type = os.path.splitext(image_path)[1].lower().replace(".", "") or "png"
        st.markdown(
            f"<div class='login-header'><img src='data:image/{image_type};base64,{base64_image}' alt='Logo' class='login-logo'></div>",
            unsafe_allow_html=True)
    else:
        st.markdown("<div class='login-header' style='color: red; font-weight: bold;'>[Logo Not Found]</div>",
                    unsafe_allow_html=True)

    st.markdown("<h2 class='login-header-text'>User Login</h2>", unsafe_allow_html=True)
    st.markdown("""
    <div class='app-description'>
        Welcome! This platform streamlines Draft Audit Report (DAR) collection and processing.
        PCOs manage MCM periods; Audit Groups upload DARs for AI-powered data extraction.
    </div>
    """, unsafe_allow_html=True)

    username = st.text_input("Username", key="login_username_styled", placeholder="Enter your username")
    password = st.text_input("Password", type="password", key="login_password_styled",
                             placeholder="Enter your password")

    if st.button("Login", key="login_button_styled", use_container_width=True):
        if username in USER_CREDENTIALS and USER_CREDENTIALS[username] == password:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.role = USER_ROLES[username]
            if st.session_state.role == "AuditGroup":
                st.session_state.audit_group_no = AUDIT_GROUP_NUMBERS[username]
            st.success(f"Logged in as {username} ({st.session_state.role})")
            st.session_state.drive_structure_initialized = False  # Reset for re-initialization after login
            st.rerun()
        else:
            st.error("Invalid username or password")
    st.markdown("</div>", unsafe_allow_html=True)  # Close login-form-container


# --- PCO Dashboard ---
def pco_dashboard(drive_service, sheets_service):
    st.markdown("<div class='sub-header'>Planning & Coordination Officer Dashboard</div>", unsafe_allow_html=True)
    mcm_periods = load_mcm_periods(drive_service)  # Pass drive_service

    with st.sidebar:
        st.image(
            "https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/Indian_Ministry_of_Finance_logo.svg/1200px-Indian_Ministry_of_Finance_logo.svg.png",
            width=80)
        st.markdown(f"**User:** {st.session_state.username}")
        st.markdown(f"**Role:** {st.session_state.role}")
        if st.button("Logout", key="pco_logout_styled", use_container_width=True):
            st.session_state.logged_in = False;
            st.session_state.username = "";
            st.session_state.role = ""
            st.session_state.drive_structure_initialized = False  # Reset
            st.rerun()
        st.markdown("---")

    selected_tab = option_menu(
        menu_title=None,
        options=["Create MCM Period", "Manage MCM Periods", "View Uploaded Reports", "Visualizations"],
        icons=["calendar-plus-fill", "sliders", "eye-fill", "bar-chart-fill"],
        menu_icon="gear-wide-connected", default_index=0, orientation="horizontal",
        styles={
            "container": {"padding": "5px !important", "background-color": "#e9ecef"},
            "icon": {"color": "#007bff", "font-size": "20px"},
            "nav-link": {"font-size": "16px", "text-align": "center", "margin": "0px", "--hover-color": "#d1e7fd"},
            "nav-link-selected": {"background-color": "#007bff", "color": "white"},
        })

    st.markdown("<div class='card'>", unsafe_allow_html=True)
    if selected_tab == "Create MCM Period":
        st.markdown("<h3>Create New MCM Period</h3>", unsafe_allow_html=True)
        current_year = datetime.datetime.now().year
        years = list(range(current_year - 1, current_year + 3))
        months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
                  "November", "December"]
        col1, col2 = st.columns(2)
        with col1:
            selected_year = st.selectbox("Select Year", options=years, index=years.index(current_year), key="pco_year")
        with col2:
            selected_month_name = st.selectbox("Select Month", options=months, index=datetime.datetime.now().month - 1,
                                               key="pco_month")
        selected_month_num = months.index(selected_month_name) + 1
        period_key = f"{selected_year}-{selected_month_num:02d}"

        if period_key in mcm_periods:
            st.warning(f"MCM Period for {selected_month_name} {selected_year} already exists.")
            # Links to Drive/Sheet would be part of mcm_periods[period_key]
        else:
            if st.button(f"Create MCM for {selected_month_name} {selected_year}", key="pco_create_mcm",
                         use_container_width=True):
                if not drive_service or not sheets_service or not st.session_state.get('master_drive_folder_id'):
                    st.error("Google Services or Master Drive Folder not available. Cannot create MCM period.")
                else:
                    with st.spinner("Creating Google Drive folder and Spreadsheet..."):
                        master_folder_id = st.session_state.master_drive_folder_id
                        folder_name = f"MCM_DARs_{selected_month_name}_{selected_year}"
                        spreadsheet_title = f"MCM_Audit_Paras_{selected_month_name}_{selected_year}"

                        folder_id, folder_url = create_drive_folder(drive_service, folder_name,
                                                                    parent_id=master_folder_id)
                        # Create spreadsheet also within the master folder context if possible, or link it.
                        # For simplicity, spreadsheet is created and then if drive_service can move it, it does.
                        # If drive_service can't move, it will just have general permissions.
                        sheet_id, sheet_url = create_spreadsheet(sheets_service, drive_service, spreadsheet_title,
                                                                 parent_folder_id=master_folder_id)

                        if folder_id and sheet_id:
                            mcm_periods[period_key] = {
                                "year": selected_year, "month_num": selected_month_num,
                                "month_name": selected_month_name,
                                "drive_folder_id": folder_id, "drive_folder_url": folder_url,
                                "spreadsheet_id": sheet_id, "spreadsheet_url": sheet_url, "active": True
                            }
                            if save_mcm_periods(drive_service, mcm_periods):  # Pass drive_service
                                st.success(
                                    f"Successfully created MCM period for {selected_month_name} {selected_year}!")
                                st.markdown(f"**Drive Folder:** <a href='{folder_url}' target='_blank'>Open Folder</a>",
                                            unsafe_allow_html=True)
                                st.markdown(f"**Spreadsheet:** <a href='{sheet_url}' target='_blank'>Open Sheet</a>",
                                            unsafe_allow_html=True)
                                st.balloons();
                                time.sleep(0.5);
                                st.rerun()
                            else:
                                st.error("Failed to save MCM period configuration to Drive.")
                        else:
                            st.error("Failed to create Drive folder or Spreadsheet.")

    elif selected_tab == "Manage MCM Periods":
        st.markdown("<h3>Manage Existing MCM Periods</h3>", unsafe_allow_html=True)
        if not mcm_periods:
            st.info("No MCM periods created yet.")
        else:
            sorted_periods_keys = sorted(mcm_periods.keys(), reverse=True)
            for period_key in sorted_periods_keys:
                data = mcm_periods[period_key]
                month_name_display = data.get('month_name', 'Unknown Month')
                year_display = data.get('year', 'Unknown Year')
                st.markdown(f"<h4>{month_name_display} {year_display}</h4>", unsafe_allow_html=True)
                #st.markdown(f"<h4>{data['month_name']} {data['year']}</h4>", unsafe_allow_html=True)
                col1, col2, col3, col4 = st.columns([2, 2, 1, 2])
                with col1:
                    st.markdown(f"<a href='{data.get('drive_folder_url', '#')}' target='_blank'>Open Drive Folder</a>",
                                unsafe_allow_html=True)
                with col2:
                    st.markdown(f"<a href='{data.get('spreadsheet_url', '#')}' target='_blank'>Open Spreadsheet</a>",
                                unsafe_allow_html=True)
                with col3:
                    is_active = data.get("active", False)
                    new_status = st.checkbox("Active", value=is_active, key=f"active_{period_key}_styled")
                    if new_status != is_active:
                        mcm_periods[period_key]["active"] = new_status
                        if save_mcm_periods(drive_service, mcm_periods):  # Pass drive_service
                            #st.success(f"Status for {data['month_name']} {data['year']} updated.");
                            
                            # NEW FIXED LINES:
                            month_name_succ = data.get('month_name', 'Unknown Period') # Provide a default
                            year_succ = data.get('year', '') # Default to empty if year is also part of the issue
                            st.success(f"Status for {month_name_succ} {year_succ} updated.");
                            st.rerun()
                        else:
                            st.error("Failed to save updated MCM period status to Drive.")
                with col4:
                    if st.button("Delete Period Record", key=f"delete_mcm_{period_key}", type="secondary"):
                        st.session_state.period_to_delete = period_key
                        st.session_state.show_delete_confirm = True
                        st.rerun()
                st.markdown("---")

            if st.session_state.get('show_delete_confirm', False) and st.session_state.get('period_to_delete'):
                period_key_to_delete = st.session_state.period_to_delete
                period_data_to_delete = mcm_periods.get(period_key_to_delete, {})
                with st.form(key=f"delete_confirm_form_{period_key_to_delete}"):
                    st.warning(
                        f"Are you sure you want to delete the MCM period record for **{period_data_to_delete.get('month_name')} {period_data_to_delete.get('year')}** from this application?")
                    st.caption(
                        "This action only removes the period from the app's tracking (from the `{MCM_PERIODS_FILENAME_ON_DRIVE}` file on Google Drive). It **does NOT delete** the actual Google Drive folder or the Google Spreadsheet.")
                    pco_password_confirm = st.text_input("Enter your PCO password:", type="password",
                                                         key=f"pco_pass_conf_{period_key_to_delete}")
                    c1, c2 = st.columns(2)
                    with c1:
                        submitted_delete = st.form_submit_button("Yes, Delete Record", use_container_width=True)
                    with c2:
                        if st.form_submit_button("Cancel", type="secondary", use_container_width=True):
                            st.session_state.show_delete_confirm = False;
                            st.session_state.period_to_delete = None;
                            st.rerun()
                    if submitted_delete:
                        if pco_password_confirm == USER_CREDENTIALS.get("planning_officer"):
                            del mcm_periods[period_key_to_delete]
                            if save_mcm_periods(drive_service, mcm_periods):  # Pass drive_service
                                st.success(
                                    f"MCM record for {period_data_to_delete.get('month_name')} {period_data_to_delete.get('year')} deleted.");
                            else:
                                st.error("Failed to save changes to Drive after deleting record locally.")
                            st.session_state.show_delete_confirm = False;
                            st.session_state.period_to_delete = None;
                            st.rerun()
                        else:
                            st.error("Incorrect password.")

    elif selected_tab == "View Uploaded Reports":
        st.markdown("<h3>View Uploaded Reports Summary</h3>", unsafe_allow_html=True)
        active_periods = {k: v for k, v in mcm_periods.items()}  # Filter for active if needed, or show all
        if not active_periods:
            st.info("No MCM periods to view reports for.")
        else:
            period_options = [
                 f"{p.get('month_name')} {p.get('year')}"
                 for k, p in sorted(active_periods.items(), key=lambda item: item[0], reverse=True)
                 if p.get('month_name') and p.get('year') # Only include if both month_name and year exist
             ]
            if not period_options:
                 st.warning("No valid MCM periods with complete month and year information found to display options.")
            # period_options = [f"{p['month_name']} {p['year']}" for k, p in
            #                   sorted(active_periods.items(), key=lambda item: item[0], reverse=True)]
            selected_period_display = st.selectbox("Select MCM Period", options=period_options,
                                                   key="pco_view_reports_period")
            if selected_period_display:
                # selected_period_key = next((k for k, p in active_periods.items() if
                #                             f"{p['month_name']} {p['year']}" == selected_period_display), None)
                selected_period_key = next((k for k, p in active_periods.items() if
                            p.get('month_name') and p.get('year') and # Ensure keys exist in p
                            f"{p.get('month_name')} {p.get('year')}" == selected_period_display), None)
                if selected_period_key and sheets_service:
                    sheet_id = mcm_periods[selected_period_key]['spreadsheet_id']
                    with st.spinner("Loading data from Google Sheet..."):
                        df = read_from_spreadsheet(sheets_service, sheet_id)
                    if not df.empty:
                        # Display logic from original code
                        st.markdown("<h4>Summary of Uploads:</h4>", unsafe_allow_html=True)
                        if 'Audit Group Number' in df.columns:
                            try:
                                df['Audit Group Number'] = pd.to_numeric(df['Audit Group Number'], errors='coerce')
                                df.dropna(subset=['Audit Group Number'], inplace=True)

                                dars_per_group = df.groupby('Audit Group Number')['DAR PDF URL'].nunique().reset_index(
                                    name='DARs Uploaded')
                                st.write("**DARs Uploaded per Audit Group:**");
                                st.dataframe(dars_per_group, use_container_width=True)
                                paras_per_group = df.groupby('Audit Group Number').size().reset_index(
                                    name='Total Para Entries')
                                st.write("**Total Para Entries per Audit Group:**");
                                st.dataframe(paras_per_group, use_container_width=True)
                                st.markdown("<h4>Detailed Data:</h4>", unsafe_allow_html=True);
                                st.dataframe(df, use_container_width=True)
                            except Exception as e:
                                st.error(f"Error processing summary: {e}"); st.dataframe(df, use_container_width=True)
                        else:
                            st.warning("Missing 'Audit Group Number' column."); st.dataframe(df,
                                                                                             use_container_width=True)
                    else:
                        st.info(f"No data in spreadsheet for {selected_period_display}.")
                elif not sheets_service:
                    st.error("Google Sheets service not available.")

    elif selected_tab == "Visualizations":  # MODIFIED
        st.markdown("<h3>Data Visualizations</h3>", unsafe_allow_html=True)
        all_mcm_periods = mcm_periods  # Use all periods for selection
        if not all_mcm_periods:
            st.info("No MCM periods to visualize data from.")
        else:
            viz_period_options = [
                f"{p.get('month_name')} {p.get('year')}"
                for k, p in sorted(all_mcm_periods.items(), key=lambda item: item[0], reverse=True)
                if p.get('month_name') and p.get('year') # Only include if both month_name and year exist
            ]
            if not viz_period_options:
                st.warning("No valid MCM periods with complete month and year information found for visualization options.")
            # viz_period_options = [f"{p['month_name']} {p['year']}" for k, p in
            #                       sorted(all_mcm_periods.items(), key=lambda item: item[0], reverse=True)]
            selected_viz_period_display = st.selectbox("Select MCM Period for Visualization",
                                                       options=viz_period_options, key="pco_viz_period")
            if selected_viz_period_display and sheets_service:
                # selected_viz_period_key = next((k for k, p in all_mcm_periods.items() if
                #                                 f"{p['month_name']} {p['year']}" == selected_viz_period_display), None)
                selected_viz_period_key = next((k for k, p in all_mcm_periods.items() if
                                p.get('month_name') and p.get('year') and # Ensure keys exist in p
                                f"{p.get('month_name')} {p.get('year')}" == selected_viz_period_display), None)
                if selected_viz_period_key:
                    sheet_id_viz = all_mcm_periods[selected_viz_period_key]['spreadsheet_id']
                    with st.spinner("Loading data for visualizations..."):
                        df_viz = read_from_spreadsheet(sheets_service, sheet_id_viz)

                    if not df_viz.empty:
                        amount_cols = ['Total Amount Detected (Overall Rs)', 'Total Amount Recovered (Overall Rs)',
                                       'Revenue Involved (Lakhs Rs)', 'Revenue Recovered (Lakhs Rs)']
                        for col in amount_cols:
                            if col in df_viz.columns: df_viz[col] = pd.to_numeric(df_viz[col], errors='coerce').fillna(
                                0)
                        if 'Audit Group Number' in df_viz.columns:
                            df_viz['Audit Group Number'] = pd.to_numeric(df_viz['Audit Group Number'],
                                                                         errors='coerce').fillna(0).astype(int)

                        st.markdown("---")
                        st.markdown("<h4>Group-wise Performance</h4>", unsafe_allow_html=True)

                        # Top 5 Groups by Total Detection
                        if 'Total Amount Detected (Overall Rs)' in df_viz.columns and 'Audit Group Number' in df_viz.columns:
                            detection_data = df_viz.groupby('Audit Group Number')[
                                'Total Amount Detected (Overall Rs)'].sum().reset_index()
                            detection_data = detection_data.sort_values(by='Total Amount Detected (Overall Rs)',
                                                                        ascending=False).nlargest(5,
                                                                                                  'Total Amount Detected (Overall Rs)')
                            if not detection_data.empty:
                                st.write("**Top 5 Groups by Total Detection Amount (Rs):**")
                                fig = px.bar(detection_data, x='Audit Group Number',
                                             y='Total Amount Detected (Overall Rs)', text_auto=True,
                                             labels={'Total Amount Detected (Overall Rs)': 'Total Detection (Rs)',
                                                     'Audit Group Number': '<b>Audit Group</b>'})
                                fig.update_layout(xaxis_title_font_size=14, yaxis_title_font_size=14,
                                                  xaxis_tickfont_size=12, yaxis_tickfont_size=12,
                                                  xaxis={'categoryorder': 'total descending'})
                                fig.update_traces(textposition='outside', marker_color='indianred')
                                st.plotly_chart(fig, use_container_width=True)
                            else:
                                st.info("Not enough data for 'Top Detection Groups' chart.")

                        # Top 5 Groups by Total Realisation
                        if 'Total Amount Recovered (Overall Rs)' in df_viz.columns and 'Audit Group Number' in df_viz.columns:
                            recovery_data = df_viz.groupby('Audit Group Number')[
                                'Total Amount Recovered (Overall Rs)'].sum().reset_index()
                            recovery_data = recovery_data.sort_values(by='Total Amount Recovered (Overall Rs)',
                                                                      ascending=False).nlargest(5,
                                                                                                'Total Amount Recovered (Overall Rs)')
                            if not recovery_data.empty:
                                st.write("**Top 5 Groups by Total Realisation Amount (Rs):**")
                                fig = px.bar(recovery_data, x='Audit Group Number',
                                             y='Total Amount Recovered (Overall Rs)', text_auto=True,
                                             labels={'Total Amount Recovered (Overall Rs)': 'Total Realisation (Rs)',
                                                     'Audit Group Number': '<b>Audit Group</b>'})
                                fig.update_layout(xaxis_title_font_size=14, yaxis_title_font_size=14,
                                                  xaxis_tickfont_size=12, yaxis_tickfont_size=12,
                                                  xaxis={'categoryorder': 'total descending'})
                                fig.update_traces(textposition='outside', marker_color='lightseagreen')
                                st.plotly_chart(fig, use_container_width=True)
                            else:
                                st.info("Not enough data for 'Top Realisation Groups' chart.")

                        # Top 5 Groups by Recovery/Detection Ratio
                        if 'Total Amount Detected (Overall Rs)' in df_viz.columns and 'Total Amount Recovered (Overall Rs)' in df_viz.columns and 'Audit Group Number' in df_viz.columns:
                            group_summary = df_viz.groupby('Audit Group Number').agg(
                                Total_Detected=('Total Amount Detected (Overall Rs)', 'sum'),
                                Total_Recovered=('Total Amount Recovered (Overall Rs)', 'sum')
                            ).reset_index()
                            group_summary['Recovery_Ratio'] = group_summary.apply(
                                lambda row: (row['Total_Recovered'] / row['Total_Detected']) * 100 if pd.notna(
                                    row['Total_Detected']) and row['Total_Detected'] > 0 and pd.notna(
                                    row['Total_Recovered']) else 0, axis=1
                            )
                            ratio_data = group_summary.sort_values(by='Recovery_Ratio', ascending=False).nlargest(5,
                                                                                                                  'Recovery_Ratio')
                            if not ratio_data.empty:
                                st.write("**Top 5 Groups by Recovery/Detection Ratio (%):**")
                                fig = px.bar(ratio_data, x='Audit Group Number', y='Recovery_Ratio', text_auto=True,
                                             labels={'Recovery_Ratio': 'Recovery Ratio (%)',
                                                     'Audit Group Number': '<b>Audit Group</b>'})
                                fig.update_layout(xaxis_title_font_size=14, yaxis_title_font_size=14,
                                                  xaxis_tickfont_size=12, yaxis_tickfont_size=12,
                                                  xaxis={'categoryorder': 'total descending'})
                                fig.update_traces(textposition='outside', marker_color='mediumpurple')
                                st.plotly_chart(fig, use_container_width=True)
                            else:
                                st.info("Not enough data for 'Top Recovery Ratio Groups' chart.")

                        st.markdown("---")
                        st.markdown("<h4>Para-wise Performance</h4>", unsafe_allow_html=True)
                        num_paras_to_show = st.number_input("Select N for Top N Paras:", min_value=1, max_value=20,
                                                            value=5, step=1, key="top_n_paras_viz")
                        df_paras_only = df_viz[df_viz['Audit Para Number'].notna() & (
                            ~df_viz['Audit Para Heading'].isin(
                                ["N/A - Header Info Only (Add Paras Manually)", "Manual Entry Required",
                                 "Manual Entry - PDF Error", "Manual Entry - PDF Upload Failed"]))]

                        if 'Revenue Involved (Lakhs Rs)' in df_paras_only.columns:
                            top_detection_paras = df_paras_only.nlargest(num_paras_to_show,
                                                                         'Revenue Involved (Lakhs Rs)')
                            if not top_detection_paras.empty:
                                st.write(f"**Top {num_paras_to_show} Detection Paras (by Revenue Involved):**")
                                st.dataframe(top_detection_paras[
                                                 ['Audit Group Number', 'Trade Name', 'Audit Para Number',
                                                  'Audit Para Heading', 'Revenue Involved (Lakhs Rs)']],
                                             use_container_width=True)
                            else:
                                st.info("Not enough data for 'Top Detection Paras' list.")
                        if 'Revenue Recovered (Lakhs Rs)' in df_paras_only.columns:
                            top_recovery_paras = df_paras_only.nlargest(num_paras_to_show,
                                                                        'Revenue Recovered (Lakhs Rs)')
                            if not top_recovery_paras.empty:
                                st.write(f"**Top {num_paras_to_show} Realisation Paras (by Revenue Recovered):**")
                                st.dataframe(top_recovery_paras[
                                                 ['Audit Group Number', 'Trade Name', 'Audit Para Number',
                                                  'Audit Para Heading', 'Revenue Recovered (Lakhs Rs)']],
                                             use_container_width=True)
                            else:
                                st.info("Not enough data for 'Top Realisation Paras' list.")
                    else:
                        st.info(f"No data in spreadsheet for {selected_viz_period_display} to visualize.")
                elif not sheets_service:
                    st.error("Google Sheets service not available.")
    st.markdown("</div>", unsafe_allow_html=True)


# --- Validation Function ---
MANDATORY_FIELDS_FOR_SHEET = {
    "audit_group_number": "Audit Group Number", "gstin": "GSTIN", "trade_name": "Trade Name", "category": "Category",
    "total_amount_detected_overall_rs": "Total Amount Detected (Overall Rs)",
    "total_amount_recovered_overall_rs": "Total Amount Recovered (Overall Rs)",
    "audit_para_number": "Audit Para Number", "audit_para_heading": "Audit Para Heading",
    "revenue_involved_lakhs_rs": "Revenue Involved (Lakhs Rs)",
    "revenue_recovered_lakhs_rs": "Revenue Recovered (Lakhs Rs)"
}
VALID_CATEGORIES = ["Large", "Medium", "Small"]


def validate_data_for_sheet(data_df_to_validate):
    validation_errors = []
    if data_df_to_validate.empty: return ["No data to validate."]
    for index, row in data_df_to_validate.iterrows():
        row_display_id = f"Row {index + 1} (Para: {row.get('audit_para_number', 'N/A')})"
        for field_key, field_name in MANDATORY_FIELDS_FOR_SHEET.items():
            value = row.get(field_key)
            is_missing = value is None or (isinstance(value, str) and not value.strip()) or pd.isna(value)
            if is_missing:
                if field_key in ["audit_para_number", "audit_para_heading", "revenue_involved_lakhs_rs",
                                 "revenue_recovered_lakhs_rs"] and \
                        row.get('audit_para_heading', "").startswith("N/A - Header Info Only") and pd.isna(
                    row.get('audit_para_number')):
                    continue
                validation_errors.append(f"{row_display_id}: '{field_name}' is missing or empty.")
        category_val = row.get('category')
        if pd.notna(category_val) and category_val.strip() and category_val not in VALID_CATEGORIES:
            validation_errors.append(
                f"{row_display_id}: 'Category' ('{category_val}') is invalid. Must be one of {VALID_CATEGORIES}.")
        elif (pd.isna(category_val) or not str(category_val).strip()) and "category" in MANDATORY_FIELDS_FOR_SHEET:
            validation_errors.append(f"{row_display_id}: 'Category' is missing.")
    if 'trade_name' in data_df_to_validate.columns and 'category' in data_df_to_validate.columns:
        trade_name_categories = {}
        for index, row in data_df_to_validate.iterrows():
            trade_name, category = row.get('trade_name'), row.get('category')
            if pd.notna(trade_name) and str(trade_name).strip() and pd.notna(category) and str(
                    category).strip() and category in VALID_CATEGORIES:
                trade_name_categories.setdefault(trade_name, set()).add(category)
        for tn, cats in trade_name_categories.items():
            if len(cats) > 1: validation_errors.append(
                f"Consistency Error: Trade Name '{tn}' has multiple categories: {', '.join(sorted(list(cats)))}.")
    return sorted(list(set(validation_errors)))


# --- Audit Group Dashboard ---
def audit_group_dashboard(drive_service, sheets_service):
    st.markdown(f"<div class='sub-header'>Audit Group {st.session_state.audit_group_no} Dashboard</div>",
                unsafe_allow_html=True)
    mcm_periods = load_mcm_periods(drive_service)  # Pass drive_service
    active_periods = {k: v for k, v in mcm_periods.items() if v.get("active")}

    with st.sidebar:
        st.image(
            "https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/Indian_Ministry_of_Finance_logo.svg/1200px-Indian_Ministry_of_Finance_logo.svg.png",
            width=80)
        st.markdown(f"**User:** {st.session_state.username}<br>**Group No:** {st.session_state.audit_group_no}",
                    unsafe_allow_html=True)
        if st.button("Logout", key="ag_logout_styled", use_container_width=True):
            for key_to_del in ['ag_current_extracted_data', 'ag_pdf_drive_url', 'ag_validation_errors',
                               'ag_editor_data', 'ag_current_mcm_key', 'ag_current_uploaded_file_name',
                               'ag_row_to_delete_details', 'ag_show_delete_confirm', 'drive_structure_initialized']:
                if key_to_del in st.session_state: del st.session_state[key_to_del]
            st.session_state.logged_in = False;
            st.session_state.username = "";
            st.session_state.role = "";
            st.session_state.audit_group_no = None
            st.rerun()
        st.markdown("---")

    selected_tab = option_menu(
        menu_title=None, options=["Upload DAR for MCM", "View My Uploaded DARs", "Delete My DAR Entries"],
        icons=["cloud-upload-fill", "eye-fill", "trash2-fill"], menu_icon="person-workspace", default_index=0,
        orientation="horizontal",
        styles={
            "container": {"padding": "5px !important", "background-color": "#e9ecef"},
            "icon": {"color": "#28a745", "font-size": "20px"},
            "nav-link": {"font-size": "16px", "text-align": "center", "margin": "0px", "--hover-color": "#d4edda"},
            "nav-link-selected": {"background-color": "#28a745", "color": "white"},
        })
    st.markdown("<div class='card'>", unsafe_allow_html=True)

    if selected_tab == "Upload DAR for MCM":
        st.markdown("<h3>Upload DAR PDF for MCM Period</h3>", unsafe_allow_html=True)
        if not active_periods:
            st.warning("No active MCM periods. Contact Planning Officer.")
        else:
            # period_options = {f"{p['month_name']} {p['year']}": k for k, p in
            #                   sorted(active_periods.items(), reverse=True)}
            period_options = {
                                 f"{p.get('month_name')} {p.get('year')}": k
                                 for k, p in sorted(active_periods.items(), reverse=True)
                                 if p.get('month_name') and p.get('year') # Only include if both keys exist
                             }
            if not period_options and active_periods: # If active_periods was not empty but options are
                st.warning("Some active MCM periods have incomplete data (missing month/year) and are not shown as options.")
            selected_period_display = st.selectbox("Select Active MCM Period", options=list(period_options.keys()),
                                                   key="ag_select_mcm_upload")
            if selected_period_display:
                selected_mcm_key = period_options[selected_period_display]
                mcm_info = mcm_periods[selected_mcm_key]
                if st.session_state.get('ag_current_mcm_key') != selected_mcm_key:  # Reset on MCM change
                    st.session_state.ag_current_extracted_data = [];
                    st.session_state.ag_pdf_drive_url = None
                    st.session_state.ag_validation_errors = [];
                    st.session_state.ag_editor_data = pd.DataFrame()
                    st.session_state.ag_current_mcm_key = selected_mcm_key;
                    st.session_state.ag_current_uploaded_file_name = None
                st.info(f"Uploading for: {mcm_info['month_name']} {mcm_info['year']}")
                uploaded_dar_file = st.file_uploader("Choose DAR PDF", type="pdf",
                                                     key=f"dar_upload_ag_{selected_mcm_key}_{st.session_state.get('uploader_key_suffix', 0)}")

                if uploaded_dar_file:
                    if st.session_state.get(
                            'ag_current_uploaded_file_name') != uploaded_dar_file.name:  # Reset on new file
                        st.session_state.ag_current_extracted_data = [];
                        st.session_state.ag_pdf_drive_url = None
                        st.session_state.ag_validation_errors = [];
                        st.session_state.ag_editor_data = pd.DataFrame()
                        st.session_state.ag_current_uploaded_file_name = uploaded_dar_file.name

                    if st.button("Extract Data from PDF", key=f"extract_ag_{selected_mcm_key}",
                                 use_container_width=True):
                        st.session_state.ag_validation_errors = []
                        with st.spinner("Processing PDF & AI extraction..."):
                            dar_pdf_bytes = uploaded_dar_file.getvalue()
                            dar_filename_on_drive = f"AG{st.session_state.audit_group_no}_{uploaded_dar_file.name}"
                            st.session_state.ag_pdf_drive_url = None
                            pdf_drive_id, pdf_drive_url_temp = upload_to_drive(drive_service, dar_pdf_bytes,
                                                                               mcm_info['drive_folder_id'],
                                                                               dar_filename_on_drive)

                            if not pdf_drive_id:
                                st.error("Failed to upload PDF to Drive.");
                                st.session_state.ag_editor_data = pd.DataFrame([{
                                                                                    "audit_group_number": st.session_state.audit_group_no,
                                                                                    "audit_para_heading": "Manual Entry - PDF Upload Failed"}])
                            else:
                                st.session_state.ag_pdf_drive_url = pdf_drive_url_temp
                                st.success(f"DAR PDF on Drive: [Link]({st.session_state.ag_pdf_drive_url})")
                                preprocessed_text = preprocess_pdf_text(BytesIO(dar_pdf_bytes))
                                if preprocessed_text.startswith("Error"):
                                    st.error(f"PDF Preprocessing Error: {preprocessed_text}");
                                    default_row = {"audit_group_number": st.session_state.audit_group_no,
                                                   "audit_para_heading": "Manual Entry - PDF Error"}
                                    st.session_state.ag_editor_data = pd.DataFrame([default_row])
                                else:
                                    parsed_report_obj = get_structured_data_with_gemini(YOUR_GEMINI_API_KEY,
                                                                                        preprocessed_text)
                                    temp_list = []
                                    ai_failed = True
                                    if parsed_report_obj.parsing_errors: st.warning(
                                        f"AI Parsing Issues: {parsed_report_obj.parsing_errors}")
                                    if parsed_report_obj and parsed_report_obj.header:
                                        h = parsed_report_obj.header;
                                        ai_failed = False
                                        if parsed_report_obj.audit_paras:
                                            for p in parsed_report_obj.audit_paras: temp_list.append(
                                                {"audit_group_number": st.session_state.audit_group_no,
                                                 "gstin": h.gstin, "trade_name": h.trade_name, "category": h.category,
                                                 "total_amount_detected_overall_rs": h.total_amount_detected_overall_rs,
                                                 "total_amount_recovered_overall_rs": h.total_amount_recovered_overall_rs,
                                                 "audit_para_number": p.audit_para_number,
                                                 "audit_para_heading": p.audit_para_heading,
                                                 "revenue_involved_lakhs_rs": p.revenue_involved_lakhs_rs,
                                                 "revenue_recovered_lakhs_rs": p.revenue_recovered_lakhs_rs})
                                        elif h.trade_name:
                                            temp_list.append({"audit_group_number": st.session_state.audit_group_no,
                                                              "gstin": h.gstin, "trade_name": h.trade_name,
                                                              "category": h.category,
                                                              "total_amount_detected_overall_rs": h.total_amount_detected_overall_rs,
                                                              "total_amount_recovered_overall_rs": h.total_amount_recovered_overall_rs,
                                                              "audit_para_heading": "N/A - Header Info Only (Add Paras Manually)"})
                                        else:
                                            st.error("AI failed to extract key header info."); ai_failed = True
                                    if ai_failed or not temp_list:
                                        st.warning("AI extraction failed or yielded no data. Please fill manually.")
                                        st.session_state.ag_editor_data = pd.DataFrame([{
                                                                                            "audit_group_number": st.session_state.audit_group_no,
                                                                                            "audit_para_heading": "Manual Entry Required"}])
                                    else:
                                        st.session_state.ag_editor_data = pd.DataFrame(temp_list); st.info(
                                            "Data extracted. Review & edit below.")
                if not isinstance(st.session_state.get('ag_editor_data'),
                                  pd.DataFrame): st.session_state.ag_editor_data = pd.DataFrame()
                if uploaded_dar_file and st.session_state.ag_editor_data.empty and st.session_state.get(
                        'ag_pdf_drive_url'):
                    st.warning("AI couldn't extract data or none loaded. Template row provided.")
                    st.session_state.ag_editor_data = pd.DataFrame(
                        [{"audit_group_number": st.session_state.audit_group_no, "audit_para_heading": "Manual Entry"}])

                if not st.session_state.ag_editor_data.empty:
                    st.markdown("<h4>Review and Edit Extracted Data:</h4>", unsafe_allow_html=True)
                    df_to_edit_ag = st.session_state.ag_editor_data.copy();
                    df_to_edit_ag["audit_group_number"] = st.session_state.audit_group_no
                    col_order = ["audit_group_number", "gstin", "trade_name", "category",
                                 "total_amount_detected_overall_rs", "total_amount_recovered_overall_rs",
                                 "audit_para_number", "audit_para_heading", "revenue_involved_lakhs_rs",
                                 "revenue_recovered_lakhs_rs"]
                    for col in col_order:
                        if col not in df_to_edit_ag.columns: df_to_edit_ag[col] = None
                    col_config = {
                        "audit_group_number": st.column_config.NumberColumn("Audit Group", disabled=True),
                        "gstin": st.column_config.TextColumn("GSTIN"),
                        "trade_name": st.column_config.TextColumn("Trade Name"),
                        "category": st.column_config.SelectboxColumn("Category", options=VALID_CATEGORIES),
                        "total_amount_detected_overall_rs": st.column_config.NumberColumn("Total Detected (Rs)",
                                                                                          format="%.2f"),
                        "total_amount_recovered_overall_rs": st.column_config.NumberColumn("Total Recovered (Rs)",
                                                                                           format="%.2f"),
                        "audit_para_number": st.column_config.NumberColumn("Para No.", format="%d"),
                        "audit_para_heading": st.column_config.TextColumn("Para Heading", width="xlarge"),
                        "revenue_involved_lakhs_rs": st.column_config.NumberColumn("Rev. Involved (Lakhs)",
                                                                                   format="%.2f"),
                        "revenue_recovered_lakhs_rs": st.column_config.NumberColumn("Rev. Recovered (Lakhs)",
                                                                                    format="%.2f")}
                    editor_key = f"ag_editor_{selected_mcm_key}_{st.session_state.ag_current_uploaded_file_name or 'no_file'}"
                    edited_df = st.data_editor(df_to_edit_ag.reindex(columns=col_order), column_config=col_config,
                                               num_rows="dynamic", key=editor_key, use_container_width=True, height=400)

                    if st.button("Validate and Submit to MCM Sheet", key=f"submit_ag_{selected_mcm_key}",
                                 use_container_width=True):
                        current_data = pd.DataFrame(edited_df);
                        current_data["audit_group_number"] = st.session_state.audit_group_no
                        val_errors = validate_data_for_sheet(current_data);
                        st.session_state.ag_validation_errors = val_errors
                        if not val_errors:
                            if not st.session_state.ag_pdf_drive_url:
                                st.error("PDF Drive URL missing. Re-extract data.")
                            else:
                                with st.spinner("Submitting to Google Sheet..."):
                                    rows_to_append = []
                                    created_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    for _, row in current_data.iterrows(): rows_to_append.append(
                                        [row.get(c) for c in col_order] + [st.session_state.ag_pdf_drive_url,
                                                                           created_date])
                                    if rows_to_append:
                                        if append_to_spreadsheet(sheets_service, mcm_info['spreadsheet_id'],
                                                                 rows_to_append):
                                            st.success(
                                                f"Data for '{st.session_state.ag_current_uploaded_file_name}' submitted!");
                                            st.balloons();
                                            time.sleep(0.5)
                                            st.session_state.ag_current_extracted_data = [];
                                            st.session_state.ag_pdf_drive_url = None;
                                            st.session_state.ag_editor_data = pd.DataFrame();
                                            st.session_state.ag_current_uploaded_file_name = None
                                            st.session_state.uploader_key_suffix = st.session_state.get(
                                                'uploader_key_suffix', 0) + 1;
                                            st.rerun()
                                        else:
                                            st.error("Failed to append to Google Sheet.")
                                    else:
                                        st.error("No data to submit after validation.")
                        else:
                            st.error("Validation Failed! Correct errors below.")
                if st.session_state.get('ag_validation_errors'):
                    st.markdown("---");
                    st.subheader("⚠️ Validation Errors:");
                    for err in st.session_state.ag_validation_errors: st.warning(err)

    elif selected_tab == "View My Uploaded DARs":
        st.markdown("<h3>My Uploaded DARs</h3>", unsafe_allow_html=True)
        if not mcm_periods:
            st.info("No MCM periods by PCO yet.")
        else:
            # all_period_options = {f"{p['month_name']} {p['year']}": k for k, p in
            #                       sorted(mcm_periods.items(), key=lambda item: item[0], reverse=True)}
            all_period_options = {
                                     f"{p.get('month_name')} {p.get('year')}": k
                                     for k, p in sorted(mcm_periods.items(), key=lambda item: item[0], reverse=True)
                                     if p.get('month_name') and p.get('year') # Only include if both keys exist
                                 }
            if not all_period_options and mcm_periods:
                st.warning("Some MCM periods have incomplete data (missing month/year) and are not shown as options for viewing.")
            if not all_period_options:
                st.info("No MCM periods found.")
            else:
                selected_view_period_display = st.selectbox("Select MCM Period",
                                                            options=list(all_period_options.keys()),
                                                            key="ag_view_my_dars_period")
                if selected_view_period_display and sheets_service:
                    selected_view_period_key = all_period_options[selected_view_period_display]
                    sheet_id = mcm_periods[selected_view_period_key]['spreadsheet_id']
                    with st.spinner("Loading your uploads..."):
                        df_all = read_from_spreadsheet(sheets_service, sheet_id)
                    if not df_all.empty and 'Audit Group Number' in df_all.columns:
                        df_all['Audit Group Number'] = df_all['Audit Group Number'].astype(str)
                        my_uploads_df = df_all[df_all['Audit Group Number'] == str(st.session_state.audit_group_no)]
                        if not my_uploads_df.empty:
                            st.markdown(f"<h4>Your Uploads for {selected_view_period_display}:</h4>",
                                        unsafe_allow_html=True)
                            df_display = my_uploads_df.copy()
                            if 'DAR PDF URL' in df_display.columns: df_display['DAR PDF URL'] = df_display[
                                'DAR PDF URL'].apply(
                                lambda x: f'<a href="{x}" target="_blank">View PDF</a>' if pd.notna(x) and str(
                                    x).startswith("http") else "No Link")
                            view_cols = ["Trade Name", "Category", "Audit Para Number", "Audit Para Heading",
                                         "DAR PDF URL", "Record Created Date"]
                            st.markdown(df_display[view_cols].to_html(escape=False, index=False),
                                        unsafe_allow_html=True)
                        else:
                            st.info(f"No DARs uploaded by you for {selected_view_period_display}.")
                    elif df_all.empty:
                        st.info(f"No data in MCM sheet for {selected_view_period_display}.")
                    else:
                        st.warning("Spreadsheet missing 'Audit Group Number' column.")
                elif not sheets_service:
                    st.error("Google Sheets service not available.")

    elif selected_tab == "Delete My DAR Entries":
        st.markdown("<h3>Delete My Uploaded DAR Entries</h3>", unsafe_allow_html=True)
        st.info("Select MCM period to view entries. Deletion removes entry from Google Sheet; PDF on Drive remains.")
        if not mcm_periods:
            st.info("No MCM periods created yet.")
        else:
            # all_period_options_del = {f"{p['month_name']} {p['year']}": k for k, p in
            #                           sorted(mcm_periods.items(), key=lambda item: item[0], reverse=True)}
            all_period_options_del = {
                                         f"{p.get('month_name')} {p.get('year')}": k
                                         for k, p in sorted(mcm_periods.items(), key=lambda item: item[0], reverse=True)
                                         if p.get('month_name') and p.get('year') # Only include if both keys exist
                                     }
            if not all_period_options_del and mcm_periods:
                st.warning("Some MCM periods have incomplete data (missing month/year) and are not shown as options for deletion.")
            selected_del_period_display = st.selectbox("Select MCM Period", options=list(all_period_options_del.keys()),
                                                       key="ag_del_dars_period")
            if selected_del_period_display and sheets_service:
                selected_del_period_key = all_period_options_del[selected_del_period_display]
                sheet_id_to_manage = mcm_periods[selected_del_period_key]['spreadsheet_id']
                first_sheet_gid = 0  # Default
                try:
                    meta = sheets_service.spreadsheets().get(spreadsheetId=sheet_id_to_manage).execute()
                    first_sheet_gid = meta.get('sheets', [{}])[0].get('properties', {}).get('sheetId', 0)
                except Exception as e_gid:
                    st.error(f"Could not fetch sheet GID: {e_gid}")

                with st.spinner("Loading your uploads..."):
                    df_all_del = read_from_spreadsheet(sheets_service, sheet_id_to_manage)
                if not df_all_del.empty and 'Audit Group Number' in df_all_del.columns:
                    df_all_del['Audit Group Number'] = df_all_del['Audit Group Number'].astype(str)
                    my_uploads_df_del = df_all_del[
                        df_all_del['Audit Group Number'] == str(st.session_state.audit_group_no)].copy()
                    my_uploads_df_del.reset_index(inplace=True);
                    my_uploads_df_del.rename(columns={'index': 'original_df_index'}, inplace=True)

                    if not my_uploads_df_del.empty:
                        st.markdown(f"<h4>Your Uploads in {selected_del_period_display} (Select to delete):</h4>",
                                    unsafe_allow_html=True)
                        options_for_del = ["--Select an entry--"]
                        st.session_state.ag_deletable_map = {}
                        for idx, row in my_uploads_df_del.iterrows():
                            ident_str = f"Entry (TN: {str(row.get('Trade Name', 'N/A'))[:20]}..., Para: {row.get('Audit Para Number', 'N/A')}, Date: {row.get('Record Created Date', 'N/A')})"
                            options_for_del.append(ident_str)
                            st.session_state.ag_deletable_map[ident_str] = {k: str(row.get(k)) for k in
                                                                            ["Trade Name", "Audit Para Number",
                                                                             "Record Created Date", "DAR PDF URL"]}
                        selected_entry_del_display = st.selectbox("Select Entry to Delete:", options_for_del,
                                                                  key=f"del_sel_{selected_del_period_key}")

                        if selected_entry_del_display != "--Select an entry--":
                            row_ident_data = st.session_state.ag_deletable_map.get(selected_entry_del_display)
                            if row_ident_data:
                                st.warning(
                                    f"Selected to delete: **{row_ident_data.get('trade_name')} - Para {row_ident_data.get('audit_para_number')}** (Uploaded: {row_ident_data.get('record_created_date')})")
                                with st.form(key=f"del_ag_form_{selected_entry_del_display.replace(' ', '_')}"):
                                    ag_pass = st.text_input("Your Password:", type="password",
                                                            key=f"ag_pass_del_{selected_entry_del_display.replace(' ', '_')}")
                                    submitted_del = st.form_submit_button("Confirm Deletion")
                                    if submitted_del:
                                        if ag_pass == USER_CREDENTIALS.get(st.session_state.username):
                                            current_sheet_df = read_from_spreadsheet(sheets_service,
                                                                                     sheet_id_to_manage)  # Re-fetch
                                            if not current_sheet_df.empty:
                                                indices_to_del_sheet = []
                                                for sheet_idx, sheet_row in current_sheet_df.iterrows():
                                                    match = all([
                                                        str(sheet_row.get('Audit Group Number')) == str(
                                                            st.session_state.audit_group_no),
                                                        str(sheet_row.get('Trade Name')) == row_ident_data.get(
                                                            'trade_name'),
                                                        str(sheet_row.get('Audit Para Number')) == row_ident_data.get(
                                                            'audit_para_number'),
                                                        str(sheet_row.get('Record Created Date')) == row_ident_data.get(
                                                            'record_created_date'),
                                                        str(sheet_row.get('DAR PDF URL')) == row_ident_data.get(
                                                            'dar_pdf_url')
                                                    ])
                                                    if match: indices_to_del_sheet.append(sheet_idx)
                                                if indices_to_del_sheet:
                                                    if delete_spreadsheet_rows(sheets_service, sheet_id_to_manage,
                                                                               first_sheet_gid, indices_to_del_sheet):
                                                        st.success(
                                                            f"Entry for '{row_ident_data.get('trade_name')}' deleted.");
                                                        time.sleep(0.5);
                                                        st.rerun()
                                                    else:
                                                        st.error("Failed to delete from sheet.")
                                                else:
                                                    st.error(
                                                        "Could not find exact entry to delete. Might be already deleted/modified.")
                                            else:
                                                st.error("Could not re-fetch sheet data for deletion.")
                                        else:
                                            st.error("Incorrect password.")
                            else:
                                st.error("Could not retrieve details for selected entry.")
                    else:
                        st.info(f"You have no uploads in {selected_del_period_display} to delete.")
                elif df_all_del.empty:
                    st.info(f"No data in MCM sheet for {selected_del_period_display}.")
                else:
                    st.warning("Spreadsheet missing 'Audit Group Number' column.")
            elif not sheets_service:
                st.error("Google Sheets service not available.")
    st.markdown("</div>", unsafe_allow_html=True)


# --- Main App Logic ---
if not st.session_state.logged_in:
    login_page()
else:
    if 'drive_service' not in st.session_state or 'sheets_service' not in st.session_state or \
            st.session_state.drive_service is None or st.session_state.sheets_service is None:
        with st.spinner("Initializing Google Services... Ensure 'credentials.json' is present."):
            st.session_state.drive_service, st.session_state.sheets_service = get_google_services()
            if st.session_state.drive_service and st.session_state.sheets_service:
                st.success("Google Services Initialized.")
                st.session_state.drive_structure_initialized = False  # Trigger re-init of Drive structure
                st.rerun()
            # else: Error messages handled by get_google_services()

    if st.session_state.drive_service and st.session_state.sheets_service:
        if not st.session_state.get('drive_structure_initialized'):
            with st.spinner(
                    f"Initializing application folder structure on Google Drive ('{MASTER_DRIVE_FOLDER_NAME}')..."):
                if initialize_drive_structure(st.session_state.drive_service):
                    st.session_state.drive_structure_initialized = True
                    st.rerun()  # Rerun to ensure dashboards load with correct IDs
                else:  # Critical failure to init drive structure
                    st.error(
                        f"Failed to initialize Google Drive structure for '{MASTER_DRIVE_FOLDER_NAME}'. Application cannot proceed safely.")
                    if st.button("Logout", key="fail_logout_drive_init"):
                        st.session_state.logged_in = False;
                        st.rerun()
                    st.stop()  # Halt execution if drive structure is not ready

        # Proceed only if drive structure is initialized
        if st.session_state.get('drive_structure_initialized'):
            if st.session_state.role == "PCO":
                pco_dashboard(st.session_state.drive_service, st.session_state.sheets_service)
            elif st.session_state.role == "AuditGroup":
                audit_group_dashboard(st.session_state.drive_service, st.session_state.sheets_service)
    elif st.session_state.logged_in:  # Services not available but logged in
        st.warning("Google services are not available. Ensure 'credentials.json' is correct and network is stable.")
        if st.button("Logout", key="main_logout_gerror_sa_alt"):
            st.session_state.logged_in = False;
            st.rerun()

# # --- Final Check for API Key ---
# if YOUR_GEMINI_API_KEY == "AIzaSyBr37Or_irHH89GXzv0JpHOCULF_vMQDUw" or YOUR_GEMINI_API_KEY == "YOUR_API_KEY_HERE":
#     # Check if there's a sidebar, if not, display as a general warning
#     try:
#         st.sidebar.error("CRITICAL: Update 'YOUR_GEMINI_API_KEY' in app.py with your actual key!")
#     except st.errors.StreamlitAPIException:  # No sidebar (e.g., during login page)
#         st.error("CRITICAL: Gemini API Key needs to be updated in the application code (app.py).")
# # app.py
# import streamlit as st
# import pandas as pd
# from io import BytesIO
# import os
# import json
# import datetime
# from PIL import Image  # For a potential logo
# import time  # For retry delay and balloon visibility
#
# # --- Google API Imports ---
# from google.oauth2 import service_account  # For Service Account
# from googleapiclient.discovery import build
# from googleapiclient.errors import HttpError
# from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload
#
# # --- Gemini AI Import ---
# import google.generativeai as genai  # Ensure this is at the top level
#
# # --- Custom Module Imports ---
# # Assuming dar_processor.py only contains preprocess_pdf_text after simplification
# from dar_processor import preprocess_pdf_text
# from models import FlattenedAuditData, DARHeaderSchema, AuditParaSchema, ParsedDARReport
#
# # --- Streamlit Option Menu for better navigation ---
# from streamlit_option_menu import option_menu
#
# # --- Configuration ---
# # !!! REPLACE WITH YOUR ACTUAL GEMINI API KEY !!!
# YOUR_GEMINI_API_KEY = "AIzaSyBr37Or_irHH89GXzv0JpHOCULF_vMQDUw"
#
# # Google API Scopes
# SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
# # This CREDENTIALS_FILE should now be your Service Account JSON Key file
# CREDENTIALS_FILE = 'credentials.json'
# MCM_PERIODS_FILE = 'mcm_periods.json'  # To store Drive/Sheet IDs
#
# # --- User Credentials (Basic - for demonstration) ---
# USER_CREDENTIALS = {
#     "planning_officer": "pco_password",
#     **{f"audit_group{i}": f"ag{i}_audit" for i in range(1, 31)}
# }
# USER_ROLES = {
#     "planning_officer": "PCO",
#     **{f"audit_group{i}": "AuditGroup" for i in range(1, 31)}
# }
# AUDIT_GROUP_NUMBERS = {
#     f"audit_group{i}": i for i in range(1, 31)
# }
#
#
# # --- Custom CSS Styling ---
# def load_custom_css():
#     st.markdown("""
#     <style>
#         /* --- Global Styles --- */
#         body {
#             font-family: 'Roboto', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; /* Modern font */
#             background-color: #eef2f7; /* Lighter, softer background */
#             color: #4A4A4A; /* Darker gray for body text for better contrast */
#             line-height: 1.6;
#         }
#         .stApp {
#              background: linear-gradient(135deg, #f0f7ff 0%, #cfe7fa 100%); /* Softer blue gradient */
#         }
#
#         /* --- Titles and Headers --- */
#         .page-main-title { /* For titles outside login box */
#             font-size: 3em; /* Even bigger */
#             color: #1A237E; /* Darker, richer blue */
#             text-align: center;
#             padding: 30px 0 10px 0;
#             font-weight: 700;
#             letter-spacing: 1.5px;
#             text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
#         }
#         .page-app-subtitle { /* For subtitle outside login box */
#             font-size: 1.3em;
#             color: #3F51B5; /* Indigo */
#             text-align: center;
#             margin-top: -5px;
#             margin-bottom: 30px;
#             font-weight: 400;
#         }
#         .app-description {
#             font-size: 1.0em; /* Slightly larger description */
#             color: #455A64; /* Bluish gray */
#             text-align: center;
#             margin-bottom: 25px;
#             padding: 0 20px;
#             max-width: 700px; /* Limit width for readability */
#             margin-left: auto;
#             margin-right: auto;
#         }
#         .sub-header {
#             font-size: 1.6em;
#             color: #2779bd;
#             border-bottom: 3px solid #5dade2;
#             padding-bottom: 12px;
#             margin-top: 35px;
#             margin-bottom: 25px;
#             font-weight: 600;
#         }
#         .card h3 {
#             margin-top: 0;
#             color: #1abc9c;
#             font-size: 1.3em;
#             font-weight: 600;
#         }
#          .card h4 {
#             color: #2980b9;
#             font-size: 1.1em;
#             margin-top: 15px;
#             margin-bottom: 8px;
#         }
#
#
#         /* --- Cards --- */
#         .card {
#             background-color: #ffffff;
#             padding: 30px;
#             border-radius: 12px;
#             box-shadow: 0 6px 12px rgba(0,0,0,0.08);
#             margin-bottom: 25px;
#             border-left: 6px solid #5dade2;
#         }
#
#         /* --- Streamlit Widgets Styling --- */
#         .stButton>button {
#             border-radius: 25px;
#             background-image: linear-gradient(to right, #1abc9c 0%, #16a085 100%);
#             color: white;
#             padding: 12px 24px;
#             font-weight: bold;
#             border: none;
#             transition: all 0.3s ease;
#             box-shadow: 0 2px 4px rgba(0,0,0,0.1);
#         }
#         .stButton>button:hover {
#             background-image: linear-gradient(to right, #16a085 0%, #1abc9c 100%);
#             transform: translateY(-2px);
#             box-shadow: 0 4px 8px rgba(0,0,0,0.15);
#         }
#         .stButton>button[kind="secondary"] {
#             background-image: linear-gradient(to right, #e74c3c 0%, #c0392b 100%);
#         }
#         .stButton>button[kind="secondary"]:hover {
#             background-image: linear-gradient(to right, #c0392b 0%, #e74c3c 100%);
#         }
#         .stButton>button:disabled {
#             background-image: none;
#             background-color: #bdc3c7;
#             color: #7f8c8d;
#             box-shadow: none;
#             transform: none;
#         }
#         .stTextInput>div>div>input, .stSelectbox>div>div>div, .stDateInput>div>div>input, .stNumberInput>div>div>input {
#             border-radius: 8px;
#             border: 1px solid #ced4da;
#             padding: 10px;
#         }
#         .stTextInput>div>div>input:focus, .stSelectbox>div>div>div:focus-within, .stNumberInput>div>div>input:focus {
#             border-color: #5dade2;
#             box-shadow: 0 0 0 0.2rem rgba(93, 173, 226, 0.25);
#         }
#         .stFileUploader>div>div>button {
#             border-radius: 25px;
#             background-image: linear-gradient(to right, #5dade2 0%, #2980b9 100%);
#             color: white;
#             padding: 10px 18px;
#         }
#         .stFileUploader>div>div>button:hover {
#             background-image: linear-gradient(to right, #2980b9 0%, #5dade2 100%);
#         }
#
#
#         /* --- Login Page Specific --- */
#         .login-form-container { /* This will be the white box */
#             max-width: 500px;
#             margin: 20px auto; /* Adjusted margin */
#             padding: 30px;
#             background-color: #ffffff;
#             border-radius: 15px;
#             box-shadow: 0 10px 25px rgba(0,0,0,0.1);
#         }
#         .login-form-container .stButton>button {
#             background-image: linear-gradient(to right, #34495e 0%, #2c3e50 100%);
#         }
#         .login-form-container .stButton>button:hover {
#             background-image: linear-gradient(to right, #2c3e50 0%, #34495e 100%);
#         }
#         .login-header-text { /* For "User Login" text inside the box */
#             text-align: center;
#             color: #1a5276;
#             font-weight: 600;
#             font-size: 1.8em;
#             margin-bottom: 25px;
#         }
#         .login-logo {
#             display: block;
#             margin-left: auto;
#             margin-right: auto;
#             max-width: 70px;
#             margin-bottom: 15px;
#             border-radius: 50%;
#             box-shadow: 0 2px 4px rgba(0,0,0,0.1);
#         }
#
#
#         /* --- Sidebar Styling --- */
#         .css-1d391kg {
#             background-color: #ffffff;
#             padding: 15px !important;
#         }
#         .sidebar .stButton>button {
#              background-image: linear-gradient(to right, #e74c3c 0%, #c0392b 100%);
#         }
#         .sidebar .stButton>button:hover {
#              background-image: linear-gradient(to right, #c0392b 0%, #e74c3c 100%);
#         }
#         .sidebar .stMarkdown > div > p > strong {
#             color: #2c3e50;
#         }
#
#         /* --- Option Menu Customization --- */
#         div[data-testid="stOptionMenu"] > ul {
#             background-color: #ffffff;
#             border-radius: 25px;
#             padding: 8px;
#             box-shadow: 0 2px 5px rgba(0,0,0,0.05);
#         }
#         div[data-testid="stOptionMenu"] > ul > li > button {
#             border-radius: 20px;
#             margin: 0 5px !important;
#             border: none !important;
#             transition: all 0.3s ease;
#         }
#         div[data-testid="stOptionMenu"] > ul > li > button.selected {
#             background-image: linear-gradient(to right, #1abc9c 0%, #16a085 100%);
#             color: white;
#             font-weight: bold;
#             box-shadow: 0 2px 4px rgba(0,0,0,0.1);
#         }
#         div[data-testid="stOptionMenu"] > ul > li > button:hover:not(.selected) {
#             background-color: #e0e0e0;
#             color: #333;
#         }
#
#         /* --- Links --- */
#         a {
#             color: #3498db;
#             text-decoration: none;
#             font-weight: 500;
#         }
#         a:hover {
#             text-decoration: underline;
#             color: #2980b9;
#         }
#
#         /* --- Info/Warning/Error Boxes --- */
#         .stAlert {
#             border-radius: 8px;
#             padding: 15px;
#             border-left-width: 5px;
#         }
#         .stAlert[data-baseweb="notification"][role="alert"] > div:nth-child(2) {
#              font-size: 1.0em;
#         }
#         .stAlert[data-testid="stNotification"] {
#             box-shadow: 0 2px 10px rgba(0,0,0,0.07);
#         }
#         .stAlert[data-baseweb="notification"][kind="info"] { border-left-color: #3498db; }
#         .stAlert[data-baseweb="notification"][kind="success"] { border-left-color: #2ecc71; }
#         .stAlert[data-baseweb="notification"][kind="warning"] { border-left-color: #f39c12; }
#         .stAlert[data-baseweb="notification"][kind="error"] { border-left-color: #e74c3c; }
#
#     </style>
#     """, unsafe_allow_html=True)
#
#
# # --- Google API Authentication and Service Initialization (Modified for Service Account) ---
# def get_google_services():
#     """Authenticates using a service account and returns Drive and Sheets services."""
#     creds = None
#     if not os.path.exists(CREDENTIALS_FILE):
#         st.error(f"Service account credentials file ('{CREDENTIALS_FILE}') not found. "
#                  "Please place your service account JSON key file in the app directory.")
#         return None, None
#     try:
#         creds = service_account.Credentials.from_service_account_file(
#             CREDENTIALS_FILE, scopes=SCOPES)
#     except Exception as e:
#         st.error(f"Failed to load service account credentials: {e}")
#         return None, None
#
#     if not creds:
#         st.error("Could not initialize credentials from service account file.")
#         return None, None
#
#     try:
#         drive_service = build('drive', 'v3', credentials=creds)
#         sheets_service = build('sheets', 'v4', credentials=creds)
#         return drive_service, sheets_service
#     except HttpError as error:
#         st.error(f"An error occurred initializing Google services with service account: {error}")
#         return None, None
#     except Exception as e:
#         st.error(f"An unexpected error occurred with Google services (service account): {e}")
#         return None, None
#
#
# # --- MCM Period Management ---
# def load_mcm_periods():
#     if os.path.exists(MCM_PERIODS_FILE):
#         with open(MCM_PERIODS_FILE, 'r') as f:
#             try:
#                 return json.load(f)
#             except json.JSONDecodeError:
#                 return {}
#     return {}
#
#
# def save_mcm_periods(periods):
#     with open(MCM_PERIODS_FILE, 'w') as f:
#         json.dump(periods, f, indent=4)
#
#
# # --- Google Drive Functions ---
# def set_public_read_permission(drive_service, file_id):
#     """Sets 'anyone with the link can view' permission for a Drive file/folder."""
#     try:
#         permission = {'type': 'anyone', 'role': 'reader'}
#         drive_service.permissions().create(fileId=file_id, body=permission).execute()
#         print(f"Permission set to public read for file ID: {file_id}")
#     except HttpError as error:
#         st.warning(
#             f"Could not set public read permission for file ID {file_id}: {error}. Link might require login or manual sharing.")
#     except Exception as e:
#         st.warning(f"Unexpected error setting public permission for file ID {file_id}: {e}")
#
#
# def create_drive_folder(drive_service, folder_name):
#     try:
#         file_metadata = {
#             'name': folder_name,
#             'mimeType': 'application/vnd.google-apps.folder'
#         }
#         folder = drive_service.files().create(body=file_metadata, fields='id, webViewLink').execute()
#         folder_id = folder.get('id')
#         if folder_id:
#             set_public_read_permission(drive_service, folder_id)
#         return folder_id, folder.get('webViewLink')
#     except HttpError as error:
#         st.error(f"An error occurred creating Drive folder: {error}")
#         return None, None
#
#
# def upload_to_drive(drive_service, file_content_or_path, folder_id, filename_on_drive):
#     try:
#         file_metadata = {'name': filename_on_drive, 'parents': [folder_id]}
#         media_body = None
#
#         if isinstance(file_content_or_path, str) and os.path.exists(file_content_or_path):
#             media_body = MediaFileUpload(file_content_or_path, mimetype='application/pdf', resumable=True)
#         elif isinstance(file_content_or_path, bytes):
#             fh = BytesIO(file_content_or_path)
#             media_body = MediaIoBaseUpload(fh, mimetype='application/pdf', resumable=True)
#         elif isinstance(file_content_or_path, BytesIO):
#             file_content_or_path.seek(0)
#             media_body = MediaIoBaseUpload(file_content_or_path, mimetype='application/pdf', resumable=True)
#         else:
#             st.error(f"Unsupported file content type for Google Drive upload: {type(file_content_or_path)}")
#             return None, None
#
#         if media_body is None:
#             st.error("Media body for upload could not be prepared.")
#             return None, None
#
#         request = drive_service.files().create(
#             body=file_metadata,
#             media_body=media_body,
#             fields='id, webViewLink'
#         )
#         file = request.execute()
#         file_id = file.get('id')
#         if file_id:
#             set_public_read_permission(drive_service, file_id)
#         return file_id, file.get('webViewLink')
#     except HttpError as error:
#         st.error(f"An API error occurred uploading to Drive: {error}")
#         return None, None
#     except Exception as e:
#         st.error(f"An unexpected error in upload_to_drive: {e}")
#         return None, None
#
#
# # --- Google Sheets Functions ---
# def create_spreadsheet(sheets_service, drive_service, title):
#     try:
#         spreadsheet_body = {
#             'properties': {'title': title}
#         }
#         spreadsheet = sheets_service.spreadsheets().create(body=spreadsheet_body,
#                                                            fields='spreadsheetId,spreadsheetUrl').execute()
#         spreadsheet_id = spreadsheet.get('spreadsheetId')
#         if spreadsheet_id and drive_service:
#             set_public_read_permission(drive_service, spreadsheet_id)
#         return spreadsheet_id, spreadsheet.get('spreadsheetUrl')
#     except HttpError as error:
#         st.error(f"An error occurred creating Spreadsheet: {error}")
#         return None, None
#     except Exception as e:
#         st.error(f"An unexpected error occurred creating Spreadsheet: {e}")
#         return None, None
#
#
# def append_to_spreadsheet(sheets_service, spreadsheet_id, values_to_append):
#     try:
#         body = {'values': values_to_append}
#         sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
#         sheets = sheet_metadata.get('sheets', '')
#         first_sheet_title = sheets[0].get("properties", {}).get("title", "Sheet1")
#
#         range_to_check = f"{first_sheet_title}!A1:L1"
#         result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,
#                                                             range=range_to_check).execute()
#         header_row_values = result.get('values', [])
#
#         if not header_row_values:
#             header_values_list = [
#                 "Audit Group Number", "GSTIN", "Trade Name", "Category",
#                 "Total Amount Detected (Overall Rs)", "Total Amount Recovered (Overall Rs)",
#                 "Audit Para Number", "Audit Para Heading",
#                 "Revenue Involved (Lakhs Rs)", "Revenue Recovered (Lakhs Rs)",
#                 "DAR PDF URL", "Record Created Date"
#             ]
#             sheets_service.spreadsheets().values().append(
#                 spreadsheetId=spreadsheet_id,
#                 range=f"{first_sheet_title}!A1",
#                 valueInputOption='USER_ENTERED',
#                 body={'values': header_values_list}
#             ).execute()
#
#         append_result = sheets_service.spreadsheets().values().append(
#             spreadsheetId=spreadsheet_id,
#             range=f"{first_sheet_title}!A1",
#             valueInputOption='USER_ENTERED',
#             body=body
#         ).execute()
#         return append_result
#     except HttpError as error:
#         st.error(f"An error occurred appending to Spreadsheet: {error}")
#         return None
#
#
# def read_from_spreadsheet(sheets_service, spreadsheet_id, sheet_name="Sheet1"):
#     try:
#         result = sheets_service.spreadsheets().values().get(
#             spreadsheetId=spreadsheet_id,
#             range=sheet_name
#         ).execute()
#         values = result.get('values', [])
#         if not values:
#             return pd.DataFrame()
#         else:
#             expected_cols = [
#                 "Audit Group Number", "GSTIN", "Trade Name", "Category",
#                 "Total Amount Detected (Overall Rs)", "Total Amount Recovered (Overall Rs)",
#                 "Audit Para Number", "Audit Para Heading",
#                 "Revenue Involved (Lakhs Rs)", "Revenue Recovered (Lakhs Rs)",
#                 "DAR PDF URL", "Record Created Date"
#             ]
#             if values and values[0] == expected_cols:
#                 return pd.DataFrame(values[1:], columns=values[0])
#             else:
#                 st.warning(
#                     f"Spreadsheet '{sheet_name}' headers do not match expected or are missing. Attempting to load data.")
#                 if values:
#                     num_cols = len(values[0]) if values else 0
#                     df_cols = values[0] if values and len(values[0]) == num_cols else [f"Col_{j + 1}" for j in
#                                                                                        range(num_cols)]
#                     data_start_row = 1 if values and values[0] == expected_cols else 0
#                     return pd.DataFrame(values[data_start_row:], columns=df_cols if data_start_row == 1 else None)
#                 else:
#                     return pd.DataFrame()
#     except HttpError as error:
#         st.error(f"An error occurred reading from Spreadsheet: {error}")
#         return pd.DataFrame()
#
#
# def delete_spreadsheet_rows(sheets_service, spreadsheet_id, sheet_id_gid, row_indices_to_delete):
#     """Deletes specified rows from a sheet. Row indices are 0-based sheet data indices (after header)."""
#     if not row_indices_to_delete:
#         return True
#
#     requests = []
#     # Sort indices in descending order to avoid shifting issues during batch deletion
#     # The row_indices_to_delete should be 0-based relative to the data (excluding header)
#     # So, if deleting the first data row, it's index 0, which is sheet row 2 (1-indexed)
#     for data_row_index in sorted(row_indices_to_delete, reverse=True):
#         sheet_row_start_index = data_row_index + 1  # +1 because Sheets API is 1-indexed for rows, and we skip header
#         requests.append({
#             "deleteDimension": {
#                 "range": {
#                     "sheetId": sheet_id_gid,
#                     "dimension": "ROWS",
#                     "startIndex": sheet_row_start_index,
#                     "endIndex": sheet_row_start_index + 1
#                 }
#             }
#         })
#
#     if requests:
#         try:
#             body = {'requests': requests}
#             sheets_service.spreadsheets().batchUpdate(
#                 spreadsheetId=spreadsheet_id, body=body).execute()
#             return True
#         except HttpError as error:
#             st.error(f"An error occurred deleting rows from Spreadsheet: {error}")
#             return False
#     return True
#
#
# # --- Gemini Data Extraction with Retry ---
# def get_structured_data_with_gemini(api_key: str, text_content: str, max_retries=2) -> ParsedDARReport:
#     """
#     Calls Gemini API with the full PDF text and parses the response.
#     Includes retry logic for JSONDecodeError.
#     """
#     if not api_key or api_key == "YOUR_API_KEY_HERE":
#         return ParsedDARReport(parsing_errors="Gemini API Key not configured in app.py.")
#     if text_content.startswith("Error processing PDF with pdfplumber:"):
#         return ParsedDARReport(parsing_errors=text_content)
#
#     genai.configure(api_key=api_key)
#     model = genai.GenerativeModel('gemini-1.5-flash-latest')
#
#     prompt = f"""
#     You are an expert GST audit report analyst. Based on the following FULL text from a Departmental Audit Report (DAR),
#     where all text from all pages, including tables, is provided, extract the specified information
#     and structure it as a JSON object. Focus on identifying narrative sections for audit para details,
#     even if they are intermingled with tabular data. Notes like "[INFO: ...]" in the text are for context only.
#
#     The JSON object should follow this structure precisely:
#     {{
#       "header": {{
#         "audit_group_number": "integer or null (e.g., if 'Group-VI' or 'Gr 6', extract 6; must be between 1 and 30)",
#         "gstin": "string or null",
#         "trade_name": "string or null",
#         "category": "string ('Large', 'Medium', 'Small') or null",
#         "total_amount_detected_overall_rs": "float or null (numeric value in Rupees)",
#         "total_amount_recovered_overall_rs": "float or null (numeric value in Rupees)"
#       }},
#       "audit_paras": [
#         {{
#           "audit_para_number": "integer or null (primary number from para heading, e.g., for 'Para-1...' use 1; must be between 1 and 50)",
#           "audit_para_heading": "string or null (the descriptive title of the para)",
#           "revenue_involved_lakhs_rs": "float or null (numeric value in Lakhs of Rupees, e.g., Rs. 50,000 becomes 0.5)",
#           "revenue_recovered_lakhs_rs": "float or null (numeric value in Lakhs of Rupees)"
#         }}
#       ],
#       "parsing_errors": "string or null (any notes about parsing issues, or if extraction is incomplete)"
#     }}
#
#     Key Instructions:
#     1.  Header Information: Extract `audit_group_number` (as integer 1-30, e.g., 'Group-VI' becomes 6), `gstin`, `trade_name`, `category`, `total_amount_detected_overall_rs`, `total_amount_recovered_overall_rs`.
#     2.  Audit Paras: Identify each distinct para. Extract `audit_para_number` (as integer 1-50), `audit_para_heading`, `revenue_involved_lakhs_rs` (converted to Lakhs), `revenue_recovered_lakhs_rs` (converted to Lakhs).
#     3.  Use null for missing values. Monetary values as float.
#     4.  If no audit paras found, `audit_paras` should be an empty list [].
#
#     DAR Text Content:
#     --- START OF DAR TEXT ---
#     {text_content}
#     --- END OF DAR TEXT ---
#
#     Provide ONLY the JSON object as your response. Do not include any explanatory text before or after the JSON.
#     """
#
#     attempt = 0
#     last_exception = None
#     while attempt <= max_retries:
#         attempt += 1
#         print(f"\n--- Calling Gemini (Attempt {attempt}/{max_retries + 1}) ---")
#         try:
#             response = model.generate_content(prompt)
#
#             cleaned_response_text = response.text.strip()
#             if cleaned_response_text.startswith("```json"):
#                 cleaned_response_text = cleaned_response_text[7:]
#             elif cleaned_response_text.startswith("`json"):
#                 cleaned_response_text = cleaned_response_text[6:]
#             if cleaned_response_text.endswith("```"):
#                 cleaned_response_text = cleaned_response_text[:-3]
#
#             if not cleaned_response_text:
#                 error_message = f"Gemini returned an empty response on attempt {attempt}."
#                 print(error_message)
#                 last_exception = ValueError(error_message)
#                 if attempt > max_retries:
#                     return ParsedDARReport(parsing_errors=error_message)
#                 time.sleep(1 + attempt)
#                 continue
#
#             json_data = json.loads(cleaned_response_text)
#             if "header" not in json_data or "audit_paras" not in json_data:
#                 error_message = f"Gemini response (Attempt {attempt}) missing 'header' or 'audit_paras' key. Response: {cleaned_response_text[:500]}"
#                 print(error_message)
#                 last_exception = ValueError(error_message)
#                 if attempt > max_retries:
#                     return ParsedDARReport(parsing_errors=error_message)
#                 time.sleep(1 + attempt)
#                 continue
#
#             parsed_report = ParsedDARReport(**json_data)
#             print(f"Gemini call (Attempt {attempt}) successful. Paras found: {len(parsed_report.audit_paras)}")
#             if parsed_report.audit_paras:
#                 for idx, para_obj in enumerate(parsed_report.audit_paras):
#                     if not para_obj.audit_para_heading:
#                         print(
#                             f"  Note: Para {idx + 1} (Number: {para_obj.audit_para_number}) has a missing heading from Gemini.")
#             return parsed_report
#         except json.JSONDecodeError as e:
#             raw_response_text = "No response text available"
#             if 'response' in locals() and hasattr(response, 'text'):
#                 raw_response_text = response.text
#             error_message = f"Gemini output (Attempt {attempt}) was not valid JSON: {e}. Response: '{raw_response_text[:1000]}...'"
#             print(error_message)
#             last_exception = e
#             if attempt > max_retries:
#                 return ParsedDARReport(parsing_errors=error_message)
#             time.sleep(attempt * 2)
#         except Exception as e:
#             raw_response_text = "No response text available"
#             if 'response' in locals() and hasattr(response, 'text'):
#                 raw_response_text = response.text
#             error_message = f"Error (Attempt {attempt}) during Gemini/Pydantic: {type(e).__name__} - {e}. Response: {raw_response_text[:500]}"
#             print(error_message)
#             last_exception = e
#             if attempt > max_retries:
#                 return ParsedDARReport(parsing_errors=error_message)
#             time.sleep(attempt * 2)
#
#     return ParsedDARReport(
#         parsing_errors=f"Gemini call failed after {max_retries + 1} attempts. Last error: {last_exception}")
#
#
# # --- Streamlit App UI and Logic ---
# st.set_page_config(layout="wide", page_title="e-MCM App - GST Audit 1")
# load_custom_css()
#
# # --- Login ---
# if 'logged_in' not in st.session_state:
#     st.session_state.logged_in = False
#     st.session_state.username = ""
#     st.session_state.role = ""
#     st.session_state.audit_group_no = None
#     st.session_state.ag_current_extracted_data = []
#     st.session_state.ag_pdf_drive_url = None
#     st.session_state.ag_validation_errors = []
#     st.session_state.ag_editor_data = pd.DataFrame()
#     st.session_state.ag_current_mcm_key = None
#     st.session_state.ag_current_uploaded_file_name = None
#
#
# def login_page():
#     # The main titles are now outside the login container
#     st.markdown("<div class='page-main-title'>e-MCM App</div>", unsafe_allow_html=True)
#     st.markdown("<h2 class='page-app-subtitle'>GST Audit 1 Commissionerate</h2>", unsafe_allow_html=True)
#
#     # The white box for login form elements
#     #st.markdown("<div class='login-form-container'>", unsafe_allow_html=True)
#     # --- Logo Display using Base64 ---
#
#     import base64
#     import os
#
#
#
#     def get_image_base64_str(img_path):
#         """Converts an image file to a base64 string for embedding in HTML."""
#         try:
#             with open(img_path, "rb") as img_file:
#                 return base64.b64encode(img_file.read()).decode('utf-8')
#         except FileNotFoundError:
#             # This error will now be more visible in the Streamlit app itself
#             st.error(
#                 f"Logo image not found at path: {img_path}. Please ensure 'img.png' is in the same directory as your script.")
#             return None
#         except Exception as e:
#             st.error(f"Error reading image file {img_path}: {e}")
#             return None
#
#     image_path = "logo.png"  # Assumes img.png is in the same directory as your script.
#     # If it's elsewhere, provide the correct relative path e.g., "assets/img.png"
#
#     base64_image = get_image_base64_str(image_path)
#
#     if base64_image:
#         # Determine image type (e.g., png, jpg) for the data URI
#         image_type = os.path.splitext(image_path)[1].lower().replace(".", "")
#         if not image_type:  # Default to png if extension is missing or unknown
#             image_type = "png"
#
#         st.markdown(
#             f"<div class='login-header'><img src='data:image/{image_type};base64,{base64_image}' alt='Logo' class='login-logo'></div>",
#             unsafe_allow_html=True
#         )
#     else:
#         # Fallback text if image can't be loaded
#         st.markdown(
#             "<div class='login-header' style='color: red; font-weight: bold;'>[Logo Image Not Found or Error Loading]</div>",
#             unsafe_allow_html=True
#         )
#         # The get_image_base64_str function will also show an st.error message.
#     # --- End of Logo Display ---
#     # ---- MODIFIED LINE BELOW ----
#
#
#
#     st.markdown("<h2 class='login-header-text'>User Login</h2>", unsafe_allow_html=True)
#     st.markdown("""
#     <div class='app-description'>
#         Welcome! This platform streamlines Draft Audit Report (DAR) collection and processing.
#         PCOs manage MCM periods; Audit Groups upload DARs for AI-powered data extraction.
#     </div>
#     """, unsafe_allow_html=True)
#
#     username = st.text_input("Username", key="login_username_styled", placeholder="Enter your username")
#     password = st.text_input("Password", type="password", key="login_password_styled",
#                              placeholder="Enter your password")
#
#     if st.button("Login", key="login_button_styled", use_container_width=True):
#         if username in USER_CREDENTIALS and USER_CREDENTIALS[username] == password:
#             st.session_state.logged_in = True
#             st.session_state.username = username
#             st.session_state.role = USER_ROLES[username]
#             if st.session_state.role == "AuditGroup":
#                 st.session_state.audit_group_no = AUDIT_GROUP_NUMBERS[username]
#             st.success(f"Logged in as {username} ({st.session_state.role})")
#             st.rerun()
#         else:
#             st.error("Invalid username or password")
#     st.markdown("</div>", unsafe_allow_html=True)  # Close login-form-container
#
#
# # --- PCO Dashboard ---
# def pco_dashboard(drive_service, sheets_service):
#     st.markdown("<div class='sub-header'>Planning & Coordination Officer Dashboard</div>", unsafe_allow_html=True)
#     mcm_periods = load_mcm_periods()
#
#     with st.sidebar:
#         st.image(
#             "https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/Indian_Ministry_of_Finance_logo.svg/1200px-Indian_Ministry_of_Finance_logo.svg.png",
#             width=80)
#         st.markdown(f"**User:** {st.session_state.username}")
#         st.markdown(f"**Role:** {st.session_state.role}")
#         if st.button("Logout", key="pco_logout_styled", use_container_width=True):
#             st.session_state.logged_in = False
#             st.session_state.username = ""
#             st.session_state.role = ""
#             st.rerun()
#         st.markdown("---")
#
#     selected_tab = option_menu(
#         menu_title=None,
#         options=["Create MCM Period", "Manage MCM Periods", "View Uploaded Reports", "Visualizations"],
#         icons=["calendar-plus-fill", "sliders", "eye-fill", "bar-chart-fill"],
#         menu_icon="gear-wide-connected",
#         default_index=0,
#         orientation="horizontal",
#         styles={
#             "container": {"padding": "5px !important", "background-color": "#e9ecef"},
#             "icon": {"color": "#007bff", "font-size": "20px"},
#             "nav-link": {"font-size": "16px", "text-align": "center", "margin": "0px", "--hover-color": "#d1e7fd"},
#             "nav-link-selected": {"background-color": "#007bff", "color": "white"},
#         }
#     )
#
#     st.markdown("<div class='card'>", unsafe_allow_html=True)
#
#     if selected_tab == "Create MCM Period":
#         st.markdown("<h3>Create New MCM Period</h3>", unsafe_allow_html=True)
#         current_year = datetime.datetime.now().year
#         years = list(range(current_year - 1, current_year + 3))
#         months = ["January", "February", "March", "April", "May", "June",
#                   "July", "August", "September", "October", "November", "December"]
#
#         col1, col2 = st.columns(2)
#         with col1:
#             selected_year = st.selectbox("Select Year", options=years, index=years.index(current_year), key="pco_year")
#         with col2:
#             selected_month_name = st.selectbox("Select Month", options=months, index=datetime.datetime.now().month - 1,
#                                                key="pco_month")
#
#         selected_month_num = months.index(selected_month_name) + 1
#         period_key = f"{selected_year}-{selected_month_num:02d}"
#
#         if period_key in mcm_periods:
#             st.warning(f"MCM Period for {selected_month_name} {selected_year} already exists.")
#             st.markdown(
#                 f"**Drive Folder:** <a href='{mcm_periods[period_key]['drive_folder_url']}' target='_blank'>Open Folder</a>",
#                 unsafe_allow_html=True)
#             st.markdown(
#                 f"**Spreadsheet:** <a href='{mcm_periods[period_key]['spreadsheet_url']}' target='_blank'>Open Sheet</a>",
#                 unsafe_allow_html=True)
#         else:
#             if st.button(f"Create MCM for {selected_month_name} {selected_year}", key="pco_create_mcm",
#                          use_container_width=True):
#                 if not drive_service or not sheets_service:
#                     st.error("Google Services not available. Cannot create MCM period.")
#                 else:
#                     with st.spinner("Creating Google Drive folder and Spreadsheet..."):
#                         folder_name = f"MCM_DARs_{selected_month_name}_{selected_year}"
#                         spreadsheet_title = f"MCM_Audit_Paras_{selected_month_name}_{selected_year}"
#
#                         folder_id, folder_url = create_drive_folder(drive_service, folder_name)
#                         sheet_id, sheet_url = create_spreadsheet(sheets_service, drive_service, spreadsheet_title)
#
#                         if folder_id and sheet_id:
#                             mcm_periods[period_key] = {
#                                 "year": selected_year, "month_num": selected_month_num,
#                                 "month_name": selected_month_name,
#                                 "drive_folder_id": folder_id, "drive_folder_url": folder_url,
#                                 "spreadsheet_id": sheet_id, "spreadsheet_url": sheet_url,
#                                 "active": True
#                             }
#                             save_mcm_periods(mcm_periods)
#                             st.success(f"Successfully created MCM period for {selected_month_name} {selected_year}!")
#                             st.markdown(f"**Drive Folder:** <a href='{folder_url}' target='_blank'>Open Folder</a>",
#                                         unsafe_allow_html=True)
#                             st.markdown(f"**Spreadsheet:** <a href='{sheet_url}' target='_blank'>Open Sheet</a>",
#                                         unsafe_allow_html=True)
#                             st.balloons()
#                             time.sleep(0.5)
#                             st.rerun()
#                         else:
#                             st.error("Failed to create Drive folder or Spreadsheet.")
#
#     elif selected_tab == "Manage MCM Periods":
#         st.markdown("<h3>Manage Existing MCM Periods</h3>", unsafe_allow_html=True)
#         if not mcm_periods:
#             st.info("No MCM periods created yet.")
#         else:
#             sorted_periods_keys = sorted(mcm_periods.keys(), reverse=True)
#
#             for period_key in sorted_periods_keys:
#                 data = mcm_periods[period_key]
#                 st.markdown(f"<h4>{data['month_name']} {data['year']}</h4>", unsafe_allow_html=True)
#                 col1, col2, col3, col4 = st.columns([2, 2, 1, 2])
#                 with col1:
#                     st.markdown(f"<a href='{data['drive_folder_url']}' target='_blank'>Open Drive Folder</a>",
#                                 unsafe_allow_html=True)
#                 with col2:
#                     st.markdown(f"<a href='{data['spreadsheet_url']}' target='_blank'>Open Spreadsheet</a>",
#                                 unsafe_allow_html=True)
#                 with col3:
#                     is_active = data.get("active", False)
#                     new_status = st.checkbox("Active", value=is_active, key=f"active_{period_key}_styled_manage")
#                     if new_status != is_active:
#                         mcm_periods[period_key]["active"] = new_status
#                         save_mcm_periods(mcm_periods)
#                         st.success(f"Status for {data['month_name']} {data['year']} updated.")
#                         st.rerun()
#                 with col4:
#                     if st.button("Delete Period Record", key=f"delete_mcm_{period_key}", type="secondary"):
#                         st.session_state.period_to_delete = period_key
#                         st.session_state.show_delete_confirm = True
#                         st.rerun()
#                 st.markdown("---")
#
#             if st.session_state.get('show_delete_confirm', False) and st.session_state.get('period_to_delete'):
#                 period_key_to_delete = st.session_state.period_to_delete
#                 period_data_to_delete = mcm_periods.get(period_key_to_delete, {})
#
#                 with st.form(key=f"delete_confirm_form_{period_key_to_delete}"):
#                     st.warning(
#                         f"Are you sure you want to delete the MCM period record for **{period_data_to_delete.get('month_name')} {period_data_to_delete.get('year')}** from this application?")
#                     st.caption(
#                         "This action only removes the period from the app's tracking. It **does NOT delete** the actual Google Drive folder or the Google Spreadsheet associated with it. Those must be managed manually in Google Drive/Sheets if needed.")
#                     pco_password_confirm = st.text_input("Enter your PCO password to confirm deletion:",
#                                                          type="password",
#                                                          key=f"pco_pass_confirm_{period_key_to_delete}")
#
#                     col_form1, col_form2 = st.columns(2)
#                     with col_form1:
#                         submitted_delete = st.form_submit_button("Yes, Delete This Period Record",
#                                                                  use_container_width=True)
#                     with col_form2:
#                         if st.form_submit_button("Cancel", type="secondary", use_container_width=True):
#                             st.session_state.show_delete_confirm = False
#                             st.session_state.period_to_delete = None
#                             st.rerun()
#
#                     if submitted_delete:
#                         if pco_password_confirm == USER_CREDENTIALS.get("planning_officer"):
#                             del mcm_periods[period_key_to_delete]
#                             save_mcm_periods(mcm_periods)
#                             st.success(
#                                 f"MCM period record for {period_data_to_delete.get('month_name')} {period_data_to_delete.get('year')} has been deleted from the app.")
#                             st.session_state.show_delete_confirm = False
#                             st.session_state.period_to_delete = None
#                             st.rerun()
#                         else:
#                             st.error("Incorrect password. Deletion cancelled.")
#
#     elif selected_tab == "View Uploaded Reports":
#         st.markdown("<h3>View Uploaded Reports Summary</h3>", unsafe_allow_html=True)
#         active_periods = {k: v for k, v in mcm_periods.items()}
#         if not active_periods:
#             st.info("No MCM periods created yet to view reports.")
#         else:
#             period_options = [f"{p_data['month_name']} {p_data['year']}" for p_key, p_data in
#                               sorted(active_periods.items(), key=lambda item: item[0], reverse=True)]
#             selected_period_display = st.selectbox("Select MCM Period to View", options=period_options,
#                                                    key="pco_view_period")
#
#             if selected_period_display:
#                 selected_period_key = None
#                 for p_key, p_data in active_periods.items():
#                     if f"{p_data['month_name']} {p_data['year']}" == selected_period_display:
#                         selected_period_key = p_key
#                         break
#
#                 if selected_period_key and sheets_service:
#                     sheet_id = mcm_periods[selected_period_key]['spreadsheet_id']
#                     st.markdown(f"**Fetching data for {selected_period_display}...**")
#                     with st.spinner("Loading data from Google Sheet..."):
#                         df = read_from_spreadsheet(sheets_service, sheet_id)
#
#                     if not df.empty:
#                         st.markdown("<h4>Summary of Uploads:</h4>", unsafe_allow_html=True)
#                         if 'Audit Group Number' in df.columns:
#                             try:
#                                 df['Audit Group Number'] = pd.to_numeric(df['Audit Group Number'], errors='coerce')
#                                 df.dropna(subset=['Audit Group Number'], inplace=True)
#
#                                 dars_per_group = df.groupby('Audit Group Number')['DAR PDF URL'].nunique().reset_index(
#                                     name='DARs Uploaded')
#                                 st.write("**DARs Uploaded per Audit Group:**")
#                                 st.dataframe(dars_per_group, use_container_width=True)
#
#                                 paras_per_group = df.groupby('Audit Group Number').size().reset_index(
#                                     name='Total Para Entries')
#                                 st.write("**Total Para Entries per Audit Group:**")
#                                 st.dataframe(paras_per_group, use_container_width=True)
#
#                                 st.markdown("<h4>Detailed Data:</h4>", unsafe_allow_html=True)
#                                 st.dataframe(df, use_container_width=True)
#
#                             except Exception as e:
#                                 st.error(f"Error processing data for summary: {e}")
#                                 st.write("Raw Data:")
#                                 st.dataframe(df, use_container_width=True)
#                         else:
#                             st.warning("Spreadsheet does not contain 'Audit Group Number' column for summary.")
#                             st.dataframe(df, use_container_width=True)
#                     else:
#                         st.info(f"No data found in the spreadsheet for {selected_period_display}.")
#                 elif not sheets_service:
#                     st.error("Google Sheets service not available.")
#
#     elif selected_tab == "Visualizations":
#         st.markdown("<h3>Data Visualizations</h3>", unsafe_allow_html=True)
#         all_mcm_periods = {k: v for k, v in mcm_periods.items()}
#         if not all_mcm_periods:
#             st.info("No MCM periods created yet to visualize data.")
#         else:
#             viz_period_options = [f"{p_data['month_name']} {p_data['year']}" for p_key, p_data in
#                                   sorted(all_mcm_periods.items(), key=lambda item: item[0], reverse=True)]
#             selected_viz_period_display = st.selectbox("Select MCM Period for Visualization",
#                                                        options=viz_period_options, key="pco_viz_period")
#
#             if selected_viz_period_display and sheets_service:
#                 selected_viz_period_key = None
#                 for p_key, p_data in all_mcm_periods.items():
#                     if f"{p_data['month_name']} {p_data['year']}" == selected_viz_period_display:
#                         selected_viz_period_key = p_key
#                         break
#
#                 if selected_viz_period_key:
#                     sheet_id_viz = all_mcm_periods[selected_viz_period_key]['spreadsheet_id']
#                     st.markdown(f"**Fetching data for {selected_viz_period_display} visualizations...**")
#                     with st.spinner("Loading data..."):
#                         df_viz = read_from_spreadsheet(sheets_service, sheet_id_viz)
#
#                     if not df_viz.empty:
#                         # Data Cleaning for numeric columns
#                         amount_cols = ['Total Amount Detected (Overall Rs)', 'Total Amount Recovered (Overall Rs)',
#                                        'Revenue Involved (Lakhs Rs)', 'Revenue Recovered (Lakhs Rs)']
#                         for col in amount_cols:
#                             if col in df_viz.columns:
#                                 df_viz[col] = pd.to_numeric(df_viz[col], errors='coerce').fillna(0)
#                         if 'Audit Group Number' in df_viz.columns:
#                             df_viz['Audit Group Number'] = pd.to_numeric(df_viz['Audit Group Number'],
#                                                                          errors='coerce').fillna(0).astype(int)
#
#                         st.markdown("---")
#                         st.markdown("<h4>Group-wise Performance</h4>", unsafe_allow_html=True)
#
#                         # Top 5 Groups by Total Detection
#                         if 'Total Amount Detected (Overall Rs)' in df_viz.columns and 'Audit Group Number' in df_viz.columns:
#                             # Summing up all 'Total Amount Detected (Overall Rs)' for each group.
#                             # This assumes that if a DAR has multiple paras, the overall total is repeated.
#                             # If overall total is meant to be unique per DAR, this logic might over-sum if not careful with data structure.
#                             # A safer way if overall total is unique per DAR:
#                             # df_unique_dars = df_viz.drop_duplicates(subset=['Audit Group Number', 'DAR PDF URL'])
#                             # top_detection_groups_data = df_unique_dars.groupby('Audit Group Number')['Total Amount Detected (Overall Rs)'].sum().nlargest(5)
#
#                             # Current approach: sum all entries, which might be what's intended if each row has the overall for that specific DAR
#                             top_detection_groups_data = df_viz.groupby('Audit Group Number')[
#                                 'Total Amount Detected (Overall Rs)'].sum().nlargest(5)
#                             if not top_detection_groups_data.empty:
#                                 st.write("**Top 5 Groups by Total Detection Amount (Rs):**")
#                                 st.bar_chart(top_detection_groups_data)
#                             else:
#                                 st.info("Not enough data for 'Top Detection Groups' chart.")
#
#                         # Top 5 Groups by Total Realisation
#                         if 'Total Amount Recovered (Overall Rs)' in df_viz.columns and 'Audit Group Number' in df_viz.columns:
#                             top_recovery_groups_data = df_viz.groupby('Audit Group Number')[
#                                 'Total Amount Recovered (Overall Rs)'].sum().nlargest(5)
#                             if not top_recovery_groups_data.empty:
#                                 st.write("**Top 5 Groups by Total Realisation Amount (Rs):**")
#                                 st.bar_chart(top_recovery_groups_data)
#                             else:
#                                 st.info("Not enough data for 'Top Realisation Groups' chart.")
#
#                         # Top 5 Groups by Recovery/Detection Ratio
#                         if 'Total Amount Detected (Overall Rs)' in df_viz.columns and \
#                                 'Total Amount Recovered (Overall Rs)' in df_viz.columns and \
#                                 'Audit Group Number' in df_viz.columns:
#                             group_summary = df_viz.groupby('Audit Group Number').agg(
#                                 Total_Detected=('Total Amount Detected (Overall Rs)', 'sum'),
#                                 Total_Recovered=('Total Amount Recovered (Overall Rs)', 'sum')
#                             ).reset_index()
#                             group_summary['Recovery_Ratio'] = group_summary.apply(
#                                 lambda row: (row['Total_Recovered'] / row['Total_Detected']) * 100 if pd.notna(
#                                     row['Total_Detected']) and row['Total_Detected'] > 0 and pd.notna(
#                                     row['Total_Recovered']) else 0, axis=1
#                             )
#                             top_ratio_groups_data = group_summary.nlargest(5, 'Recovery_Ratio')[
#                                 ['Audit Group Number', 'Recovery_Ratio']].set_index('Audit Group Number')
#                             if not top_ratio_groups_data.empty:
#                                 st.write("**Top 5 Groups by Recovery/Detection Ratio (%):**")
#                                 st.bar_chart(top_ratio_groups_data['Recovery_Ratio'])
#                             else:
#                                 st.info("Not enough data for 'Top Recovery Ratio Groups' chart.")
#
#                         st.markdown("---")
#                         st.markdown("<h4>Para-wise Performance</h4>", unsafe_allow_html=True)
#                         num_paras_to_show = st.number_input("Select N for Top N Paras:", min_value=1, max_value=20,
#                                                             value=5, step=1, key="top_n_paras_viz")
#
#                         # Top N Detection Paras
#                         if 'Revenue Involved (Lakhs Rs)' in df_viz.columns:
#                             df_paras_only = df_viz[df_viz['Audit Para Number'].notna() & (df_viz[
#                                                                                               'Audit Para Heading'] != "N/A - Header Info Only (Add Paras Manually)") & (
#                                                                df_viz[
#                                                                    'Audit Para Heading'] != "Manual Entry Required") & (
#                                                                df_viz[
#                                                                    'Audit Para Heading'] != "Manual Entry - PDF Error") & (
#                                                                df_viz[
#                                                                    'Audit Para Heading'] != "Manual Entry - PDF Upload Failed")]
#                             top_detection_paras = df_paras_only.nlargest(num_paras_to_show,
#                                                                          'Revenue Involved (Lakhs Rs)')
#                             if not top_detection_paras.empty:
#                                 st.write(f"**Top {num_paras_to_show} Detection Paras (by Revenue Involved):**")
#                                 st.dataframe(top_detection_paras[
#                                                  ['Audit Group Number', 'Trade Name', 'Audit Para Number',
#                                                   'Audit Para Heading', 'Revenue Involved (Lakhs Rs)']],
#                                              use_container_width=True)
#                             else:
#                                 st.info("Not enough data for 'Top Detection Paras' list.")
#
#                         # Top N Realisation Paras
#                         if 'Revenue Recovered (Lakhs Rs)' in df_viz.columns:
#                             df_paras_only = df_viz[df_viz['Audit Para Number'].notna() & (df_viz[
#                                                                                               'Audit Para Heading'] != "N/A - Header Info Only (Add Paras Manually)") & (
#                                                                df_viz[
#                                                                    'Audit Para Heading'] != "Manual Entry Required") & (
#                                                                df_viz[
#                                                                    'Audit Para Heading'] != "Manual Entry - PDF Error") & (
#                                                                df_viz[
#                                                                    'Audit Para Heading'] != "Manual Entry - PDF Upload Failed")]
#                             top_recovery_paras = df_paras_only.nlargest(num_paras_to_show,
#                                                                         'Revenue Recovered (Lakhs Rs)')
#                             if not top_recovery_paras.empty:
#                                 st.write(f"**Top {num_paras_to_show} Realisation Paras (by Revenue Recovered):**")
#                                 st.dataframe(top_recovery_paras[
#                                                  ['Audit Group Number', 'Trade Name', 'Audit Para Number',
#                                                   'Audit Para Heading', 'Revenue Recovered (Lakhs Rs)']],
#                                              use_container_width=True)
#                             else:
#                                 st.info("Not enough data for 'Top Realisation Paras' list.")
#                     else:
#                         st.info(
#                             f"No data found in the spreadsheet for {selected_viz_period_display} to generate visualizations.")
#                 elif not sheets_service:
#                     st.error("Google Sheets service not available for visualization.")
#
#     st.markdown("</div>", unsafe_allow_html=True)
#
#
# # --- Validation Function (reusable) ---
# MANDATORY_FIELDS_FOR_SHEET = {
#     "audit_group_number": "Audit Group Number",
#     "gstin": "GSTIN",
#     "trade_name": "Trade Name",
#     "category": "Category",
#     "total_amount_detected_overall_rs": "Total Amount Detected (Overall Rs)",
#     "total_amount_recovered_overall_rs": "Total Amount Recovered (Overall Rs)",
#     "audit_para_number": "Audit Para Number",
#     "audit_para_heading": "Audit Para Heading",
#     "revenue_involved_lakhs_rs": "Revenue Involved (Lakhs Rs)",
#     "revenue_recovered_lakhs_rs": "Revenue Recovered (Lakhs Rs)"
# }
# VALID_CATEGORIES = ["Large", "Medium", "Small"]
#
#
# def validate_data_for_sheet(data_df_to_validate):
#     validation_errors = []
#     if data_df_to_validate.empty:
#         validation_errors.append("No data to validate. Please extract or enter data first.")
#         return validation_errors
#
#     # Per-row validations
#     for index, row in data_df_to_validate.iterrows():
#         row_display_id = f"Row {index + 1} (Para: {row.get('audit_para_number', 'N/A')})"
#
#         # Mandatory field checks
#         for field_key, field_name in MANDATORY_FIELDS_FOR_SHEET.items():
#             value = row.get(field_key)
#             is_missing = False
#             if value is None:
#                 is_missing = True
#             elif isinstance(value, str) and not value.strip():
#                 is_missing = True
#             elif pd.isna(value):
#                 is_missing = True
#
#             if is_missing:
#                 if field_key in ["audit_para_number", "audit_para_heading", "revenue_involved_lakhs_rs",
#                                  "revenue_recovered_lakhs_rs"]:
#                     if row.get('audit_para_heading', "").startswith("N/A - Header Info Only") and pd.isna(
#                             row.get('audit_para_number')):
#                         continue
#                 validation_errors.append(f"{row_display_id}: '{field_name}' is missing or empty.")
#
#         # Category specific validation
#         category_val = row.get('category')
#         if pd.notna(category_val) and category_val.strip() and category_val not in VALID_CATEGORIES:
#             validation_errors.append(
#                 f"{row_display_id}: 'Category' ('{category_val}') is invalid. Must be one of {VALID_CATEGORIES} (case-sensitive)."
#             )
#         elif pd.isna(category_val) or (isinstance(category_val, str) and not category_val.strip()):
#             if "category" in MANDATORY_FIELDS_FOR_SHEET:  # Check if category was in mandatory list
#                 validation_errors.append(f"{row_display_id}: 'Category' is missing. Please select a valid category.")
#
#     # Category Consistency Check per Trade Name
#     if 'trade_name' in data_df_to_validate.columns and 'category' in data_df_to_validate.columns:
#         trade_name_categories = {}
#         for index, row in data_df_to_validate.iterrows():
#             trade_name = row.get('trade_name')
#             category = row.get('category')
#             if pd.notna(trade_name) and trade_name.strip() and \
#                     pd.notna(category) and category.strip() and category in VALID_CATEGORIES:
#                 if trade_name not in trade_name_categories:
#                     trade_name_categories[trade_name] = set()
#                 trade_name_categories[trade_name].add(category)
#
#         for tn, cats in trade_name_categories.items():
#             if len(cats) > 1:
#                 validation_errors.append(
#                     f"Consistency Error: Trade Name '{tn}' has multiple valid categories assigned: {', '.join(sorted(list(cats)))}. Please ensure a single, consistent category for this trade name across all its para entries."
#                 )
#     return sorted(list(set(validation_errors)))
#
#
# # --- Audit Group Dashboard ---
# def audit_group_dashboard(drive_service, sheets_service):
#     st.markdown(f"<div class='sub-header'>Audit Group {st.session_state.audit_group_no} Dashboard</div>",
#                 unsafe_allow_html=True)
#     mcm_periods = load_mcm_periods()
#
#     active_periods = {k: v for k, v in mcm_periods.items() if v.get("active")}
#
#     with st.sidebar:
#         st.image(
#             "https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/Indian_Ministry_of_Finance_logo.svg/1200px-Indian_Ministry_of_Finance_logo.svg.png",
#             width=80)
#         st.markdown(f"**User:** {st.session_state.username}")
#         st.markdown(f"**Group No:** {st.session_state.audit_group_no}")
#         if st.button("Logout", key="ag_logout_styled", use_container_width=True):
#             for key in ['ag_current_extracted_data', 'ag_pdf_drive_url', 'ag_validation_errors', 'ag_editor_data',
#                         'ag_current_mcm_key', 'ag_current_uploaded_file_name', 'ag_row_to_delete_details',
#                         'ag_show_delete_confirm']:
#                 if key in st.session_state:
#                     del st.session_state[key]
#             st.session_state.logged_in = False;
#             st.session_state.username = "";
#             st.session_state.role = "";
#             st.session_state.audit_group_no = None
#             st.rerun()
#         st.markdown("---")
#
#     selected_tab = option_menu(
#         menu_title=None,
#         options=["Upload DAR for MCM", "View My Uploaded DARs", "Delete My DAR Entries"],  # Changed Tab Name
#         icons=["cloud-upload-fill", "eye-fill", "trash2-fill"],  # Updated Icon
#         menu_icon="person-workspace",
#         default_index=0,
#         orientation="horizontal",
#         styles={
#             "container": {"padding": "5px !important", "background-color": "#e9ecef"},
#             "icon": {"color": "#28a745", "font-size": "20px"},
#             "nav-link": {"font-size": "16px", "text-align": "center", "margin": "0px", "--hover-color": "#d4edda"},
#             "nav-link-selected": {"background-color": "#28a745", "color": "white"},
#         }
#     )
#     st.markdown("<div class='card'>", unsafe_allow_html=True)
#
#     if selected_tab == "Upload DAR for MCM":
#         st.markdown("<h3>Upload DAR PDF for MCM Period</h3>", unsafe_allow_html=True)
#
#         if not active_periods:
#             st.warning("No active MCM periods available for upload. Please contact the Planning Officer.")
#         else:
#             period_options = {f"{p_data['month_name']} {p_data['year']}": p_key for p_key, p_data in
#                               sorted(active_periods.items(), reverse=True)}
#             selected_period_display = st.selectbox("Select Active MCM Period", options=list(period_options.keys()),
#                                                    key="ag_select_mcm_period_upload")
#
#             if selected_period_display:
#                 selected_mcm_key = period_options[selected_period_display]
#                 mcm_info = mcm_periods[selected_mcm_key]
#
#                 if st.session_state.get('ag_current_mcm_key') != selected_mcm_key:
#                     st.session_state.ag_current_extracted_data = []
#                     st.session_state.ag_pdf_drive_url = None
#                     st.session_state.ag_validation_errors = []
#                     st.session_state.ag_editor_data = pd.DataFrame()
#                     st.session_state.ag_current_mcm_key = selected_mcm_key
#                     st.session_state.ag_current_uploaded_file_name = None
#
#                 st.info(f"Preparing to upload for: {mcm_info['month_name']} {mcm_info['year']}")
#                 uploaded_dar_file = st.file_uploader("Choose a DAR PDF file", type="pdf",
#                                                      key=f"dar_upload_ag_{selected_mcm_key}_{st.session_state.get('uploader_key_suffix', 0)}")
#
#                 if uploaded_dar_file is not None:
#                     if st.session_state.get('ag_current_uploaded_file_name') != uploaded_dar_file.name:
#                         st.session_state.ag_current_extracted_data = []
#                         st.session_state.ag_pdf_drive_url = None
#                         st.session_state.ag_validation_errors = []
#                         st.session_state.ag_editor_data = pd.DataFrame()
#                         st.session_state.ag_current_uploaded_file_name = uploaded_dar_file.name
#
#                     if st.button("Extract Data from PDF", key=f"extract_btn_ag_{selected_mcm_key}",
#                                  use_container_width=True):
#                         st.session_state.ag_validation_errors = []
#
#                         with st.spinner("Processing PDF and extracting data with AI..."):
#                             dar_pdf_bytes = uploaded_dar_file.getvalue()
#                             dar_filename_on_drive = f"AG{st.session_state.audit_group_no}_{uploaded_dar_file.name}"
#
#                             st.session_state.ag_pdf_drive_url = None
#                             pdf_drive_id, pdf_drive_url_temp = upload_to_drive(drive_service, dar_pdf_bytes,
#                                                                                mcm_info['drive_folder_id'],
#                                                                                dar_filename_on_drive)
#
#                             if not pdf_drive_id:
#                                 st.error("Failed to upload PDF to Google Drive. Cannot proceed.")
#                                 st.session_state.ag_editor_data = pd.DataFrame([{
#                                     "audit_group_number": st.session_state.audit_group_no, "gstin": None,
#                                     "trade_name": None, "category": None,
#                                     "total_amount_detected_overall_rs": None, "total_amount_recovered_overall_rs": None,
#                                     "audit_para_number": None, "audit_para_heading": "Manual Entry - PDF Upload Failed",
#                                     "revenue_involved_lakhs_rs": None, "revenue_recovered_lakhs_rs": None
#                                 }])
#                             else:
#                                 st.session_state.ag_pdf_drive_url = pdf_drive_url_temp
#                                 st.success(
#                                     f"DAR PDF uploaded to Google Drive: [Link]({st.session_state.ag_pdf_drive_url})")
#
#                                 preprocessed_text = preprocess_pdf_text(BytesIO(dar_pdf_bytes))
#                                 if preprocessed_text.startswith("Error processing PDF with pdfplumber:") or \
#                                         preprocessed_text.startswith("Error in preprocess_pdf_text_"):
#                                     st.error(f"PDF Preprocessing Error: {preprocessed_text}")
#                                     st.warning("AI extraction cannot proceed. Please fill data manually.")
#                                     default_row = {
#                                         "audit_group_number": st.session_state.audit_group_no, "gstin": None,
#                                         "trade_name": None, "category": None,
#                                         "total_amount_detected_overall_rs": None,
#                                         "total_amount_recovered_overall_rs": None,
#                                         "audit_para_number": None, "audit_para_heading": "Manual Entry - PDF Error",
#                                         "revenue_involved_lakhs_rs": None, "revenue_recovered_lakhs_rs": None
#                                     }
#                                     st.session_state.ag_editor_data = pd.DataFrame([default_row])
#                                 else:
#                                     parsed_report_obj = get_structured_data_with_gemini(YOUR_GEMINI_API_KEY,
#                                                                                         preprocessed_text)
#
#                                     temp_extracted_list = []
#                                     ai_extraction_failed_completely = True
#
#                                     if parsed_report_obj.parsing_errors:
#                                         st.warning(f"AI Parsing Issues: {parsed_report_obj.parsing_errors}")
#
#                                     if parsed_report_obj and parsed_report_obj.header:
#                                         header = parsed_report_obj.header
#                                         ai_extraction_failed_completely = False
#                                         if parsed_report_obj.audit_paras:
#                                             for para_idx, para in enumerate(parsed_report_obj.audit_paras):
#                                                 flat_item_data = {
#                                                     "audit_group_number": st.session_state.audit_group_no,
#                                                     "gstin": header.gstin, "trade_name": header.trade_name,
#                                                     "category": header.category,
#                                                     "total_amount_detected_overall_rs": header.total_amount_detected_overall_rs,
#                                                     "total_amount_recovered_overall_rs": header.total_amount_recovered_overall_rs,
#                                                     "audit_para_number": para.audit_para_number,
#                                                     "audit_para_heading": para.audit_para_heading,
#                                                     "revenue_involved_lakhs_rs": para.revenue_involved_lakhs_rs,
#                                                     "revenue_recovered_lakhs_rs": para.revenue_recovered_lakhs_rs
#                                                 }
#                                                 temp_extracted_list.append(flat_item_data)
#                                         elif header.trade_name:
#                                             temp_extracted_list.append({
#                                                 "audit_group_number": st.session_state.audit_group_no,
#                                                 "gstin": header.gstin, "trade_name": header.trade_name,
#                                                 "category": header.category,
#                                                 "total_amount_detected_overall_rs": header.total_amount_detected_overall_rs,
#                                                 "total_amount_recovered_overall_rs": header.total_amount_recovered_overall_rs,
#                                                 "audit_para_number": None,
#                                                 "audit_para_heading": "N/A - Header Info Only (Add Paras Manually)",
#                                                 "revenue_involved_lakhs_rs": None, "revenue_recovered_lakhs_rs": None
#                                             })
#                                         else:
#                                             st.error("AI failed to extract key header information (like Trade Name).")
#                                             ai_extraction_failed_completely = True
#
#                                     if ai_extraction_failed_completely or not temp_extracted_list:
#                                         st.warning(
#                                             "AI extraction failed or yielded no data. Please fill details manually in the table below.")
#                                         default_row = {
#                                             "audit_group_number": st.session_state.audit_group_no, "gstin": None,
#                                             "trade_name": None, "category": None,
#                                             "total_amount_detected_overall_rs": None,
#                                             "total_amount_recovered_overall_rs": None,
#                                             "audit_para_number": None, "audit_para_heading": "Manual Entry Required",
#                                             "revenue_involved_lakhs_rs": None, "revenue_recovered_lakhs_rs": None
#                                         }
#                                         st.session_state.ag_editor_data = pd.DataFrame([default_row])
#                                     else:
#                                         st.session_state.ag_editor_data = pd.DataFrame(temp_extracted_list)
#                                         st.info(
#                                             "Data extracted. Please review and edit in the table below before submitting.")
#
#                 if not isinstance(st.session_state.get('ag_editor_data'), pd.DataFrame):
#                     st.session_state.ag_editor_data = pd.DataFrame()
#
#                 if uploaded_dar_file is not None and st.session_state.ag_editor_data.empty and \
#                         st.session_state.get('ag_pdf_drive_url'):
#                     st.warning(
#                         "AI could not extract data, or no data was previously loaded. A template row is provided for manual entry.")
#                     default_row = {
#                         "audit_group_number": st.session_state.audit_group_no, "gstin": None, "trade_name": None,
#                         "category": None,
#                         "total_amount_detected_overall_rs": None, "total_amount_recovered_overall_rs": None,
#                         "audit_para_number": None, "audit_para_heading": "Manual Entry",
#                         "revenue_involved_lakhs_rs": None, "revenue_recovered_lakhs_rs": None
#                     }
#                     st.session_state.ag_editor_data = pd.DataFrame([default_row])
#
#                 if not st.session_state.ag_editor_data.empty:
#                     st.markdown("<h4>Review and Edit Extracted Data:</h4>", unsafe_allow_html=True)
#                     st.caption("Note: 'Audit Group Number' is pre-filled from your login. Add new rows for more paras.")
#
#                     ag_column_order = [
#                         "audit_group_number", "gstin", "trade_name", "category",
#                         "total_amount_detected_overall_rs", "total_amount_recovered_overall_rs",
#                         "audit_para_number", "audit_para_heading",
#                         "revenue_involved_lakhs_rs", "revenue_recovered_lakhs_rs"
#                     ]
#
#                     df_to_edit_ag = st.session_state.ag_editor_data.copy()
#                     df_to_edit_ag["audit_group_number"] = st.session_state.audit_group_no
#
#                     for col in ag_column_order:
#                         if col not in df_to_edit_ag.columns:
#                             df_to_edit_ag[col] = None
#
#                     ag_column_config = {
#                         "audit_group_number": st.column_config.NumberColumn("Audit Group", format="%d", width="small",
#                                                                             disabled=True),
#                         "gstin": st.column_config.TextColumn("GSTIN", width="medium", help="Enter 15-digit GSTIN"),
#                         "trade_name": st.column_config.TextColumn("Trade Name", width="large"),
#                         "category": st.column_config.SelectboxColumn("Category", options=VALID_CATEGORIES,
#                                                                      width="medium",
#                                                                      help=f"Select from {VALID_CATEGORIES}"),
#                         "total_amount_detected_overall_rs": st.column_config.NumberColumn("Total Detected (Rs)",
#                                                                                           format="%.2f",
#                                                                                           width="medium"),
#                         "total_amount_recovered_overall_rs": st.column_config.NumberColumn("Total Recovered (Rs)",
#                                                                                            format="%.2f",
#                                                                                            width="medium"),
#                         "audit_para_number": st.column_config.NumberColumn("Para No.", format="%d", width="small",
#                                                                            help="Enter para number as integer"),
#                         "audit_para_heading": st.column_config.TextColumn("Para Heading", width="xlarge"),
#                         "revenue_involved_lakhs_rs": st.column_config.NumberColumn("Rev. Involved (Lakhs)",
#                                                                                    format="%.2f", width="medium"),
#                         "revenue_recovered_lakhs_rs": st.column_config.NumberColumn("Rev. Recovered (Lakhs)",
#                                                                                     format="%.2f", width="medium"),
#                     }
#
#                     editor_instance_key = f"ag_data_editor_{selected_mcm_key}_{st.session_state.ag_current_uploaded_file_name or 'no_file'}"
#
#                     edited_ag_df_output = st.data_editor(
#                         df_to_edit_ag.reindex(columns=ag_column_order, fill_value=None),
#                         column_config=ag_column_config,
#                         num_rows="dynamic",
#                         key=editor_instance_key,
#                         use_container_width=True,
#                         height=400
#                     )
#
#                     if st.button("Validate and Submit to MCM Sheet", key=f"submit_ag_btn_{selected_mcm_key}",
#                                  use_container_width=True):
#                         current_data_for_validation = pd.DataFrame(edited_ag_df_output)
#                         current_data_for_validation["audit_group_number"] = st.session_state.audit_group_no
#
#                         validation_errors = validate_data_for_sheet(current_data_for_validation)
#                         st.session_state.ag_validation_errors = validation_errors
#
#                         if not validation_errors:
#                             if not st.session_state.ag_pdf_drive_url:
#                                 st.error(
#                                     "PDF Drive URL is missing. This usually means the PDF was not uploaded to Drive. Please click 'Extract Data from PDF' again.")
#                             else:
#                                 with st.spinner("Submitting data to Google Sheet..."):
#                                     rows_to_append_final = []
#                                     record_created_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#                                     for index, row_to_save in current_data_for_validation.iterrows():
#                                         sheet_row = [
#                                             row_to_save.get("audit_group_number"),
#                                             row_to_save.get("gstin"), row_to_save.get("trade_name"),
#                                             row_to_save.get("category"),
#                                             row_to_save.get("total_amount_detected_overall_rs"),
#                                             row_to_save.get("total_amount_recovered_overall_rs"),
#                                             row_to_save.get("audit_para_number"), row_to_save.get("audit_para_heading"),
#                                             row_to_save.get("revenue_involved_lakhs_rs"),
#                                             row_to_save.get("revenue_recovered_lakhs_rs"),
#                                             st.session_state.ag_pdf_drive_url,
#                                             record_created_date
#                                         ]
#                                         rows_to_append_final.append(sheet_row)
#
#                                     if rows_to_append_final:
#                                         sheet_id = mcm_info['spreadsheet_id']
#                                         append_result = append_to_spreadsheet(sheets_service, sheet_id,
#                                                                               rows_to_append_final)
#                                         if append_result:
#                                             st.success(
#                                                 f"Data for '{st.session_state.ag_current_uploaded_file_name}' successfully submitted to MCM sheet for {mcm_info['month_name']} {mcm_info['year']}!")
#                                             st.balloons()
#                                             time.sleep(0.5)
#                                             st.session_state.ag_current_extracted_data = []
#                                             st.session_state.ag_pdf_drive_url = None
#                                             st.session_state.ag_editor_data = pd.DataFrame()
#                                             st.session_state.ag_current_uploaded_file_name = None
#                                             st.session_state.uploader_key_suffix = st.session_state.get(
#                                                 'uploader_key_suffix', 0) + 1
#                                             st.rerun()
#                                         else:
#                                             st.error("Failed to append data to Google Sheet after validation.")
#                                     else:
#                                         st.error("No data to submit after validation, though validation passed.")
#                         else:
#                             st.error(
#                                 "Validation Failed! Please correct the errors displayed below and try submitting again.")
#
#                 if st.session_state.get('ag_validation_errors'):
#                     st.markdown("---")
#                     st.subheader("⚠️ Validation Errors - Please Fix Before Submitting:")
#                     for err in st.session_state.ag_validation_errors:
#                         st.warning(err)
#
#
#     elif selected_tab == "View My Uploaded DARs":
#         st.markdown("<h3>My Uploaded DARs</h3>", unsafe_allow_html=True)
#         if not active_periods and not mcm_periods:
#             st.info("No MCM periods have been created by the Planning Officer yet.")
#         else:
#             all_period_options = {f"{p_data['month_name']} {p_data['year']}": p_key for p_key, p_data in
#                                   sorted(mcm_periods.items(), key=lambda item: item[0], reverse=True)}
#             if not all_period_options:
#                 st.info("No MCM periods found.")
#             else:
#                 selected_view_period_display = st.selectbox("Select MCM Period to View Your Uploads",
#                                                             options=list(all_period_options.keys()),
#                                                             key="ag_view_my_dars_period")
#
#                 if selected_view_period_display and sheets_service:
#                     selected_view_period_key = all_period_options[selected_view_period_display]
#                     view_mcm_info = mcm_periods[selected_view_period_key]
#                     sheet_id_to_view = view_mcm_info['spreadsheet_id']
#
#                     st.markdown(f"**Fetching your uploads for {selected_view_period_display}...**")
#                     with st.spinner("Loading data from Google Sheet..."):
#                         df_all_uploads = read_from_spreadsheet(sheets_service, sheet_id_to_view)
#
#                     if not df_all_uploads.empty:
#                         if 'Audit Group Number' in df_all_uploads.columns:
#                             df_all_uploads['Audit Group Number'] = df_all_uploads['Audit Group Number'].astype(str)
#                             my_group_uploads_df = df_all_uploads[
#                                 df_all_uploads['Audit Group Number'] == str(st.session_state.audit_group_no)]
#
#                             if not my_group_uploads_df.empty:
#                                 st.markdown(f"<h4>Your Uploads for {selected_view_period_display}:</h4>",
#                                             unsafe_allow_html=True)
#                                 df_display_my_uploads = my_group_uploads_df.copy()
#                                 if 'DAR PDF URL' in df_display_my_uploads.columns:
#                                     df_display_my_uploads['DAR PDF URL'] = df_display_my_uploads['DAR PDF URL'].apply(
#                                         lambda x: f'<a href="{x}" target="_blank">View PDF</a>' if pd.notna(
#                                             x) and isinstance(x, str) and x.startswith("http") else "No Link"
#                                     )
#                                 view_cols = ["Trade Name", "Category", "Audit Para Number", "Audit Para Heading",
#                                              "DAR PDF URL", "Record Created Date"]
#                                 for col_v in view_cols:
#                                     if col_v not in df_display_my_uploads.columns:
#                                         df_display_my_uploads[col_v] = "N/A"
#
#                                 st.markdown(df_display_my_uploads[view_cols].to_html(escape=False, index=False),
#                                             unsafe_allow_html=True)
#                             else:
#                                 st.info(f"You have not uploaded any DARs for {selected_view_period_display} yet.")
#                         else:
#                             st.warning(
#                                 "The MCM spreadsheet is missing the 'Audit Group Number' column. Cannot filter your uploads.")
#                     else:
#                         st.info(f"No data found in the MCM spreadsheet for {selected_view_period_display}.")
#                 elif not sheets_service:
#                     st.error("Google Sheets service not available.")
#
#     elif selected_tab == "Delete My DAR Entries":
#         st.markdown("<h3>Delete My Uploaded DAR Entries</h3>", unsafe_allow_html=True)
#         st.info(
#             "Select an MCM period to view your entries. You can then delete specific entries after password confirmation. This action removes the entry from the Google Sheet only; the PDF in Google Drive will NOT be deleted.")
#
#         if not mcm_periods:
#             st.info("No MCM periods have been created yet.")
#         else:
#             all_period_options_del = {f"{p_data['month_name']} {p_data['year']}": p_key for p_key, p_data in
#                                       sorted(mcm_periods.items(), key=lambda item: item[0], reverse=True)}
#             selected_del_period_display = st.selectbox("Select MCM Period to Manage Entries",
#                                                        options=list(all_period_options_del.keys()),
#                                                        key="ag_delete_my_dars_period_select")
#
#             if selected_del_period_display and sheets_service:
#                 selected_del_period_key = all_period_options_del[selected_del_period_display]
#                 del_mcm_info = mcm_periods[selected_del_period_key]
#                 sheet_id_to_manage = del_mcm_info['spreadsheet_id']
#
#                 try:
#                     sheet_metadata_for_gid = sheets_service.spreadsheets().get(
#                         spreadsheetId=sheet_id_to_manage).execute()
#                     first_sheet_gid = sheet_metadata_for_gid.get('sheets', [{}])[0].get('properties', {}).get('sheetId',
#                                                                                                               0)
#                 except Exception as e_gid:
#                     st.error(f"Could not fetch sheet GID for deletion: {e_gid}")
#                     first_sheet_gid = 0
#
#                 st.markdown(f"**Fetching your uploads for {selected_del_period_display}...**")
#                 with st.spinner("Loading data..."):
#                     df_all_uploads_del = read_from_spreadsheet(sheets_service, sheet_id_to_manage)
#
#                 if not df_all_uploads_del.empty and 'Audit Group Number' in df_all_uploads_del.columns:
#                     df_all_uploads_del['Audit Group Number'] = df_all_uploads_del['Audit Group Number'].astype(str)
#                     my_group_uploads_df_del = df_all_uploads_del[df_all_uploads_del['Audit Group Number'] == str(
#                         st.session_state.audit_group_no)].copy()  # Use .copy()
#                     my_group_uploads_df_del.reset_index(
#                         inplace=True)  # Keep original DataFrame index as 'original_df_index'
#                     my_group_uploads_df_del.rename(columns={'index': 'original_df_index'}, inplace=True)
#
#                     if not my_group_uploads_df_del.empty:
#                         st.markdown(
#                             f"<h4>Your Uploads in {selected_del_period_display} (Select an entry to delete):</h4>",
#                             unsafe_allow_html=True)
#
#                         options_for_deletion = ["--Select an entry--"]
#                         # Store a mapping of display option to actual DataFrame index and key data for identification
#                         st.session_state.ag_deletable_entries_map = {}
#
#                         for idx, row_data in my_group_uploads_df_del.iterrows():
#                             # Create a unique identifier string for the selectbox option
#                             identifier_str = (
#                                 f"Entry (TN: {str(row_data.get('Trade Name', 'N/A'))[:20]}..., "
#                                 f"Para: {row_data.get('Audit Para Number', 'N/A')}, "
#                                 f"Date: {row_data.get('Record Created Date', 'N/A')})"
#                             )
#                             options_for_deletion.append(identifier_str)
#                             # Store the necessary info to find this row later in the full sheet
#                             st.session_state.ag_deletable_entries_map[identifier_str] = {
#                                 "trade_name": str(row_data.get('Trade Name')),
#                                 "audit_para_number": str(row_data.get('Audit Para Number')),  # Compare as strings
#                                 "record_created_date": str(row_data.get('Record Created Date')),
#                                 "dar_pdf_url": str(row_data.get('DAR PDF URL'))  # For very precise matching
#                             }
#
#                         selected_entry_to_delete_display = st.selectbox(
#                             "Select Entry to Delete:",
#                             options_for_deletion,
#                             key=f"del_select_{selected_del_period_key}"
#                         )
#
#                         if selected_entry_to_delete_display != "--Select an entry--":
#                             row_to_delete_identifier_data = st.session_state.ag_deletable_entries_map.get(
#                                 selected_entry_to_delete_display)
#
#                             if row_to_delete_identifier_data:
#                                 st.warning(
#                                     f"You selected to delete entry related to: **{row_to_delete_identifier_data.get('trade_name')} - Para {row_to_delete_identifier_data.get('audit_para_number')}** (Uploaded: {row_to_delete_identifier_data.get('record_created_date')})")
#
#                                 with st.form(
#                                         key=f"delete_ag_entry_form_{selected_entry_to_delete_display.replace(' ', '_')}"):  # Unique form key
#                                     st.write("Please confirm your password to delete this entry from the Google Sheet.")
#                                     ag_password_confirm = st.text_input("Your Password:", type="password",
#                                                                         key=f"ag_pass_confirm_del_{selected_entry_to_delete_display.replace(' ', '_')}")
#                                     submitted_ag_delete = st.form_submit_button("Confirm Deletion of This Entry")
#
#                                     if submitted_ag_delete:
#                                         if ag_password_confirm == USER_CREDENTIALS.get(st.session_state.username):
#                                             # Re-fetch the entire sheet to get current row indices accurately
#                                             current_sheet_data_df = read_from_spreadsheet(sheets_service,
#                                                                                           sheet_id_to_manage)
#                                             if not current_sheet_data_df.empty:
#                                                 indices_to_delete_from_sheet = []  # 0-based indices for batchUpdate
#                                                 for sheet_idx, sheet_row in current_sheet_data_df.iterrows():
#                                                     # Robust matching based on multiple fields
#                                                     match = (
#                                                             str(sheet_row.get('Audit Group Number')) == str(
#                                                         st.session_state.audit_group_no) and
#                                                             str(sheet_row.get(
#                                                                 'Trade Name')) == row_to_delete_identifier_data.get(
#                                                         'trade_name') and
#                                                             str(sheet_row.get(
#                                                                 'Audit Para Number')) == row_to_delete_identifier_data.get(
#                                                         'audit_para_number') and
#                                                             str(sheet_row.get(
#                                                                 'Record Created Date')) == row_to_delete_identifier_data.get(
#                                                         'record_created_date') and
#                                                             str(sheet_row.get(
#                                                                 'DAR PDF URL')) == row_to_delete_identifier_data.get(
#                                                         'dar_pdf_url')
#                                                     )
#                                                     if match:
#                                                         indices_to_delete_from_sheet.append(
#                                                             sheet_idx)  # 0-based index of data row
#
#                                                 if indices_to_delete_from_sheet:
#                                                     if delete_spreadsheet_rows(sheets_service, sheet_id_to_manage,
#                                                                                first_sheet_gid,
#                                                                                indices_to_delete_from_sheet):
#                                                         st.success(
#                                                             f"Entry for '{row_to_delete_identifier_data.get('trade_name')}' - Para '{row_to_delete_identifier_data.get('audit_para_number')}' deleted successfully from the sheet.")
#                                                         time.sleep(0.5)
#                                                         st.rerun()
#                                                     else:
#                                                         st.error(
#                                                             "Failed to delete the entry from the sheet. See console/app logs for details.")
#                                                 else:
#                                                     st.error(
#                                                         "Could not find the exact entry in the sheet to delete. It might have been already modified or deleted by another process.")
#                                             else:
#                                                 st.error("Could not re-fetch sheet data for deletion verification.")
#                                         else:
#                                             st.error("Incorrect password. Deletion cancelled.")
#                             else:
#                                 st.error("Could not retrieve details for the selected entry. Please try again.")
#                     else:
#                         st.info(f"You have no uploads in {selected_del_period_display} to request deletion for.")
#                 elif df_all_uploads_del.empty:
#                     st.info(f"No data found in the MCM spreadsheet for {selected_del_period_display}.")
#                 else:
#                     st.warning(
#                         "Spreadsheet is missing 'Audit Group Number' column. Cannot identify your uploads for deletion.")
#             elif not sheets_service:
#                 st.error("Google Sheets service not available.")
#
#     st.markdown("</div>", unsafe_allow_html=True)
#
#
# # --- Main App Logic ---
# if not st.session_state.logged_in:
#     login_page()
# else:
#     if 'drive_service' not in st.session_state or 'sheets_service' not in st.session_state or \
#             st.session_state.drive_service is None or st.session_state.sheets_service is None:
#         with st.spinner(
#                 "Initializing Google Services... Please ensure your 'credentials.json' (Service Account Key) is in the app directory."):
#             st.session_state.drive_service, st.session_state.sheets_service = get_google_services()
#             if st.session_state.drive_service and st.session_state.sheets_service:
#                 st.success("Google Services Initialized.")
#                 st.rerun()
#             else:
#                 pass
#
#     if st.session_state.drive_service and st.session_state.sheets_service:
#         if st.session_state.role == "PCO":
#             pco_dashboard(st.session_state.drive_service, st.session_state.sheets_service)
#         elif st.session_state.role == "AuditGroup":
#             audit_group_dashboard(st.session_state.drive_service, st.session_state.sheets_service)
#     elif st.session_state.logged_in:
#         st.warning(
#             "Google services are not yet available or failed to initialize. Please ensure 'credentials.json' (Service Account Key) is correctly placed and configured.")
#         if st.button("Logout", key="main_logout_gerror_sa"):
#             st.session_state.logged_in = False;
#             st.rerun()
#
# # --- Final Check for API Key ---
# if YOUR_GEMINI_API_KEY == "YOUR_API_KEY_HERE":
#     st.sidebar.error("CRITICAL: Update 'YOUR_GEMINI_API_KEY' in app.py")
# # # app.py
# # import streamlit as st
# # import pandas as pd
# # from io import BytesIO
# # import os
# # import json
# # import datetime
# # from PIL import Image  # For a potential logo
# # import time  # For retry delay and balloon visibility
# #
# # # --- Google API Imports ---
# # from google.oauth2 import service_account  # For Service Account
# # from googleapiclient.discovery import build
# # from googleapiclient.errors import HttpError
# # from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload
# #
# # # --- Gemini AI Import ---
# # import google.generativeai as genai  # Ensure this is at the top level
# #
# # # --- Custom Module Imports ---
# # # Assuming dar_processor.py only contains preprocess_pdf_text after simplification
# # from dar_processor import preprocess_pdf_text
# # from models import FlattenedAuditData, DARHeaderSchema, AuditParaSchema, ParsedDARReport
# #
# # # --- Streamlit Option Menu for better navigation ---
# # from streamlit_option_menu import option_menu
# #
# # # --- Configuration ---
# # # !!! REPLACE WITH YOUR ACTUAL GEMINI API KEY !!!
# # YOUR_GEMINI_API_KEY = "AIzaSyBr37Or_irHH89GXzv0JpHOCULF_vMQDUw"
# #
# # # Google API Scopes
# # SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
# # # This CREDENTIALS_FILE should now be your Service Account JSON Key file
# # CREDENTIALS_FILE = 'credentials.json'
# # MCM_PERIODS_FILE = 'mcm_periods.json'  # To store Drive/Sheet IDs
# #
# # # --- User Credentials (Basic - for demonstration) ---
# # USER_CREDENTIALS = {
# #     "planning_officer": "pco_password",
# #     **{f"audit_group{i}": f"ag{i}_audit" for i in range(1, 31)}
# # }
# # USER_ROLES = {
# #     "planning_officer": "PCO",
# #     **{f"audit_group{i}": "AuditGroup" for i in range(1, 31)}
# # }
# # AUDIT_GROUP_NUMBERS = {
# #     f"audit_group{i}": i for i in range(1, 31)
# # }
# #
# #
# # # --- Custom CSS Styling ---
# # def load_custom_css():
# #     st.markdown("""
# #     <style>
# #         /* --- Global Styles --- */
# #         body {
# #             font-family: 'Roboto', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; /* Modern font */
# #             background-color: #eef2f7; /* Lighter, softer background */
# #             color: #4A4A4A; /* Darker gray for body text for better contrast */
# #             line-height: 1.6;
# #         }
# #         .stApp {
# #              background: linear-gradient(135deg, #f0f7ff 0%, #cfe7fa 100%); /* Softer blue gradient */
# #         }
# #
# #         /* --- Titles and Headers --- */
# #         .page-main-title { /* For titles outside login box */
# #             font-size: 3em; /* Even bigger */
# #             color: #1A237E; /* Darker, richer blue */
# #             text-align: center;
# #             padding: 30px 0 10px 0;
# #             font-weight: 700;
# #             letter-spacing: 1.5px;
# #             text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
# #         }
# #         .page-app-subtitle { /* For subtitle outside login box */
# #             font-size: 1.3em;
# #             color: #3F51B5; /* Indigo */
# #             text-align: center;
# #             margin-top: -5px;
# #             margin-bottom: 30px;
# #             font-weight: 400;
# #         }
# #         .app-description {
# #             font-size: 1.0em; /* Slightly larger description */
# #             color: #455A64; /* Bluish gray */
# #             text-align: center;
# #             margin-bottom: 25px;
# #             padding: 0 20px;
# #             max-width: 700px; /* Limit width for readability */
# #             margin-left: auto;
# #             margin-right: auto;
# #         }
# #         .sub-header {
# #             font-size: 1.6em;
# #             color: #2779bd;
# #             border-bottom: 3px solid #5dade2;
# #             padding-bottom: 12px;
# #             margin-top: 35px;
# #             margin-bottom: 25px;
# #             font-weight: 600;
# #         }
# #         .card h3 {
# #             margin-top: 0;
# #             color: #1abc9c;
# #             font-size: 1.3em;
# #             font-weight: 600;
# #         }
# #          .card h4 {
# #             color: #2980b9;
# #             font-size: 1.1em;
# #             margin-top: 15px;
# #             margin-bottom: 8px;
# #         }
# #
# #
# #         /* --- Cards --- */
# #         .card {
# #             background-color: #ffffff;
# #             padding: 30px;
# #             border-radius: 12px;
# #             box-shadow: 0 6px 12px rgba(0,0,0,0.08);
# #             margin-bottom: 25px;
# #             border-left: 6px solid #5dade2;
# #         }
# #
# #         /* --- Streamlit Widgets Styling --- */
# #         .stButton>button {
# #             border-radius: 25px;
# #             background-image: linear-gradient(to right, #1abc9c 0%, #16a085 100%);
# #             color: white;
# #             padding: 12px 24px;
# #             font-weight: bold;
# #             border: none;
# #             transition: all 0.3s ease;
# #             box-shadow: 0 2px 4px rgba(0,0,0,0.1);
# #         }
# #         .stButton>button:hover {
# #             background-image: linear-gradient(to right, #16a085 0%, #1abc9c 100%);
# #             transform: translateY(-2px);
# #             box-shadow: 0 4px 8px rgba(0,0,0,0.15);
# #         }
# #         .stButton>button[kind="secondary"] {
# #             background-image: linear-gradient(to right, #e74c3c 0%, #c0392b 100%);
# #         }
# #         .stButton>button[kind="secondary"]:hover {
# #             background-image: linear-gradient(to right, #c0392b 0%, #e74c3c 100%);
# #         }
# #         .stButton>button:disabled {
# #             background-image: none;
# #             background-color: #bdc3c7;
# #             color: #7f8c8d;
# #             box-shadow: none;
# #             transform: none;
# #         }
# #         .stTextInput>div>div>input, .stSelectbox>div>div>div, .stDateInput>div>div>input, .stNumberInput>div>div>input {
# #             border-radius: 8px;
# #             border: 1px solid #ced4da;
# #             padding: 10px;
# #         }
# #         .stTextInput>div>div>input:focus, .stSelectbox>div>div>div:focus-within, .stNumberInput>div>div>input:focus {
# #             border-color: #5dade2;
# #             box-shadow: 0 0 0 0.2rem rgba(93, 173, 226, 0.25);
# #         }
# #         .stFileUploader>div>div>button {
# #             border-radius: 25px;
# #             background-image: linear-gradient(to right, #5dade2 0%, #2980b9 100%);
# #             color: white;
# #             padding: 10px 18px;
# #         }
# #         .stFileUploader>div>div>button:hover {
# #             background-image: linear-gradient(to right, #2980b9 0%, #5dade2 100%);
# #         }
# #
# #
# #         /* --- Login Page Specific --- */
# #         .login-container {
# #             max-width: 500px;
# #             margin: 20px auto; /* Adjusted margin */
# #             padding: 30px;
# #             background-color: #ffffff;
# #             border-radius: 15px;
# #             box-shadow: 0 10px 25px rgba(0,0,0,0.1);
# #         }
# #         .login-container .stButton>button {
# #             background-image: linear-gradient(to right, #34495e 0%, #2c3e50 100%);
# #         }
# #         .login-container .stButton>button:hover {
# #             background-image: linear-gradient(to right, #2c3e50 0%, #34495e 100%);
# #         }
# #         .login-header-text { /* For "User Login" text inside the box */
# #             text-align: center;
# #             color: #1a5276;
# #             font-weight: 600;
# #             font-size: 1.8em;
# #             margin-bottom: 25px;
# #         }
# #         .login-logo {
# #             display: block;
# #             margin-left: auto;
# #             margin-right: auto;
# #             max-width: 70px;
# #             margin-bottom: 15px;
# #             border-radius: 50%;
# #             box-shadow: 0 2px 4px rgba(0,0,0,0.1);
# #         }
# #
# #
# #         /* --- Sidebar Styling --- */
# #         .css-1d391kg {
# #             background-color: #ffffff;
# #             padding: 15px !important;
# #         }
# #         .sidebar .stButton>button {
# #              background-image: linear-gradient(to right, #e74c3c 0%, #c0392b 100%);
# #         }
# #         .sidebar .stButton>button:hover {
# #              background-image: linear-gradient(to right, #c0392b 0%, #e74c3c 100%);
# #         }
# #         .sidebar .stMarkdown > div > p > strong {
# #             color: #2c3e50;
# #         }
# #
# #         /* --- Option Menu Customization --- */
# #         div[data-testid="stOptionMenu"] > ul {
# #             background-color: #ffffff;
# #             border-radius: 25px;
# #             padding: 8px;
# #             box-shadow: 0 2px 5px rgba(0,0,0,0.05);
# #         }
# #         div[data-testid="stOptionMenu"] > ul > li > button {
# #             border-radius: 20px;
# #             margin: 0 5px !important;
# #             border: none !important;
# #             transition: all 0.3s ease;
# #         }
# #         div[data-testid="stOptionMenu"] > ul > li > button.selected {
# #             background-image: linear-gradient(to right, #1abc9c 0%, #16a085 100%);
# #             color: white;
# #             font-weight: bold;
# #             box-shadow: 0 2px 4px rgba(0,0,0,0.1);
# #         }
# #         div[data-testid="stOptionMenu"] > ul > li > button:hover:not(.selected) {
# #             background-color: #e0e0e0;
# #             color: #333;
# #         }
# #
# #         /* --- Links --- */
# #         a {
# #             color: #3498db;
# #             text-decoration: none;
# #             font-weight: 500;
# #         }
# #         a:hover {
# #             text-decoration: underline;
# #             color: #2980b9;
# #         }
# #
# #         /* --- Info/Warning/Error Boxes --- */
# #         .stAlert {
# #             border-radius: 8px;
# #             padding: 15px;
# #             border-left-width: 5px;
# #         }
# #         .stAlert[data-baseweb="notification"][role="alert"] > div:nth-child(2) {
# #              font-size: 1.0em;
# #         }
# #         .stAlert[data-testid="stNotification"] {
# #             box-shadow: 0 2px 10px rgba(0,0,0,0.07);
# #         }
# #         .stAlert[data-baseweb="notification"][kind="info"] { border-left-color: #3498db; }
# #         .stAlert[data-baseweb="notification"][kind="success"] { border-left-color: #2ecc71; }
# #         .stAlert[data-baseweb="notification"][kind="warning"] { border-left-color: #f39c12; }
# #         .stAlert[data-baseweb="notification"][kind="error"] { border-left-color: #e74c3c; }
# #
# #     </style>
# #     """, unsafe_allow_html=True)
# #
# #
# # # --- Google API Authentication and Service Initialization (Modified for Service Account) ---
# # def get_google_services():
# #     """Authenticates using a service account and returns Drive and Sheets services."""
# #     creds = None
# #     if not os.path.exists(CREDENTIALS_FILE):
# #         st.error(f"Service account credentials file ('{CREDENTIALS_FILE}') not found. "
# #                  "Please place your service account JSON key file in the app directory.")
# #         return None, None
# #     try:
# #         creds = service_account.Credentials.from_service_account_file(
# #             CREDENTIALS_FILE, scopes=SCOPES)
# #     except Exception as e:
# #         st.error(f"Failed to load service account credentials: {e}")
# #         return None, None
# #
# #     if not creds:
# #         st.error("Could not initialize credentials from service account file.")
# #         return None, None
# #
# #     try:
# #         drive_service = build('drive', 'v3', credentials=creds)
# #         sheets_service = build('sheets', 'v4', credentials=creds)
# #         return drive_service, sheets_service
# #     except HttpError as error:
# #         st.error(f"An error occurred initializing Google services with service account: {error}")
# #         return None, None
# #     except Exception as e:
# #         st.error(f"An unexpected error occurred with Google services (service account): {e}")
# #         return None, None
# #
# #
# # # --- MCM Period Management ---
# # def load_mcm_periods():
# #     if os.path.exists(MCM_PERIODS_FILE):
# #         with open(MCM_PERIODS_FILE, 'r') as f:
# #             try:
# #                 return json.load(f)
# #             except json.JSONDecodeError:
# #                 return {}
# #     return {}
# #
# #
# # def save_mcm_periods(periods):
# #     with open(MCM_PERIODS_FILE, 'w') as f:
# #         json.dump(periods, f, indent=4)
# #
# #
# # # --- Google Drive Functions ---
# # def set_public_read_permission(drive_service, file_id):
# #     """Sets 'anyone with the link can view' permission for a Drive file/folder."""
# #     try:
# #         permission = {'type': 'anyone', 'role': 'reader'}
# #         drive_service.permissions().create(fileId=file_id, body=permission).execute()
# #         print(f"Permission set to public read for file ID: {file_id}")
# #     except HttpError as error:
# #         st.warning(
# #             f"Could not set public read permission for file ID {file_id}: {error}. Link might require login or manual sharing.")
# #     except Exception as e:
# #         st.warning(f"Unexpected error setting public permission for file ID {file_id}: {e}")
# #
# #
# # def create_drive_folder(drive_service, folder_name):
# #     try:
# #         file_metadata = {
# #             'name': folder_name,
# #             'mimeType': 'application/vnd.google-apps.folder'
# #         }
# #         folder = drive_service.files().create(body=file_metadata, fields='id, webViewLink').execute()
# #         folder_id = folder.get('id')
# #         if folder_id:
# #             set_public_read_permission(drive_service, folder_id)
# #         return folder_id, folder.get('webViewLink')
# #     except HttpError as error:
# #         st.error(f"An error occurred creating Drive folder: {error}")
# #         return None, None
# #
# #
# # def upload_to_drive(drive_service, file_content_or_path, folder_id, filename_on_drive):
# #     try:
# #         file_metadata = {'name': filename_on_drive, 'parents': [folder_id]}
# #         media_body = None
# #
# #         if isinstance(file_content_or_path, str) and os.path.exists(file_content_or_path):
# #             media_body = MediaFileUpload(file_content_or_path, mimetype='application/pdf', resumable=True)
# #         elif isinstance(file_content_or_path, bytes):
# #             fh = BytesIO(file_content_or_path)
# #             media_body = MediaIoBaseUpload(fh, mimetype='application/pdf', resumable=True)
# #         elif isinstance(file_content_or_path, BytesIO):
# #             file_content_or_path.seek(0)
# #             media_body = MediaIoBaseUpload(file_content_or_path, mimetype='application/pdf', resumable=True)
# #         else:
# #             st.error(f"Unsupported file content type for Google Drive upload: {type(file_content_or_path)}")
# #             return None, None
# #
# #         if media_body is None:
# #             st.error("Media body for upload could not be prepared.")
# #             return None, None
# #
# #         request = drive_service.files().create(
# #             body=file_metadata,
# #             media_body=media_body,
# #             fields='id, webViewLink'
# #         )
# #         file = request.execute()
# #         file_id = file.get('id')
# #         if file_id:
# #             set_public_read_permission(drive_service, file_id)
# #         return file_id, file.get('webViewLink')
# #     except HttpError as error:
# #         st.error(f"An API error occurred uploading to Drive: {error}")
# #         return None, None
# #     except Exception as e:
# #         st.error(f"An unexpected error in upload_to_drive: {e}")
# #         return None, None
# #
# #
# # # --- Google Sheets Functions ---
# # def create_spreadsheet(sheets_service, drive_service, title):
# #     try:
# #         spreadsheet_body = {
# #             'properties': {'title': title}
# #         }
# #         spreadsheet = sheets_service.spreadsheets().create(body=spreadsheet_body,
# #                                                            fields='spreadsheetId,spreadsheetUrl').execute()
# #         spreadsheet_id = spreadsheet.get('spreadsheetId')
# #         if spreadsheet_id and drive_service:
# #             set_public_read_permission(drive_service, spreadsheet_id)
# #         return spreadsheet_id, spreadsheet.get('spreadsheetUrl')
# #     except HttpError as error:
# #         st.error(f"An error occurred creating Spreadsheet: {error}")
# #         return None, None
# #     except Exception as e:
# #         st.error(f"An unexpected error occurred creating Spreadsheet: {e}")
# #         return None, None
# #
# #
# # def append_to_spreadsheet(sheets_service, spreadsheet_id, values_to_append):
# #     try:
# #         body = {'values': values_to_append}
# #         sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
# #         sheets = sheet_metadata.get('sheets', '')
# #         first_sheet_title = sheets[0].get("properties", {}).get("title", "Sheet1")
# #
# #         range_to_check = f"{first_sheet_title}!A1:L1"
# #         result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,
# #                                                             range=range_to_check).execute()
# #         header_row_values = result.get('values', [])
# #
# #         if not header_row_values:
# #             header_values_list = [
# #                 "Audit Group Number", "GSTIN", "Trade Name", "Category",
# #                 "Total Amount Detected (Overall Rs)", "Total Amount Recovered (Overall Rs)",
# #                 "Audit Para Number", "Audit Para Heading",
# #                 "Revenue Involved (Lakhs Rs)", "Revenue Recovered (Lakhs Rs)",
# #                 "DAR PDF URL", "Record Created Date"
# #             ]
# #             sheets_service.spreadsheets().values().append(
# #                 spreadsheetId=spreadsheet_id,
# #                 range=f"{first_sheet_title}!A1",
# #                 valueInputOption='USER_ENTERED',
# #                 body={'values': header_values_list}
# #             ).execute()
# #
# #         append_result = sheets_service.spreadsheets().values().append(
# #             spreadsheetId=spreadsheet_id,
# #             range=f"{first_sheet_title}!A1",
# #             valueInputOption='USER_ENTERED',
# #             body=body
# #         ).execute()
# #         return append_result
# #     except HttpError as error:
# #         st.error(f"An error occurred appending to Spreadsheet: {error}")
# #         return None
# #
# #
# # def read_from_spreadsheet(sheets_service, spreadsheet_id, sheet_name="Sheet1"):
# #     try:
# #         result = sheets_service.spreadsheets().values().get(
# #             spreadsheetId=spreadsheet_id,
# #             range=sheet_name
# #         ).execute()
# #         values = result.get('values', [])
# #         if not values:
# #             return pd.DataFrame()
# #         else:
# #             expected_cols = [
# #                 "Audit Group Number", "GSTIN", "Trade Name", "Category",
# #                 "Total Amount Detected (Overall Rs)", "Total Amount Recovered (Overall Rs)",
# #                 "Audit Para Number", "Audit Para Heading",
# #                 "Revenue Involved (Lakhs Rs)", "Revenue Recovered (Lakhs Rs)",
# #                 "DAR PDF URL", "Record Created Date"
# #             ]
# #             if values and values[0] == expected_cols:
# #                 return pd.DataFrame(values[1:], columns=values[0])
# #             else:
# #                 st.warning(
# #                     f"Spreadsheet '{sheet_name}' headers do not match expected or are missing. Attempting to load data.")
# #                 if values:
# #                     num_cols = len(values[0]) if values else 0
# #                     df_cols = values[0] if values and len(values[0]) == num_cols else [f"Col_{j + 1}" for j in
# #                                                                                        range(num_cols)]
# #                     data_start_row = 1 if values and values[0] == expected_cols else 0
# #                     return pd.DataFrame(values[data_start_row:], columns=df_cols if data_start_row == 1 else None)
# #                 else:
# #                     return pd.DataFrame()
# #     except HttpError as error:
# #         st.error(f"An error occurred reading from Spreadsheet: {error}")
# #         return pd.DataFrame()
# #
# #
# # def delete_spreadsheet_rows(sheets_service, spreadsheet_id, sheet_id_gid, row_indices_to_delete):
# #     """Deletes specified rows from a sheet. Row indices are 0-based."""
# #     if not row_indices_to_delete:
# #         return True  # Nothing to delete
# #
# #     requests = []
# #     # Sort indices in descending order to avoid shifting issues during batch deletion
# #     for row_index in sorted(row_indices_to_delete, reverse=True):
# #         requests.append({
# #             "deleteDimension": {
# #                 "range": {
# #                     "sheetId": sheet_id_gid,  # This is the GID of the sheet, not its name
# #                     "dimension": "ROWS",
# #                     "startIndex": row_index,  # 0-indexed
# #                     "endIndex": row_index + 1
# #                 }
# #             }
# #         })
# #
# #     if requests:
# #         try:
# #             body = {'requests': requests}
# #             sheets_service.spreadsheets().batchUpdate(
# #                 spreadsheetId=spreadsheet_id, body=body).execute()
# #             return True
# #         except HttpError as error:
# #             st.error(f"An error occurred deleting rows from Spreadsheet: {error}")
# #             return False
# #     return True
# #
# #
# # # --- Gemini Data Extraction with Retry ---
# # def get_structured_data_with_gemini(api_key: str, text_content: str, max_retries=2) -> ParsedDARReport:
# #     """
# #     Calls Gemini API with the full PDF text and parses the response.
# #     Includes retry logic for JSONDecodeError.
# #     """
# #     if not api_key or api_key == "YOUR_API_KEY_HERE":
# #         return ParsedDARReport(parsing_errors="Gemini API Key not configured in app.py.")
# #     if text_content.startswith("Error processing PDF with pdfplumber:"):
# #         return ParsedDARReport(parsing_errors=text_content)
# #
# #     genai.configure(api_key=api_key)
# #     model = genai.GenerativeModel('gemini-1.5-flash-latest')
# #
# #     prompt = f"""
# #     You are an expert GST audit report analyst. Based on the following FULL text from a Departmental Audit Report (DAR),
# #     where all text from all pages, including tables, is provided, extract the specified information
# #     and structure it as a JSON object. Focus on identifying narrative sections for audit para details,
# #     even if they are intermingled with tabular data. Notes like "[INFO: ...]" in the text are for context only.
# #
# #     The JSON object should follow this structure precisely:
# #     {{
# #       "header": {{
# #         "audit_group_number": "integer or null (e.g., if 'Group-VI' or 'Gr 6', extract 6; must be between 1 and 30)",
# #         "gstin": "string or null",
# #         "trade_name": "string or null",
# #         "category": "string ('Large', 'Medium', 'Small') or null",
# #         "total_amount_detected_overall_rs": "float or null (numeric value in Rupees)",
# #         "total_amount_recovered_overall_rs": "float or null (numeric value in Rupees)"
# #       }},
# #       "audit_paras": [
# #         {{
# #           "audit_para_number": "integer or null (primary number from para heading, e.g., for 'Para-1...' use 1; must be between 1 and 50)",
# #           "audit_para_heading": "string or null (the descriptive title of the para)",
# #           "revenue_involved_lakhs_rs": "float or null (numeric value in Lakhs of Rupees, e.g., Rs. 50,000 becomes 0.5)",
# #           "revenue_recovered_lakhs_rs": "float or null (numeric value in Lakhs of Rupees)"
# #         }}
# #       ],
# #       "parsing_errors": "string or null (any notes about parsing issues, or if extraction is incomplete)"
# #     }}
# #
# #     Key Instructions:
# #     1.  Header Information: Extract `audit_group_number` (as integer 1-30, e.g., 'Group-VI' becomes 6), `gstin`, `trade_name`, `category`, `total_amount_detected_overall_rs`, `total_amount_recovered_overall_rs`.
# #     2.  Audit Paras: Identify each distinct para. Extract `audit_para_number` (as integer 1-50), `audit_para_heading`, `revenue_involved_lakhs_rs` (converted to Lakhs), `revenue_recovered_lakhs_rs` (converted to Lakhs).
# #     3.  Use null for missing values. Monetary values as float.
# #     4.  If no audit paras found, `audit_paras` should be an empty list [].
# #
# #     DAR Text Content:
# #     --- START OF DAR TEXT ---
# #     {text_content}
# #     --- END OF DAR TEXT ---
# #
# #     Provide ONLY the JSON object as your response. Do not include any explanatory text before or after the JSON.
# #     """
# #
# #     attempt = 0
# #     last_exception = None
# #     while attempt <= max_retries:
# #         attempt += 1
# #         print(f"\n--- Calling Gemini (Attempt {attempt}/{max_retries + 1}) ---")
# #         try:
# #             response = model.generate_content(prompt)
# #
# #             cleaned_response_text = response.text.strip()
# #             if cleaned_response_text.startswith("```json"):
# #                 cleaned_response_text = cleaned_response_text[7:]
# #             elif cleaned_response_text.startswith("`json"):
# #                 cleaned_response_text = cleaned_response_text[6:]
# #             if cleaned_response_text.endswith("```"):
# #                 cleaned_response_text = cleaned_response_text[:-3]
# #
# #             if not cleaned_response_text:
# #                 error_message = f"Gemini returned an empty response on attempt {attempt}."
# #                 print(error_message)
# #                 last_exception = ValueError(error_message)
# #                 if attempt > max_retries:
# #                     return ParsedDARReport(parsing_errors=error_message)
# #                 time.sleep(1 + attempt)
# #                 continue
# #
# #             json_data = json.loads(cleaned_response_text)
# #             if "header" not in json_data or "audit_paras" not in json_data:
# #                 error_message = f"Gemini response (Attempt {attempt}) missing 'header' or 'audit_paras' key. Response: {cleaned_response_text[:500]}"
# #                 print(error_message)
# #                 last_exception = ValueError(error_message)
# #                 if attempt > max_retries:
# #                     return ParsedDARReport(parsing_errors=error_message)
# #                 time.sleep(1 + attempt)
# #                 continue
# #
# #             parsed_report = ParsedDARReport(**json_data)
# #             print(f"Gemini call (Attempt {attempt}) successful. Paras found: {len(parsed_report.audit_paras)}")
# #             if parsed_report.audit_paras:
# #                 for idx, para_obj in enumerate(parsed_report.audit_paras):
# #                     if not para_obj.audit_para_heading:
# #                         print(
# #                             f"  Note: Para {idx + 1} (Number: {para_obj.audit_para_number}) has a missing heading from Gemini.")
# #             return parsed_report
# #         except json.JSONDecodeError as e:
# #             raw_response_text = "No response text available"
# #             if 'response' in locals() and hasattr(response, 'text'):
# #                 raw_response_text = response.text
# #             error_message = f"Gemini output (Attempt {attempt}) was not valid JSON: {e}. Response: '{raw_response_text[:1000]}...'"
# #             print(error_message)
# #             last_exception = e
# #             if attempt > max_retries:
# #                 return ParsedDARReport(parsing_errors=error_message)
# #             time.sleep(attempt * 2)
# #         except Exception as e:
# #             raw_response_text = "No response text available"
# #             if 'response' in locals() and hasattr(response, 'text'):
# #                 raw_response_text = response.text
# #             error_message = f"Error (Attempt {attempt}) during Gemini/Pydantic: {type(e).__name__} - {e}. Response: {raw_response_text[:500]}"
# #             print(error_message)
# #             last_exception = e
# #             if attempt > max_retries:
# #                 return ParsedDARReport(parsing_errors=error_message)
# #             time.sleep(attempt * 2)
# #
# #     return ParsedDARReport(
# #         parsing_errors=f"Gemini call failed after {max_retries + 1} attempts. Last error: {last_exception}")
# #
# #
# # # --- Streamlit App UI and Logic ---
# # st.set_page_config(layout="wide", page_title="e-MCM App - GST Audit 1")
# # load_custom_css()
# #
# # # --- Login ---
# # if 'logged_in' not in st.session_state:
# #     st.session_state.logged_in = False
# #     st.session_state.username = ""
# #     st.session_state.role = ""
# #     st.session_state.audit_group_no = None
# #     st.session_state.ag_current_extracted_data = []
# #     st.session_state.ag_pdf_drive_url = None
# #     st.session_state.ag_validation_errors = []
# #     st.session_state.ag_editor_data = pd.DataFrame()
# #     st.session_state.ag_current_mcm_key = None
# #     st.session_state.ag_current_uploaded_file_name = None
# #
# #
# # def login_page():
# #     # The main titles are now outside the login container
# #     st.markdown("<div class='page-main-title'>e-MCM App</div>", unsafe_allow_html=True)
# #     st.markdown("<h2 class='page-app-subtitle'>GST Audit 1 Commissionerate</h2>", unsafe_allow_html=True)
# #
# #     with st.container():
# #         st.markdown("<div class='login-container'>", unsafe_allow_html=True)  # White box for login form
# #         st.markdown(
# #             "<div class='login-header'><img src='https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/Indian_Ministry_of_Finance_logo.svg/1200px-Indian_Ministry_of_Finance_logo.svg.png' alt='Logo' class='login-logo'></div>",
# #             unsafe_allow_html=True)
# #         st.markdown("<h2 class='login-header-text'>User Login</h2>", unsafe_allow_html=True)
# #         st.markdown("""
# #         <div class='app-description'>
# #             Welcome! This platform streamlines Draft Audit Report (DAR) collection and processing.
# #             PCOs manage MCM periods; Audit Groups upload DARs for AI-powered data extraction.
# #         </div>
# #         """, unsafe_allow_html=True)
# #
# #         username = st.text_input("Username", key="login_username_styled", placeholder="Enter your username")
# #         password = st.text_input("Password", type="password", key="login_password_styled",
# #                                  placeholder="Enter your password")
# #
# #         if st.button("Login", key="login_button_styled", use_container_width=True):
# #             if username in USER_CREDENTIALS and USER_CREDENTIALS[username] == password:
# #                 st.session_state.logged_in = True
# #                 st.session_state.username = username
# #                 st.session_state.role = USER_ROLES[username]
# #                 if st.session_state.role == "AuditGroup":
# #                     st.session_state.audit_group_no = AUDIT_GROUP_NUMBERS[username]
# #                 st.success(f"Logged in as {username} ({st.session_state.role})")
# #                 st.rerun()
# #             else:
# #                 st.error("Invalid username or password")
# #         st.markdown("</div>", unsafe_allow_html=True)
# #
# #
# # # --- PCO Dashboard ---
# # def pco_dashboard(drive_service, sheets_service):
# #     st.markdown("<div class='sub-header'>Planning & Coordination Officer Dashboard</div>", unsafe_allow_html=True)
# #     mcm_periods = load_mcm_periods()
# #
# #     with st.sidebar:
# #         st.image(
# #             "https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/Indian_Ministry_of_Finance_logo.svg/1200px-Indian_Ministry_of_Finance_logo.svg.png",
# #             width=80)
# #         st.markdown(f"**User:** {st.session_state.username}")
# #         st.markdown(f"**Role:** {st.session_state.role}")
# #         if st.button("Logout", key="pco_logout_styled", use_container_width=True):
# #             st.session_state.logged_in = False
# #             st.session_state.username = ""
# #             st.session_state.role = ""
# #             st.rerun()
# #         st.markdown("---")
# #
# #     selected_tab = option_menu(
# #         menu_title=None,
# #         options=["Create MCM Period", "Manage MCM Periods", "View Uploaded Reports", "Visualizations"],
# #         icons=["calendar-plus-fill", "sliders", "eye-fill", "bar-chart-fill"],
# #         menu_icon="gear-wide-connected",
# #         default_index=0,
# #         orientation="horizontal",
# #         styles={
# #             "container": {"padding": "5px !important", "background-color": "#e9ecef"},
# #             "icon": {"color": "#007bff", "font-size": "20px"},
# #             "nav-link": {"font-size": "16px", "text-align": "center", "margin": "0px", "--hover-color": "#d1e7fd"},
# #             "nav-link-selected": {"background-color": "#007bff", "color": "white"},
# #         }
# #     )
# #
# #     st.markdown("<div class='card'>", unsafe_allow_html=True)
# #
# #     if selected_tab == "Create MCM Period":
# #         st.markdown("<h3>Create New MCM Period</h3>", unsafe_allow_html=True)
# #         current_year = datetime.datetime.now().year
# #         years = list(range(current_year - 1, current_year + 3))
# #         months = ["January", "February", "March", "April", "May", "June",
# #                   "July", "August", "September", "October", "November", "December"]
# #
# #         col1, col2 = st.columns(2)
# #         with col1:
# #             selected_year = st.selectbox("Select Year", options=years, index=years.index(current_year), key="pco_year")
# #         with col2:
# #             selected_month_name = st.selectbox("Select Month", options=months, index=datetime.datetime.now().month - 1,
# #                                                key="pco_month")
# #
# #         selected_month_num = months.index(selected_month_name) + 1
# #         period_key = f"{selected_year}-{selected_month_num:02d}"
# #
# #         if period_key in mcm_periods:
# #             st.warning(f"MCM Period for {selected_month_name} {selected_year} already exists.")
# #             st.markdown(
# #                 f"**Drive Folder:** <a href='{mcm_periods[period_key]['drive_folder_url']}' target='_blank'>Open Folder</a>",
# #                 unsafe_allow_html=True)
# #             st.markdown(
# #                 f"**Spreadsheet:** <a href='{mcm_periods[period_key]['spreadsheet_url']}' target='_blank'>Open Sheet</a>",
# #                 unsafe_allow_html=True)
# #         else:
# #             if st.button(f"Create MCM for {selected_month_name} {selected_year}", key="pco_create_mcm",
# #                          use_container_width=True):
# #                 if not drive_service or not sheets_service:
# #                     st.error("Google Services not available. Cannot create MCM period.")
# #                 else:
# #                     with st.spinner("Creating Google Drive folder and Spreadsheet..."):
# #                         folder_name = f"MCM_DARs_{selected_month_name}_{selected_year}"
# #                         spreadsheet_title = f"MCM_Audit_Paras_{selected_month_name}_{selected_year}"
# #
# #                         folder_id, folder_url = create_drive_folder(drive_service, folder_name)
# #                         sheet_id, sheet_url = create_spreadsheet(sheets_service, drive_service, spreadsheet_title)
# #
# #                         if folder_id and sheet_id:
# #                             mcm_periods[period_key] = {
# #                                 "year": selected_year, "month_num": selected_month_num,
# #                                 "month_name": selected_month_name,
# #                                 "drive_folder_id": folder_id, "drive_folder_url": folder_url,
# #                                 "spreadsheet_id": sheet_id, "spreadsheet_url": sheet_url,
# #                                 "active": True
# #                             }
# #                             save_mcm_periods(mcm_periods)
# #                             st.success(f"Successfully created MCM period for {selected_month_name} {selected_year}!")
# #                             st.markdown(f"**Drive Folder:** <a href='{folder_url}' target='_blank'>Open Folder</a>",
# #                                         unsafe_allow_html=True)
# #                             st.markdown(f"**Spreadsheet:** <a href='{sheet_url}' target='_blank'>Open Sheet</a>",
# #                                         unsafe_allow_html=True)
# #                             st.balloons()
# #                             time.sleep(0.5)
# #                             st.rerun()
# #                         else:
# #                             st.error("Failed to create Drive folder or Spreadsheet.")
# #
# #     elif selected_tab == "Manage MCM Periods":
# #         st.markdown("<h3>Manage Existing MCM Periods</h3>", unsafe_allow_html=True)
# #         if not mcm_periods:
# #             st.info("No MCM periods created yet.")
# #         else:
# #             sorted_periods_keys = sorted(mcm_periods.keys(), reverse=True)
# #
# #             for period_key in sorted_periods_keys:
# #                 data = mcm_periods[period_key]
# #                 st.markdown(f"<h4>{data['month_name']} {data['year']}</h4>", unsafe_allow_html=True)
# #                 col1, col2, col3, col4 = st.columns([2, 2, 1, 2])
# #                 with col1:
# #                     st.markdown(f"<a href='{data['drive_folder_url']}' target='_blank'>Open Drive Folder</a>",
# #                                 unsafe_allow_html=True)
# #                 with col2:
# #                     st.markdown(f"<a href='{data['spreadsheet_url']}' target='_blank'>Open Spreadsheet</a>",
# #                                 unsafe_allow_html=True)
# #                 with col3:
# #                     is_active = data.get("active", False)
# #                     new_status = st.checkbox("Active", value=is_active, key=f"active_{period_key}_styled_manage")
# #                     if new_status != is_active:
# #                         mcm_periods[period_key]["active"] = new_status
# #                         save_mcm_periods(mcm_periods)
# #                         st.success(f"Status for {data['month_name']} {data['year']} updated.")
# #                         st.rerun()
# #                 with col4:
# #                     if st.button("Delete Period Record", key=f"delete_mcm_{period_key}", type="secondary"):
# #                         st.session_state.period_to_delete = period_key
# #                         st.session_state.show_delete_confirm = True
# #                         st.rerun()
# #                 st.markdown("---")
# #
# #             if st.session_state.get('show_delete_confirm', False) and st.session_state.get('period_to_delete'):
# #                 period_key_to_delete = st.session_state.period_to_delete
# #                 period_data_to_delete = mcm_periods.get(period_key_to_delete, {})
# #
# #                 with st.form(key=f"delete_confirm_form_{period_key_to_delete}"):
# #                     st.warning(
# #                         f"Are you sure you want to delete the MCM period record for **{period_data_to_delete.get('month_name')} {period_data_to_delete.get('year')}** from this application?")
# #                     st.caption(
# #                         "This action only removes the period from the app's tracking. It **does NOT delete** the actual Google Drive folder or the Google Spreadsheet associated with it. Those must be managed manually in Google Drive/Sheets if needed.")
# #                     pco_password_confirm = st.text_input("Enter your PCO password to confirm deletion:",
# #                                                          type="password",
# #                                                          key=f"pco_pass_confirm_{period_key_to_delete}")
# #
# #                     col_form1, col_form2 = st.columns(2)
# #                     with col_form1:
# #                         submitted_delete = st.form_submit_button("Yes, Delete This Period Record",
# #                                                                  use_container_width=True)
# #                     with col_form2:
# #                         if st.form_submit_button("Cancel", type="secondary", use_container_width=True):
# #                             st.session_state.show_delete_confirm = False
# #                             st.session_state.period_to_delete = None
# #                             st.rerun()
# #
# #                     if submitted_delete:
# #                         if pco_password_confirm == USER_CREDENTIALS.get("planning_officer"):
# #                             del mcm_periods[period_key_to_delete]
# #                             save_mcm_periods(mcm_periods)
# #                             st.success(
# #                                 f"MCM period record for {period_data_to_delete.get('month_name')} {period_data_to_delete.get('year')} has been deleted from the app.")
# #                             st.session_state.show_delete_confirm = False
# #                             st.session_state.period_to_delete = None
# #                             st.rerun()
# #                         else:
# #                             st.error("Incorrect password. Deletion cancelled.")
# #
# #     elif selected_tab == "View Uploaded Reports":
# #         st.markdown("<h3>View Uploaded Reports Summary</h3>", unsafe_allow_html=True)
# #         active_periods = {k: v for k, v in mcm_periods.items()}
# #         if not active_periods:
# #             st.info("No MCM periods created yet to view reports.")
# #         else:
# #             period_options = [f"{p_data['month_name']} {p_data['year']}" for p_key, p_data in
# #                               sorted(active_periods.items(), key=lambda item: item[0], reverse=True)]
# #             selected_period_display = st.selectbox("Select MCM Period to View", options=period_options,
# #                                                    key="pco_view_period")
# #
# #             if selected_period_display:
# #                 selected_period_key = None
# #                 for p_key, p_data in active_periods.items():
# #                     if f"{p_data['month_name']} {p_data['year']}" == selected_period_display:
# #                         selected_period_key = p_key
# #                         break
# #
# #                 if selected_period_key and sheets_service:
# #                     sheet_id = mcm_periods[selected_period_key]['spreadsheet_id']
# #                     st.markdown(f"**Fetching data for {selected_period_display}...**")
# #                     with st.spinner("Loading data from Google Sheet..."):
# #                         df = read_from_spreadsheet(sheets_service, sheet_id)
# #
# #                     if not df.empty:
# #                         st.markdown("<h4>Summary of Uploads:</h4>", unsafe_allow_html=True)
# #                         if 'Audit Group Number' in df.columns:
# #                             try:
# #                                 df['Audit Group Number'] = pd.to_numeric(df['Audit Group Number'], errors='coerce')
# #                                 df.dropna(subset=['Audit Group Number'], inplace=True)
# #
# #                                 dars_per_group = df.groupby('Audit Group Number')['DAR PDF URL'].nunique().reset_index(
# #                                     name='DARs Uploaded')
# #                                 st.write("**DARs Uploaded per Audit Group:**")
# #                                 st.dataframe(dars_per_group, use_container_width=True)
# #
# #                                 paras_per_group = df.groupby('Audit Group Number').size().reset_index(
# #                                     name='Total Para Entries')
# #                                 st.write("**Total Para Entries per Audit Group:**")
# #                                 st.dataframe(paras_per_group, use_container_width=True)
# #
# #                                 st.markdown("<h4>Detailed Data:</h4>", unsafe_allow_html=True)
# #                                 st.dataframe(df, use_container_width=True)
# #
# #                             except Exception as e:
# #                                 st.error(f"Error processing data for summary: {e}")
# #                                 st.write("Raw Data:")
# #                                 st.dataframe(df, use_container_width=True)
# #                         else:
# #                             st.warning("Spreadsheet does not contain 'Audit Group Number' column for summary.")
# #                             st.dataframe(df, use_container_width=True)
# #                     else:
# #                         st.info(f"No data found in the spreadsheet for {selected_period_display}.")
# #                 elif not sheets_service:
# #                     st.error("Google Sheets service not available.")
# #
# #     elif selected_tab == "Visualizations":
# #         st.markdown("<h3>Data Visualizations</h3>", unsafe_allow_html=True)
# #         all_mcm_periods = {k: v for k, v in mcm_periods.items()}
# #         if not all_mcm_periods:
# #             st.info("No MCM periods created yet to visualize data.")
# #         else:
# #             viz_period_options = [f"{p_data['month_name']} {p_data['year']}" for p_key, p_data in
# #                                   sorted(all_mcm_periods.items(), key=lambda item: item[0], reverse=True)]
# #             selected_viz_period_display = st.selectbox("Select MCM Period for Visualization",
# #                                                        options=viz_period_options, key="pco_viz_period")
# #
# #             if selected_viz_period_display and sheets_service:
# #                 selected_viz_period_key = None
# #                 for p_key, p_data in all_mcm_periods.items():
# #                     if f"{p_data['month_name']} {p_data['year']}" == selected_viz_period_display:
# #                         selected_viz_period_key = p_key
# #                         break
# #
# #                 if selected_viz_period_key:
# #                     sheet_id_viz = all_mcm_periods[selected_viz_period_key]['spreadsheet_id']
# #                     st.markdown(f"**Fetching data for {selected_viz_period_display} visualizations...**")
# #                     with st.spinner("Loading data..."):
# #                         df_viz = read_from_spreadsheet(sheets_service, sheet_id_viz)
# #
# #                     if not df_viz.empty:
# #                         # Data Cleaning for numeric columns
# #                         amount_cols = ['Total Amount Detected (Overall Rs)', 'Total Amount Recovered (Overall Rs)',
# #                                        'Revenue Involved (Lakhs Rs)', 'Revenue Recovered (Lakhs Rs)']
# #                         for col in amount_cols:
# #                             if col in df_viz.columns:
# #                                 df_viz[col] = pd.to_numeric(df_viz[col], errors='coerce').fillna(0)
# #                         if 'Audit Group Number' in df_viz.columns:
# #                             df_viz['Audit Group Number'] = pd.to_numeric(df_viz['Audit Group Number'],
# #                                                                          errors='coerce').fillna(0).astype(int)
# #
# #                         st.markdown("---")
# #                         st.markdown("<h4>Group-wise Performance</h4>", unsafe_allow_html=True)
# #
# #                         # Top 5 Groups by Total Detection
# #                         if 'Total Amount Detected (Overall Rs)' in df_viz.columns and 'Audit Group Number' in df_viz.columns:
# #                             top_detection_groups = df_viz.groupby('Audit Group Number')[
# #                                 'Total Amount Detected (Overall Rs)'].sum().nlargest(5)
# #                             if not top_detection_groups.empty:
# #                                 st.write("**Top 5 Groups by Total Detection Amount (Rs):**")
# #                                 st.bar_chart(top_detection_groups)
# #                             else:
# #                                 st.info("Not enough data for 'Top Detection Groups' chart.")
# #
# #                         # Top 5 Groups by Total Realisation
# #                         if 'Total Amount Recovered (Overall Rs)' in df_viz.columns and 'Audit Group Number' in df_viz.columns:
# #                             top_recovery_groups = df_viz.groupby('Audit Group Number')[
# #                                 'Total Amount Recovered (Overall Rs)'].sum().nlargest(5)
# #                             if not top_recovery_groups.empty:
# #                                 st.write("**Top 5 Groups by Total Realisation Amount (Rs):**")
# #                                 st.bar_chart(top_recovery_groups)
# #                             else:
# #                                 st.info("Not enough data for 'Top Realisation Groups' chart.")
# #
# #                         # Top 5 Groups by Recovery/Detection Ratio
# #                         if 'Total Amount Detected (Overall Rs)' in df_viz.columns and \
# #                                 'Total Amount Recovered (Overall Rs)' in df_viz.columns and \
# #                                 'Audit Group Number' in df_viz.columns:
# #                             group_summary = df_viz.groupby('Audit Group Number').agg(
# #                                 Total_Detected=('Total Amount Detected (Overall Rs)', 'sum'),
# #                                 Total_Recovered=('Total Amount Recovered (Overall Rs)', 'sum')
# #                             ).reset_index()
# #                             group_summary['Recovery_Ratio'] = group_summary.apply(
# #                                 lambda row: (row['Total_Recovered'] / row['Total_Detected']) * 100 if row[
# #                                                                                                           'Total_Detected'] > 0 else 0,
# #                                 axis=1
# #                             )
# #                             top_ratio_groups = group_summary.nlargest(5, 'Recovery_Ratio')[
# #                                 ['Audit Group Number', 'Recovery_Ratio']].set_index('Audit Group Number')
# #                             if not top_ratio_groups.empty:
# #                                 st.write("**Top 5 Groups by Recovery/Detection Ratio (%):**")
# #                                 st.bar_chart(top_ratio_groups)
# #                             else:
# #                                 st.info("Not enough data for 'Top Recovery Ratio Groups' chart.")
# #
# #                         st.markdown("---")
# #                         st.markdown("<h4>Para-wise Performance</h4>", unsafe_allow_html=True)
# #                         num_paras_to_show = st.number_input("Select N for Top N Paras:", min_value=1, max_value=20,
# #                                                             value=5, step=1, key="top_n_paras_viz")
# #
# #                         # Top N Detection Paras
# #                         if 'Revenue Involved (Lakhs Rs)' in df_viz.columns:
# #                             top_detection_paras = df_viz.nlargest(num_paras_to_show, 'Revenue Involved (Lakhs Rs)')
# #                             if not top_detection_paras.empty:
# #                                 st.write(f"**Top {num_paras_to_show} Detection Paras (by Revenue Involved):**")
# #                                 st.dataframe(top_detection_paras[
# #                                                  ['Audit Group Number', 'Trade Name', 'Audit Para Number',
# #                                                   'Audit Para Heading', 'Revenue Involved (Lakhs Rs)']],
# #                                              use_container_width=True)
# #                             else:
# #                                 st.info("Not enough data for 'Top Detection Paras' list.")
# #
# #                         # Top N Realisation Paras
# #                         if 'Revenue Recovered (Lakhs Rs)' in df_viz.columns:
# #                             top_recovery_paras = df_viz.nlargest(num_paras_to_show, 'Revenue Recovered (Lakhs Rs)')
# #                             if not top_recovery_paras.empty:
# #                                 st.write(f"**Top {num_paras_to_show} Realisation Paras (by Revenue Recovered):**")
# #                                 st.dataframe(top_recovery_paras[
# #                                                  ['Audit Group Number', 'Trade Name', 'Audit Para Number',
# #                                                   'Audit Para Heading', 'Revenue Recovered (Lakhs Rs)']],
# #                                              use_container_width=True)
# #                             else:
# #                                 st.info("Not enough data for 'Top Realisation Paras' list.")
# #                     else:
# #                         st.info(
# #                             f"No data found in the spreadsheet for {selected_viz_period_display} to generate visualizations.")
# #                 elif not sheets_service:
# #                     st.error("Google Sheets service not available for visualization.")
# #
# #     st.markdown("</div>", unsafe_allow_html=True)
# #
# #
# # # --- Validation Function (reusable) ---
# # MANDATORY_FIELDS_FOR_SHEET = {
# #     "audit_group_number": "Audit Group Number",
# #     "gstin": "GSTIN",
# #     "trade_name": "Trade Name",
# #     "category": "Category",
# #     "total_amount_detected_overall_rs": "Total Amount Detected (Overall Rs)",
# #     "total_amount_recovered_overall_rs": "Total Amount Recovered (Overall Rs)",
# #     "audit_para_number": "Audit Para Number",
# #     "audit_para_heading": "Audit Para Heading",
# #     "revenue_involved_lakhs_rs": "Revenue Involved (Lakhs Rs)",
# #     "revenue_recovered_lakhs_rs": "Revenue Recovered (Lakhs Rs)"
# # }
# # VALID_CATEGORIES = ["Large", "Medium", "Small"]
# #
# #
# # def validate_data_for_sheet(data_df_to_validate):
# #     validation_errors = []
# #     if data_df_to_validate.empty:
# #         validation_errors.append("No data to validate. Please extract or enter data first.")
# #         return validation_errors
# #
# #     # Per-row validations
# #     for index, row in data_df_to_validate.iterrows():
# #         row_display_id = f"Row {index + 1} (Para: {row.get('audit_para_number', 'N/A')})"
# #
# #         # Mandatory field checks
# #         for field_key, field_name in MANDATORY_FIELDS_FOR_SHEET.items():
# #             value = row.get(field_key)
# #             is_missing = False
# #             if value is None:
# #                 is_missing = True
# #             elif isinstance(value, str) and not value.strip():
# #                 is_missing = True
# #             elif pd.isna(value):
# #                 is_missing = True
# #
# #             if is_missing:
# #                 if field_key in ["audit_para_number", "audit_para_heading", "revenue_involved_lakhs_rs",
# #                                  "revenue_recovered_lakhs_rs"]:
# #                     if row.get('audit_para_heading', "").startswith("N/A - Header Info Only") and pd.isna(
# #                             row.get('audit_para_number')):
# #                         continue
# #                 validation_errors.append(f"{row_display_id}: '{field_name}' is missing or empty.")
# #
# #         # Category specific validation
# #         category_val = row.get('category')
# #         if pd.notna(category_val) and category_val.strip() and category_val not in VALID_CATEGORIES:
# #             validation_errors.append(
# #                 f"{row_display_id}: 'Category' ('{category_val}') is invalid. Must be one of {VALID_CATEGORIES} (case-sensitive)."
# #             )
# #         elif pd.isna(category_val) or (isinstance(category_val, str) and not category_val.strip()):
# #             if "category" in MANDATORY_FIELDS_FOR_SHEET:  # Check if category was in mandatory list
# #                 validation_errors.append(f"{row_display_id}: 'Category' is missing. Please select a valid category.")
# #
# #     # Category Consistency Check per Trade Name
# #     if 'trade_name' in data_df_to_validate.columns and 'category' in data_df_to_validate.columns:
# #         trade_name_categories = {}
# #         for index, row in data_df_to_validate.iterrows():
# #             trade_name = row.get('trade_name')
# #             category = row.get('category')
# #             if pd.notna(trade_name) and trade_name.strip() and \
# #                     pd.notna(category) and category.strip() and category in VALID_CATEGORIES:
# #                 if trade_name not in trade_name_categories:
# #                     trade_name_categories[trade_name] = set()
# #                 trade_name_categories[trade_name].add(category)
# #
# #         for tn, cats in trade_name_categories.items():
# #             if len(cats) > 1:
# #                 validation_errors.append(
# #                     f"Consistency Error: Trade Name '{tn}' has multiple valid categories assigned: {', '.join(sorted(list(cats)))}. Please ensure a single, consistent category for this trade name across all its para entries."
# #                 )
# #     return sorted(list(set(validation_errors)))
# #
# #
# # # --- Audit Group Dashboard ---
# # def audit_group_dashboard(drive_service, sheets_service):
# #     st.markdown(f"<div class='sub-header'>Audit Group {st.session_state.audit_group_no} Dashboard</div>",
# #                 unsafe_allow_html=True)
# #     mcm_periods = load_mcm_periods()
# #
# #     active_periods = {k: v for k, v in mcm_periods.items() if v.get("active")}
# #
# #     with st.sidebar:
# #         st.image(
# #             "https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/Indian_Ministry_of_Finance_logo.svg/1200px-Indian_Ministry_of_Finance_logo.svg.png",
# #             width=80)
# #         st.markdown(f"**User:** {st.session_state.username}")
# #         st.markdown(f"**Group No:** {st.session_state.audit_group_no}")
# #         if st.button("Logout", key="ag_logout_styled", use_container_width=True):
# #             for key in ['ag_current_extracted_data', 'ag_pdf_drive_url', 'ag_validation_errors', 'ag_editor_data',
# #                         'ag_current_mcm_key', 'ag_current_uploaded_file_name', 'ag_row_to_delete_details',
# #                         'ag_show_delete_confirm']:
# #                 if key in st.session_state:
# #                     del st.session_state[key]
# #             st.session_state.logged_in = False;
# #             st.session_state.username = "";
# #             st.session_state.role = "";
# #             st.session_state.audit_group_no = None
# #             st.rerun()
# #         st.markdown("---")
# #
# #     selected_tab = option_menu(
# #         menu_title=None,
# #         options=["Upload DAR for MCM", "View My Uploaded DARs", "Delete My DAR Entries"],  # Changed Tab Name
# #         icons=["cloud-upload-fill", "eye-fill", "trash2-fill"],  # Updated Icon
# #         menu_icon="person-workspace",
# #         default_index=0,
# #         orientation="horizontal",
# #         styles={
# #             "container": {"padding": "5px !important", "background-color": "#e9ecef"},
# #             "icon": {"color": "#28a745", "font-size": "20px"},
# #             "nav-link": {"font-size": "16px", "text-align": "center", "margin": "0px", "--hover-color": "#d4edda"},
# #             "nav-link-selected": {"background-color": "#28a745", "color": "white"},
# #         }
# #     )
# #     st.markdown("<div class='card'>", unsafe_allow_html=True)
# #
# #     if selected_tab == "Upload DAR for MCM":
# #         st.markdown("<h3>Upload DAR PDF for MCM Period</h3>", unsafe_allow_html=True)
# #
# #         if not active_periods:
# #             st.warning("No active MCM periods available for upload. Please contact the Planning Officer.")
# #         else:
# #             period_options = {f"{p_data['month_name']} {p_data['year']}": p_key for p_key, p_data in
# #                               sorted(active_periods.items(), reverse=True)}
# #             selected_period_display = st.selectbox("Select Active MCM Period", options=list(period_options.keys()),
# #                                                    key="ag_select_mcm_period_upload")
# #
# #             if selected_period_display:
# #                 selected_mcm_key = period_options[selected_period_display]
# #                 mcm_info = mcm_periods[selected_mcm_key]
# #
# #                 if st.session_state.get('ag_current_mcm_key') != selected_mcm_key:
# #                     st.session_state.ag_current_extracted_data = []
# #                     st.session_state.ag_pdf_drive_url = None
# #                     st.session_state.ag_validation_errors = []
# #                     st.session_state.ag_editor_data = pd.DataFrame()
# #                     st.session_state.ag_current_mcm_key = selected_mcm_key
# #                     st.session_state.ag_current_uploaded_file_name = None
# #
# #                 st.info(f"Preparing to upload for: {mcm_info['month_name']} {mcm_info['year']}")
# #                 uploaded_dar_file = st.file_uploader("Choose a DAR PDF file", type="pdf",
# #                                                      key=f"dar_upload_ag_{selected_mcm_key}_{st.session_state.get('uploader_key_suffix', 0)}")
# #
# #                 if uploaded_dar_file is not None:
# #                     if st.session_state.get('ag_current_uploaded_file_name') != uploaded_dar_file.name:
# #                         st.session_state.ag_current_extracted_data = []
# #                         st.session_state.ag_pdf_drive_url = None
# #                         st.session_state.ag_validation_errors = []
# #                         st.session_state.ag_editor_data = pd.DataFrame()
# #                         st.session_state.ag_current_uploaded_file_name = uploaded_dar_file.name
# #
# #                     if st.button("Extract Data from PDF", key=f"extract_btn_ag_{selected_mcm_key}",
# #                                  use_container_width=True):
# #                         st.session_state.ag_validation_errors = []
# #
# #                         with st.spinner("Processing PDF and extracting data with AI..."):
# #                             dar_pdf_bytes = uploaded_dar_file.getvalue()
# #                             dar_filename_on_drive = f"AG{st.session_state.audit_group_no}_{uploaded_dar_file.name}"
# #
# #                             st.session_state.ag_pdf_drive_url = None
# #                             pdf_drive_id, pdf_drive_url_temp = upload_to_drive(drive_service, dar_pdf_bytes,
# #                                                                                mcm_info['drive_folder_id'],
# #                                                                                dar_filename_on_drive)
# #
# #                             if not pdf_drive_id:
# #                                 st.error("Failed to upload PDF to Google Drive. Cannot proceed.")
# #                                 st.session_state.ag_editor_data = pd.DataFrame([{
# #                                     "audit_group_number": st.session_state.audit_group_no, "gstin": None,
# #                                     "trade_name": None, "category": None,
# #                                     "total_amount_detected_overall_rs": None, "total_amount_recovered_overall_rs": None,
# #                                     "audit_para_number": None, "audit_para_heading": "Manual Entry - PDF Upload Failed",
# #                                     "revenue_involved_lakhs_rs": None, "revenue_recovered_lakhs_rs": None
# #                                 }])
# #                             else:
# #                                 st.session_state.ag_pdf_drive_url = pdf_drive_url_temp
# #                                 st.success(
# #                                     f"DAR PDF uploaded to Google Drive: [Link]({st.session_state.ag_pdf_drive_url})")
# #
# #                                 preprocessed_text = preprocess_pdf_text(BytesIO(dar_pdf_bytes))
# #                                 if preprocessed_text.startswith("Error processing PDF with pdfplumber:") or \
# #                                         preprocessed_text.startswith("Error in preprocess_pdf_text_"):
# #                                     st.error(f"PDF Preprocessing Error: {preprocessed_text}")
# #                                     st.warning("AI extraction cannot proceed. Please fill data manually.")
# #                                     default_row = {
# #                                         "audit_group_number": st.session_state.audit_group_no, "gstin": None,
# #                                         "trade_name": None, "category": None,
# #                                         "total_amount_detected_overall_rs": None,
# #                                         "total_amount_recovered_overall_rs": None,
# #                                         "audit_para_number": None, "audit_para_heading": "Manual Entry - PDF Error",
# #                                         "revenue_involved_lakhs_rs": None, "revenue_recovered_lakhs_rs": None
# #                                     }
# #                                     st.session_state.ag_editor_data = pd.DataFrame([default_row])
# #                                 else:
# #                                     parsed_report_obj = get_structured_data_with_gemini(YOUR_GEMINI_API_KEY,
# #                                                                                         preprocessed_text)
# #
# #                                     temp_extracted_list = []
# #                                     ai_extraction_failed_completely = True
# #
# #                                     if parsed_report_obj.parsing_errors:
# #                                         st.warning(f"AI Parsing Issues: {parsed_report_obj.parsing_errors}")
# #
# #                                     if parsed_report_obj and parsed_report_obj.header:
# #                                         header = parsed_report_obj.header
# #                                         ai_extraction_failed_completely = False
# #                                         if parsed_report_obj.audit_paras:
# #                                             for para_idx, para in enumerate(parsed_report_obj.audit_paras):
# #                                                 flat_item_data = {
# #                                                     "audit_group_number": st.session_state.audit_group_no,
# #                                                     "gstin": header.gstin, "trade_name": header.trade_name,
# #                                                     "category": header.category,
# #                                                     "total_amount_detected_overall_rs": header.total_amount_detected_overall_rs,
# #                                                     "total_amount_recovered_overall_rs": header.total_amount_recovered_overall_rs,
# #                                                     "audit_para_number": para.audit_para_number,
# #                                                     "audit_para_heading": para.audit_para_heading,
# #                                                     "revenue_involved_lakhs_rs": para.revenue_involved_lakhs_rs,
# #                                                     "revenue_recovered_lakhs_rs": para.revenue_recovered_lakhs_rs
# #                                                 }
# #                                                 temp_extracted_list.append(flat_item_data)
# #                                         elif header.trade_name:
# #                                             temp_extracted_list.append({
# #                                                 "audit_group_number": st.session_state.audit_group_no,
# #                                                 "gstin": header.gstin, "trade_name": header.trade_name,
# #                                                 "category": header.category,
# #                                                 "total_amount_detected_overall_rs": header.total_amount_detected_overall_rs,
# #                                                 "total_amount_recovered_overall_rs": header.total_amount_recovered_overall_rs,
# #                                                 "audit_para_number": None,
# #                                                 "audit_para_heading": "N/A - Header Info Only (Add Paras Manually)",
# #                                                 "revenue_involved_lakhs_rs": None, "revenue_recovered_lakhs_rs": None
# #                                             })
# #                                         else:
# #                                             st.error("AI failed to extract key header information (like Trade Name).")
# #                                             ai_extraction_failed_completely = True
# #
# #                                     if ai_extraction_failed_completely or not temp_extracted_list:
# #                                         st.warning(
# #                                             "AI extraction failed or yielded no data. Please fill details manually in the table below.")
# #                                         default_row = {
# #                                             "audit_group_number": st.session_state.audit_group_no, "gstin": None,
# #                                             "trade_name": None, "category": None,
# #                                             "total_amount_detected_overall_rs": None,
# #                                             "total_amount_recovered_overall_rs": None,
# #                                             "audit_para_number": None, "audit_para_heading": "Manual Entry Required",
# #                                             "revenue_involved_lakhs_rs": None, "revenue_recovered_lakhs_rs": None
# #                                         }
# #                                         st.session_state.ag_editor_data = pd.DataFrame([default_row])
# #                                     else:
# #                                         st.session_state.ag_editor_data = pd.DataFrame(temp_extracted_list)
# #                                         st.info(
# #                                             "Data extracted. Please review and edit in the table below before submitting.")
# #
# #                 if not isinstance(st.session_state.get('ag_editor_data'), pd.DataFrame):
# #                     st.session_state.ag_editor_data = pd.DataFrame()
# #
# #                 if uploaded_dar_file is not None and st.session_state.ag_editor_data.empty and \
# #                         st.session_state.get('ag_pdf_drive_url'):
# #                     st.warning(
# #                         "AI could not extract data, or no data was previously loaded. A template row is provided for manual entry.")
# #                     default_row = {
# #                         "audit_group_number": st.session_state.audit_group_no, "gstin": None, "trade_name": None,
# #                         "category": None,
# #                         "total_amount_detected_overall_rs": None, "total_amount_recovered_overall_rs": None,
# #                         "audit_para_number": None, "audit_para_heading": "Manual Entry",
# #                         "revenue_involved_lakhs_rs": None, "revenue_recovered_lakhs_rs": None
# #                     }
# #                     st.session_state.ag_editor_data = pd.DataFrame([default_row])
# #
# #                 if not st.session_state.ag_editor_data.empty:
# #                     st.markdown("<h4>Review and Edit Extracted Data:</h4>", unsafe_allow_html=True)
# #                     st.caption("Note: 'Audit Group Number' is pre-filled from your login. Add new rows for more paras.")
# #
# #                     ag_column_order = [
# #                         "audit_group_number", "gstin", "trade_name", "category",
# #                         "total_amount_detected_overall_rs", "total_amount_recovered_overall_rs",
# #                         "audit_para_number", "audit_para_heading",
# #                         "revenue_involved_lakhs_rs", "revenue_recovered_lakhs_rs"
# #                     ]
# #
# #                     df_to_edit_ag = st.session_state.ag_editor_data.copy()
# #                     df_to_edit_ag["audit_group_number"] = st.session_state.audit_group_no
# #
# #                     for col in ag_column_order:
# #                         if col not in df_to_edit_ag.columns:
# #                             df_to_edit_ag[col] = None
# #
# #                     ag_column_config = {
# #                         "audit_group_number": st.column_config.NumberColumn("Audit Group", format="%d", width="small",
# #                                                                             disabled=True),
# #                         "gstin": st.column_config.TextColumn("GSTIN", width="medium", help="Enter 15-digit GSTIN"),
# #                         "trade_name": st.column_config.TextColumn("Trade Name", width="large"),
# #                         "category": st.column_config.SelectboxColumn("Category", options=VALID_CATEGORIES,
# #                                                                      width="medium",
# #                                                                      help=f"Select from {VALID_CATEGORIES}"),
# #                         "total_amount_detected_overall_rs": st.column_config.NumberColumn("Total Detected (Rs)",
# #                                                                                           format="%.2f",
# #                                                                                           width="medium"),
# #                         "total_amount_recovered_overall_rs": st.column_config.NumberColumn("Total Recovered (Rs)",
# #                                                                                            format="%.2f",
# #                                                                                            width="medium"),
# #                         "audit_para_number": st.column_config.NumberColumn("Para No.", format="%d", width="small",
# #                                                                            help="Enter para number as integer"),
# #                         "audit_para_heading": st.column_config.TextColumn("Para Heading", width="xlarge"),
# #                         "revenue_involved_lakhs_rs": st.column_config.NumberColumn("Rev. Involved (Lakhs)",
# #                                                                                    format="%.2f", width="medium"),
# #                         "revenue_recovered_lakhs_rs": st.column_config.NumberColumn("Rev. Recovered (Lakhs)",
# #                                                                                     format="%.2f", width="medium"),
# #                     }
# #
# #                     editor_instance_key = f"ag_data_editor_{selected_mcm_key}_{st.session_state.ag_current_uploaded_file_name or 'no_file'}"
# #
# #                     edited_ag_df_output = st.data_editor(
# #                         df_to_edit_ag.reindex(columns=ag_column_order, fill_value=None),
# #                         column_config=ag_column_config,
# #                         num_rows="dynamic",
# #                         key=editor_instance_key,
# #                         use_container_width=True,
# #                         height=400
# #                     )
# #
# #                     if st.button("Validate and Submit to MCM Sheet", key=f"submit_ag_btn_{selected_mcm_key}",
# #                                  use_container_width=True):
# #                         current_data_for_validation = pd.DataFrame(edited_ag_df_output)
# #                         current_data_for_validation["audit_group_number"] = st.session_state.audit_group_no
# #
# #                         validation_errors = validate_data_for_sheet(current_data_for_validation)
# #                         st.session_state.ag_validation_errors = validation_errors
# #
# #                         if not validation_errors:
# #                             if not st.session_state.ag_pdf_drive_url:
# #                                 st.error(
# #                                     "PDF Drive URL is missing. This usually means the PDF was not uploaded to Drive. Please click 'Extract Data from PDF' again.")
# #                             else:
# #                                 with st.spinner("Submitting data to Google Sheet..."):
# #                                     rows_to_append_final = []
# #                                     record_created_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
# #                                     for index, row_to_save in current_data_for_validation.iterrows():
# #                                         sheet_row = [
# #                                             row_to_save.get("audit_group_number"),
# #                                             row_to_save.get("gstin"), row_to_save.get("trade_name"),
# #                                             row_to_save.get("category"),
# #                                             row_to_save.get("total_amount_detected_overall_rs"),
# #                                             row_to_save.get("total_amount_recovered_overall_rs"),
# #                                             row_to_save.get("audit_para_number"), row_to_save.get("audit_para_heading"),
# #                                             row_to_save.get("revenue_involved_lakhs_rs"),
# #                                             row_to_save.get("revenue_recovered_lakhs_rs"),
# #                                             st.session_state.ag_pdf_drive_url,
# #                                             record_created_date
# #                                         ]
# #                                         rows_to_append_final.append(sheet_row)
# #
# #                                     if rows_to_append_final:
# #                                         sheet_id = mcm_info['spreadsheet_id']
# #                                         append_result = append_to_spreadsheet(sheets_service, sheet_id,
# #                                                                               rows_to_append_final)
# #                                         if append_result:
# #                                             st.success(
# #                                                 f"Data for '{st.session_state.ag_current_uploaded_file_name}' successfully submitted to MCM sheet for {mcm_info['month_name']} {mcm_info['year']}!")
# #                                             st.balloons()
# #                                             time.sleep(0.5)
# #                                             st.session_state.ag_current_extracted_data = []
# #                                             st.session_state.ag_pdf_drive_url = None
# #                                             st.session_state.ag_editor_data = pd.DataFrame()
# #                                             st.session_state.ag_current_uploaded_file_name = None
# #                                             st.session_state.uploader_key_suffix = st.session_state.get(
# #                                                 'uploader_key_suffix', 0) + 1
# #                                             st.rerun()
# #                                         else:
# #                                             st.error("Failed to append data to Google Sheet after validation.")
# #                                     else:
# #                                         st.error("No data to submit after validation, though validation passed.")
# #                         else:
# #                             st.error(
# #                                 "Validation Failed! Please correct the errors displayed below and try submitting again.")
# #
# #                 if st.session_state.get('ag_validation_errors'):
# #                     st.markdown("---")
# #                     st.subheader("⚠️ Validation Errors - Please Fix Before Submitting:")
# #                     for err in st.session_state.ag_validation_errors:
# #                         st.warning(err)
# #
# #
# #     elif selected_tab == "View My Uploaded DARs":
# #         st.markdown("<h3>My Uploaded DARs</h3>", unsafe_allow_html=True)
# #         if not active_periods and not mcm_periods:
# #             st.info("No MCM periods have been created by the Planning Officer yet.")
# #         else:
# #             all_period_options = {f"{p_data['month_name']} {p_data['year']}": p_key for p_key, p_data in
# #                                   sorted(mcm_periods.items(), key=lambda item: item[0], reverse=True)}
# #             if not all_period_options:
# #                 st.info("No MCM periods found.")
# #             else:
# #                 selected_view_period_display = st.selectbox("Select MCM Period to View Your Uploads",
# #                                                             options=list(all_period_options.keys()),
# #                                                             key="ag_view_my_dars_period")
# #
# #                 if selected_view_period_display and sheets_service:
# #                     selected_view_period_key = all_period_options[selected_view_period_display]
# #                     view_mcm_info = mcm_periods[selected_view_period_key]
# #                     sheet_id_to_view = view_mcm_info['spreadsheet_id']
# #
# #                     st.markdown(f"**Fetching your uploads for {selected_view_period_display}...**")
# #                     with st.spinner("Loading data from Google Sheet..."):
# #                         df_all_uploads = read_from_spreadsheet(sheets_service, sheet_id_to_view)
# #
# #                     if not df_all_uploads.empty:
# #                         if 'Audit Group Number' in df_all_uploads.columns:
# #                             df_all_uploads['Audit Group Number'] = df_all_uploads['Audit Group Number'].astype(str)
# #                             my_group_uploads_df = df_all_uploads[
# #                                 df_all_uploads['Audit Group Number'] == str(st.session_state.audit_group_no)]
# #
# #                             if not my_group_uploads_df.empty:
# #                                 st.markdown(f"<h4>Your Uploads for {selected_view_period_display}:</h4>",
# #                                             unsafe_allow_html=True)
# #                                 df_display_my_uploads = my_group_uploads_df.copy()
# #                                 if 'DAR PDF URL' in df_display_my_uploads.columns:
# #                                     df_display_my_uploads['DAR PDF URL'] = df_display_my_uploads['DAR PDF URL'].apply(
# #                                         lambda x: f'<a href="{x}" target="_blank">View PDF</a>' if pd.notna(
# #                                             x) and isinstance(x, str) and x.startswith("http") else "No Link"
# #                                     )
# #                                 view_cols = ["Trade Name", "Category", "Audit Para Number", "Audit Para Heading",
# #                                              "DAR PDF URL", "Record Created Date"]
# #                                 for col_v in view_cols:
# #                                     if col_v not in df_display_my_uploads.columns:
# #                                         df_display_my_uploads[col_v] = "N/A"
# #
# #                                 st.markdown(df_display_my_uploads[view_cols].to_html(escape=False, index=False),
# #                                             unsafe_allow_html=True)
# #                             else:
# #                                 st.info(f"You have not uploaded any DARs for {selected_view_period_display} yet.")
# #                         else:
# #                             st.warning(
# #                                 "The MCM spreadsheet is missing the 'Audit Group Number' column. Cannot filter your uploads.")
# #                     else:
# #                         st.info(f"No data found in the MCM spreadsheet for {selected_view_period_display}.")
# #                 elif not sheets_service:
# #                     st.error("Google Sheets service not available.")
# #
# #     elif selected_tab == "Delete My DAR Entries":  # MODIFIED TAB NAME
# #         st.markdown("<h3>Delete My Uploaded DAR Entries</h3>", unsafe_allow_html=True)
# #         st.info(
# #             "Select an MCM period to view your entries. You can then delete specific entries after password confirmation. This action removes the entry from the Google Sheet only; the PDF in Google Drive will NOT be deleted.")
# #
# #         if not mcm_periods:
# #             st.info("No MCM periods have been created yet.")
# #         else:
# #             all_period_options_del = {f"{p_data['month_name']} {p_data['year']}": p_key for p_key, p_data in
# #                                       sorted(mcm_periods.items(), key=lambda item: item[0], reverse=True)}
# #             selected_del_period_display = st.selectbox("Select MCM Period to Manage Entries",
# #                                                        options=list(all_period_options_del.keys()),
# #                                                        key="ag_delete_my_dars_period_select")
# #
# #             if selected_del_period_display and sheets_service:
# #                 selected_del_period_key = all_period_options_del[selected_del_period_display]
# #                 del_mcm_info = mcm_periods[selected_del_period_key]
# #                 sheet_id_to_manage = del_mcm_info['spreadsheet_id']
# #
# #                 # Fetch sheet GID (needed for row deletion)
# #                 try:
# #                     sheet_metadata_for_gid = sheets_service.spreadsheets().get(
# #                         spreadsheetId=sheet_id_to_manage).execute()
# #                     first_sheet_gid = sheet_metadata_for_gid.get('sheets', [{}])[0].get('properties', {}).get('sheetId',
# #                                                                                                               0)
# #                 except Exception as e_gid:
# #                     st.error(f"Could not fetch sheet GID for deletion: {e_gid}")
# #                     first_sheet_gid = 0  # Fallback, might not work if not default sheet
# #
# #                 st.markdown(f"**Fetching your uploads for {selected_del_period_display}...**")
# #                 with st.spinner("Loading data..."):
# #                     df_all_uploads_del = read_from_spreadsheet(sheets_service, sheet_id_to_manage)
# #
# #                 if not df_all_uploads_del.empty and 'Audit Group Number' in df_all_uploads_del.columns:
# #                     df_all_uploads_del['Audit Group Number'] = df_all_uploads_del['Audit Group Number'].astype(str)
# #                     # Add a temporary unique identifier for selection if needed, or use index
# #                     my_group_uploads_df_del = df_all_uploads_del[
# #                         df_all_uploads_del['Audit Group Number'] == str(st.session_state.audit_group_no)].reset_index(
# #                         drop=True)  # Original index from sheet is lost here
# #                     my_group_uploads_df_del[
# #                         'original_sheet_row_index_approx'] = my_group_uploads_df_del.index + 2  # Approximate, +1 for header, +1 for 0-based to 1-based
# #
# #                     if not my_group_uploads_df_del.empty:
# #                         st.markdown(
# #                             f"<h4>Your Uploads in {selected_del_period_display} (Select an entry to delete):</h4>",
# #                             unsafe_allow_html=True)
# #
# #                         # Create a list of options for the selectbox
# #                         options_for_deletion = []
# #                         for idx, row_data in my_group_uploads_df_del.iterrows():
# #                             # Create a more robust identifier for the row
# #                             identifier = f"Row {idx + 1} (Sheet Approx. Row {row_data['original_sheet_row_index_approx']}) - TN: {row_data.get('Trade Name', 'N/A')}, Para: {row_data.get('Audit Para Number', 'N/A')}, Date: {row_data.get('Record Created Date', 'N/A')}"
# #                             options_for_deletion.append(identifier)
# #
# #                         selected_entry_to_delete_display = st.selectbox("Select Entry to Delete:",
# #                                                                         ["--Select an entry--"] + options_for_deletion,
# #                                                                         key=f"del_select_{selected_del_period_key}")
# #
# #                         if selected_entry_to_delete_display != "--Select an entry--":
# #                             # Find the actual row data and its approximate original index
# #                             selected_idx = options_for_deletion.index(selected_entry_to_delete_display)
# #                             row_to_delete_details = my_group_uploads_df_del.iloc[selected_idx]
# #
# #                             st.warning(
# #                                 f"You selected to delete: **{row_to_delete_details.get('Trade Name', '')} - Para {row_to_delete_details.get('Audit Para Number', '')}** (Uploaded: {row_to_delete_details.get('Record Created Date', '')})")
# #
# #                             with st.form(key=f"delete_ag_entry_form_{selected_idx}"):
# #                                 st.write("Please confirm your password to delete this entry from the Google Sheet.")
# #                                 ag_password_confirm = st.text_input("Your Password:", type="password",
# #                                                                     key=f"ag_pass_confirm_del_{selected_idx}")
# #                                 submitted_ag_delete = st.form_submit_button("Confirm Deletion of This Entry")
# #
# #                                 if submitted_ag_delete:
# #                                     if ag_password_confirm == USER_CREDENTIALS.get(st.session_state.username):
# #                                         # Re-fetch the sheet to get current row indices accurately before deletion
# #                                         df_before_delete = read_from_spreadsheet(sheets_service, sheet_id_to_manage)
# #                                         if not df_before_delete.empty and 'Audit Group Number' in df_before_delete.columns:
# #                                             df_before_delete['Audit Group Number'] = df_before_delete[
# #                                                 'Audit Group Number'].astype(str)
# #
# #                                             # Identify the actual 0-based index in the current sheet data
# #                                             # This matching needs to be robust. Using multiple fields.
# #                                             indices_to_delete_from_sheet = []
# #                                             for sheet_idx, sheet_row in df_before_delete.iterrows():
# #                                                 match = True
# #                                                 if str(sheet_row.get('Audit Group Number')) != str(
# #                                                     st.session_state.audit_group_no): match = False
# #                                                 if str(sheet_row.get('Trade Name')) != str(
# #                                                     row_to_delete_details.get('Trade Name')): match = False
# #                                                 if str(sheet_row.get('Audit Para Number')) != str(
# #                                                     row_to_delete_details.get('Audit Para Number')): match = False
# #                                                 if str(sheet_row.get('Record Created Date')) != str(
# #                                                     row_to_delete_details.get('Record Created Date')): match = False
# #
# #                                                 if match:
# #                                                     indices_to_delete_from_sheet.append(
# #                                                         sheet_idx + 1)  # +1 because read_from_spreadsheet skips header
# #
# #                                             if indices_to_delete_from_sheet:
# #                                                 if delete_spreadsheet_rows(sheets_service, sheet_id_to_manage,
# #                                                                            first_sheet_gid,
# #                                                                            indices_to_delete_from_sheet):
# #                                                     st.success(
# #                                                         f"Entry for '{row_to_delete_details.get('Trade Name')}' - Para '{row_to_delete_details.get('Audit Para Number')}' deleted successfully from the sheet.")
# #                                                     time.sleep(0.5)
# #                                                     st.rerun()
# #                                                 else:
# #                                                     st.error(
# #                                                         "Failed to delete the entry from the sheet. See console/app logs for details.")
# #                                             else:
# #                                                 st.error(
# #                                                     "Could not find the exact entry in the sheet to delete. It might have been already modified or deleted.")
# #                                         else:
# #                                             st.error("Could not re-fetch sheet data for deletion.")
# #                                     else:
# #                                         st.error("Incorrect password. Deletion cancelled.")
# #                     else:
# #                         st.info(f"You have no uploads in {selected_del_period_display} to request deletion for.")
# #                 elif df_all_uploads_del.empty:
# #                     st.info(f"No data found in the MCM spreadsheet for {selected_del_period_display}.")
# #                 else:
# #                     st.warning(
# #                         "Spreadsheet is missing 'Audit Group Number' column. Cannot identify your uploads for deletion.")
# #             elif not sheets_service:
# #                 st.error("Google Sheets service not available.")
# #
# #     st.markdown("</div>", unsafe_allow_html=True)
# #
# #
# # # --- Main App Logic ---
# # if not st.session_state.logged_in:
# #     login_page()
# # else:
# #     if 'drive_service' not in st.session_state or 'sheets_service' not in st.session_state or \
# #             st.session_state.drive_service is None or st.session_state.sheets_service is None:
# #         with st.spinner(
# #                 "Initializing Google Services... Please ensure your 'credentials.json' (Service Account Key) is in the app directory."):
# #             st.session_state.drive_service, st.session_state.sheets_service = get_google_services()
# #             if st.session_state.drive_service and st.session_state.sheets_service:
# #                 st.success("Google Services Initialized.")
# #                 st.rerun()
# #             else:
# #                 pass
# #
# #     if st.session_state.drive_service and st.session_state.sheets_service:
# #         if st.session_state.role == "PCO":
# #             pco_dashboard(st.session_state.drive_service, st.session_state.sheets_service)
# #         elif st.session_state.role == "AuditGroup":
# #             audit_group_dashboard(st.session_state.drive_service, st.session_state.sheets_service)
# #     elif st.session_state.logged_in:
# #         st.warning(
# #             "Google services are not yet available or failed to initialize. Please ensure 'credentials.json' (Service Account Key) is correctly placed and configured.")
# #         if st.button("Logout", key="main_logout_gerror_sa"):
# #             st.session_state.logged_in = False;
# #             st.rerun()
