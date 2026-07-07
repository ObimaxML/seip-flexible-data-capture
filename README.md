# SEIP Flexible Data Capture App

This is a simple Streamlit app that reads the attached SEIP Excel data collection template and creates a capture interface for each collection sheet.

## Login details

Default login:

```text
Username: seip_admin
Password: seip_admin123
```

Change the password before using the app with real data.

To change it, edit:

```text
.streamlit/secrets.toml
```

Update this line:

```toml
SEIP_ADMIN_PASSWORD = "your_new_password_here"
```

The username is fixed in `app.py` as:

```python
AUTH_USERNAME = "seip_admin"
```

## What it does

- Requires login before accessing the data capture screens.
- Detects all template sheets dynamically from `*_fields` sheets.
- Creates a separate form for each capture area, for example Job Seeker, Business, Training Provider, and Informal Business.
- Uses field definitions from the workbook to choose the right input type.
- Uses reference values from the workbook for dropdowns where available.
- Appends captured records back into an Excel workbook.
- Allows you to view, download, and delete captured rows.
- Can export the active sheet to CSV.

## How to run in VS Code

1. Unzip this folder.
2. Open the folder in VS Code.
3. Open the VS Code terminal.
4. Create and activate a virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

5. Install dependencies:

```powershell
pip install -r requirements.txt
```

6. Run the app:

```powershell
streamlit run app.py
```

The app will open in your browser.

## Template file

The default template is stored here:

```text
data/SEIP_Data_Collection_Template.xlsx
```

You can replace it with an updated template, or upload another compatible Excel template from the app sidebar.

## Data file

When you capture records, the app writes to:

```text
data/SEIP_Captured_Data.xlsx
```

The original template is copied to this file the first time you save a record.
