import asyncio
import pathlib
import yaml
import os
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def load_config():
    config_path = f"{pathlib.Path(__file__).parent.absolute()}/config.yml"
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Configuration file not found: {config_path}")    
    with open(config_path, "r") as file:
        data = yaml.safe_load(file)
    return data

config = load_config()

def get_credentials():
    creds_path = os.path.join(
        pathlib.Path(__file__).parent.absolute(),
        config["Google"]["GOOGLE_SERVICE_ACCOUNT_FILE"]
    )
    credentials = service_account.Credentials.from_service_account_file(
        creds_path, 
        scopes=[
            "https://www.googleapis.com/auth/forms.responses.readonly",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.readonly"
        ]
    )
    return credentials

def flatten_response(response):
    row = {}
    row["responseId"] = "..." + response.get("responseId", "")[-3:]
    row["createTime"] = response.get("createTime", "")
    
    answers = response.get("answers", {})
    for question_id, answer in answers.items():
        text = ""
        if "textAnswers" in answer:
            text_answers = answer["textAnswers"].get("answers", [])
            if text_answers:
                values = [ans.get("value", "") for ans in text_answers]
                text = ", ".join(values)
        else:
            text = str(answer)

        row[question_id] = text
    
    return row

def export_using_forms_api(form_id, credentials):
    try:
        forms_service = build("forms", "v1", credentials=credentials)
        response = forms_service.forms().responses().list(formId=form_id).execute()
        responses = response.get("responses", [])
    
        if not responses:
            print("No responses found via Forms API.")
            return None
    
        rows = [flatten_response(resp) for resp in responses]
    
        df = pd.DataFrame(rows).fillna("")
    
        return df
    except HttpError as err:
        print(f"Forms API error: {err}")
        return None

def export_using_sheet_api(linked_sheet_id, credentials):
    try:
        sheets_service = build("sheets", "v4", credentials=credentials)
        metadata = sheets_service.spreadsheets().get(spreadsheetId=linked_sheet_id).execute()
        first_sheet = metadata['sheets'][0]['properties']['title']

        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=linked_sheet_id,
            range=first_sheet
        ).execute()
    
        values = result.get('values', [])
    
        if not values:
            print("No data found in the linked sheet.")
            return None
    
        df = pd.DataFrame(values)
        df = df.reset_index(drop=True)
    
        if not df.empty:
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)
        
        return df
    except HttpError as err:
        print(f"Sheets API error: {err}")
        return None

def get_linked_sheet_id(form_id, credentials):
    try:
        forms_service = build("forms", "v1", credentials=credentials)
        form_metadata = forms_service.forms().get(formId=form_id).execute()
        destination = form_metadata.get("responseDestination", {})
        if destination.get("destinationType") == "SPREADSHEET":
            linked_sheet = destination.get("spreadsheet")
            if linked_sheet:
                print(f"Found linked sheet ID: {linked_sheet}")
                return linked_sheet
        print("No linked sheet found in form metadata.")
        return None
    except HttpError as err:
        print(f"Error retrieving form metadata: {err}")
        return None

def export_to_fixed_width_txt(df, file_path):
    df_str = df.astype(str)
    
    col_widths = {
        col: max(df_str[col].str.len().max(), len(col))
        for col in df_str.columns
    }
    
    with open(file_path, 'w', encoding='utf-8') as f:
        header = "".join(
            col.ljust(col_widths[col] + 2) for col in df_str.columns
        )
        f.write(header.rstrip() + '\n')
        
        for _, row in df_str.iterrows():
            line = "".join(
                row[col].ljust(col_widths[col] + 2) for col in df_str.columns
            )
            f.write(line.rstrip() + '\n')

def export_to_spreadsheet(df, spreadsheet_id, sheet_name, credentials):
    try:
        sheets_service = build("sheets", "v4", credentials=credentials)
        
        data_to_write = [df.columns.values.tolist()] + df.values.tolist()
        
        range_to_clear = f"'{sheet_name}'!A1:Z"
        
        print(f"Clearing sheet '{sheet_name}' in spreadsheet {spreadsheet_id}...")
        sheets_service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=range_to_clear,
            body={}
        ).execute()

        print(f"Writing {len(df)} rows to sheet '{sheet_name}'...")
        body = {"values": data_to_write}
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_name}'!A1",
            valueInputOption="USER_ENTERED",
            body=body
        ).execute()
        
    except HttpError as err:
        print(f"Google Sheets API error: {err}")

async def export_form(form_config):
    form_id = form_config["GOOGLE_FORM_ID"]
    export_format = form_config.get("ExportFormat", "csv")
    print(f"Exporting form {form_id}...")
    credentials = get_credentials()

    df = export_using_forms_api(form_id, credentials)

    if df is None:
        linked_sheet_id = get_linked_sheet_id(form_id, credentials)
        if linked_sheet_id:
            df = export_using_sheet_api(linked_sheet_id, credentials)
        else:
            print("Unable to retrieve responses by any method.")
            return

    if df is None or df.empty:
        print("No data to export.")
        return
    
    if "createTime" in df.columns:
        df.sort_values(by="createTime", ascending=True, inplace=True)
    
    mapping_overrides = config.get("MappingOverrides", {})
    if mapping_overrides:
        df.rename(columns=mapping_overrides, inplace=True)

    print("Response columns:", df.columns.tolist())

    if export_format == "spreadsheet":
        spreadsheet_id = form_config.get("GOOGLE_SPREADSHEET_ID")
        sheet_name = form_config.get("SHEET_NAME", "Sheet1")
        if not spreadsheet_id:
            print(f"Error: GOOGLE_SPREADSHEET_ID not specified for form {form_id}.")
            return
        export_to_spreadsheet(df, spreadsheet_id, sheet_name, credentials)
        print(f"Successfully exported form responses to Google Sheet ID: {spreadsheet_id}")
    else:
        file_name = form_config.get("FILE_NAME")
        if not file_name:
            print(f"Error: FILE_NAME not specified for form {form_id}.")
            return
        
        file_path = f"{pathlib.Path(__file__).parent.absolute()}/{file_name}"

        if export_format == "csv":
            export_to_fixed_width_txt(df, file_path)
            print(f"Exported form responses to {file_path} as perfectly aligned text.")
        elif export_format == "xlsx":
            df.to_excel(file_path, index=False)
            print(f"Exported form responses to {file_path} as XLSX.")
        else:
            print(f"Error: Unsupported export format '{export_format}'.")

async def initial_export():
    for form_config in config["Forms"]:
        await export_form(form_config)

async def run_every_5_minutes():
    while True:
        try:
            print("Running 5-minute export...")
            for form_config in config["Forms"]:
                await export_form(form_config)
            print("Sleeping for 5 minutes...")
            await asyncio.sleep(300)
        except Exception as e:
            print(f"Error in scheduled task: {e}")

async def main():
    await initial_export()
    await run_every_5_minutes()

if __name__ == '__main__':
    asyncio.run(main())