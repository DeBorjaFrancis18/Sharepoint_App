from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import os
import fnmatch
from datetime import datetime
from openpyxl import load_workbook
import sys
import tkinter as tk
from tkinter import messagebox

def show_popup(title, message):
    """Display a popup message box"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    messagebox.showinfo(title, message)
    root.destroy()

def read_config_file(file_path):
    config_values = {}
    try:
        with open(file_path, 'r') as file:
            lines = file.readlines()
            for line in lines:
                parts = line.strip().split("=")
                if len(parts) >= 2:
                    key = parts[0].strip()
                    value = "=".join(parts[1:]).strip().strip('"')
                    config_values[key] = value
                else:
                    print(f"Skipping malformed line: {line.strip()}")
    except FileNotFoundError:
        error_msg = f"Config file '{file_path}' not found."
        print(error_msg)
        show_popup("Error", error_msg)
    except Exception as e:
        error_msg = f"Error reading config file: {str(e)}"
        print(error_msg)
        show_popup("Error", error_msg)
    return config_values

def get_sharepoint_context_using_app(config_values):
    sharepoint_url = config_values.get('DestinationSiteURL')
    client_credentials = ClientCredential(
        config_values.get('Client Id'), 
        config_values.get('Client Secret')
    )
    ctx = ClientContext(sharepoint_url).with_credentials(client_credentials)
    return ctx

def update_log_sheet(log_sheet, file_name, status):
    log_sheet.insert_rows(2)
    log_sheet['A2'] = file_name
    log_sheet['B2'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_sheet['C2'] = status

def is_file_large(file_path, max_size_mb=250):
    """Check if a file is larger than the specified size in MB."""
    file_size_mb = os.path.getsize(file_path) / (1024 * 1024)  # Convert bytes to MB
    return file_size_mb > max_size_mb

def upload_file_in_chunks(ctx, target_folder, file_path, file_name, chunk_size_mb=10):
    """Upload a file to SharePoint in chunks using modern API."""
    chunk_size = chunk_size_mb * 1024 * 1024  # Convert MB to bytes
    file_size = os.path.getsize(file_path)
    offset = 0  # Initialize offset to ensure it always has a value

    try:
        print(f"Starting chunked upload for '{file_name}' ({file_size / 1024 / 1024:.2f} MB total)")

        with open(file_path, 'rb') as file:
            # Start the upload session
            first_chunk = file.read(chunk_size)
            uploaded_file = target_folder.files.create_upload_session(file_name, len(first_chunk)).execute_query()
            upload_id = uploaded_file.upload_id

            offset += len(first_chunk)
            print(f"Uploaded first chunk: {offset / 1024 / 1024:.2f}MB")

            # Continue uploading chunks
            while offset < file_size:
                chunk = file.read(chunk_size)
                if not chunk:
                    break

                if offset + len(chunk) < file_size:
                    uploaded_file = target_folder.files.continue_upload(upload_id, offset, chunk).execute_query()
                else:
                    # Finalize the upload with the last chunk
                    uploaded_file = target_folder.files.finish_upload(upload_id, offset, chunk).execute_query()

                offset += len(chunk)
                print(f"Uploaded {offset / 1024 / 1024:.2f}MB of {file_size / 1024 / 1024:.2f}MB")

        print(f"Successfully uploaded '{file_name}'")
        return uploaded_file

    except Exception as e:
        print(f"Upload failed at {offset / 1024 / 1024:.2f}MB: {str(e)}")
        raise

def upload_files_with_wildcard(file_path=None):
    # Get the directory of the current script (upload.exe)
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    config_file_path = os.path.join(script_dir, "config.txt")
    
    # Verify config file exists before proceeding
    if not os.path.exists(config_file_path):
        error_msg = f"ERROR: Config file not found at {config_file_path}"
        print(error_msg)
        show_popup("Config File Error", error_msg)
        return

    config_values = read_config_file(config_file_path)
    
    if file_path:
        source_folder_path = os.path.dirname(file_path)
        wildcard_pattern = os.path.basename(file_path)
    else:
        source_folder_path = config_values.get('SourceFolderPath')
        wildcard_pattern = config_values.get('FileName')
    
    ctx = get_sharepoint_context_using_app(config_values)
    target_folder_url = config_values.get('DestinationFolderURL')
    target_folder = ctx.web.get_folder_by_server_relative_url(target_folder_url)

    log_file_path = config_values.get('LogFilePath')
    log_workbook = load_workbook(log_file_path) if log_file_path and os.path.exists(log_file_path) else None
    log_sheet = log_workbook.active if log_workbook else None

    print(f"Target folder URL: {target_folder_url}")

    success_count = 0
    failure_count = 0
    processed_files = []

    try:
        for file_name in os.listdir(source_folder_path):
            if fnmatch.fnmatch(file_name, wildcard_pattern):
                file_path_to_upload = os.path.join(source_folder_path, file_name)
                
                try:
                    print(f"\nProcessing file: {file_name}")
                    
                    if is_file_large(file_path_to_upload):
                        print("Large file detected, using chunked upload...")
                        upload_file_in_chunks(ctx, target_folder, file_path_to_upload, file_name)
                    else:
                        print("Small file, using standard upload...")
                        with open(file_path_to_upload, 'rb') as content_file:
                            file_content = content_file.read()
                            target_folder.upload_file(file_name, file_content).execute_query()
                    
                    processed_files.append(f"✓ {file_name}")
                    success_count += 1
                    if log_sheet:
                        update_log_sheet(log_sheet, file_name, 'Successful')
                        log_workbook.save(log_file_path)
                        
                except Exception as file_error:
                    error_msg = f"Failed to upload {file_name}: {str(file_error)}"
                    print(error_msg)
                    processed_files.append(f"✗ {file_name}")
                    failure_count += 1
                    if log_sheet:
                        update_log_sheet(log_sheet, file_name, 'Failed')
                        log_workbook.save(log_file_path)
        
        # Show summary
        summary_msg = f"Upload completed!\n\nSuccess: {success_count}\nFailed: {failure_count}"
        if processed_files:
            summary_msg += "\n\nFiles processed:\n" + "\n".join(processed_files)
        else:
            summary_msg = "No files matching the pattern were found to upload."
        
        show_popup("Upload Summary", summary_msg)

    except Exception as e:
        error_msg = f"Critical error: {str(e)}"
        print(error_msg)
        show_popup("Error", error_msg)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        upload_files_with_wildcard(sys.argv[1])
    else:
        upload_files_with_wildcard()
