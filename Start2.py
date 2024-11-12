import PySimpleGUI as sg
import subprocess
import os
import winreg
import win32con
from fuzzywuzzy import fuzz
from packaging import version
import re
import operator
import shutil
import datetime
from dataclasses import dataclass
import json
from pathlib import Path
import ctypes, sys
import glob
import threading
import time

@dataclass
class RegistryData:
    key: str
    value: bytes  # Expecting binary data
    type: str

    def __post_init__(self):
        # Convert bytes to a hexadecimal string
        if isinstance(self.value, bytes):
            self.value = self.value.hex()


REG_TYPE_MAP = {
    winreg.REG_SZ: "REG_SZ",  #
    winreg.REG_EXPAND_SZ: "REG_EXPAND_SZ",  #
    winreg.REG_BINARY: "REG_BINARY",  #
    winreg.REG_DWORD: "REG_DWORD",
    winreg.REG_DWORD_LITTLE_ENDIAN: "REG_DWORD_LITTLE_ENDIAN",  #
    winreg.REG_DWORD_BIG_ENDIAN: "REG_DWORD_BIG_ENDIAN",
    winreg.REG_LINK: "REG_LINK",
    winreg.REG_MULTI_SZ: "REG_MULTI_SZ",  #
    winreg.REG_RESOURCE_LIST: "REG_RESOURCE_LIST",
    winreg.REG_FULL_RESOURCE_DESCRIPTOR: "REG_FULL_RESOURCE_DESCRIPTOR",
    winreg.REG_RESOURCE_REQUIREMENTS_LIST: "REG_RESOURCE_REQUIREMENTS_LIST",
    winreg.REG_QWORD: "REG_QWORD",
    winreg.REG_QWORD_LITTLE_ENDIAN: "REG_QWORD_LITTLE_ENDIAN",  #
}


class CustomEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, bytes):
            return obj.decode('utf-8')  # Convert bytes to UTF-8 string
        return super().default(obj)

table_changed = False
search_text_page1 = ""  # For Page 1
search_text_page2 = ""  # For Page 2
search_text_page3 = ""  # For Page 3
search_text_page4 = ""
folder_path = ""  # Define as Global Variable
file_path = ""
BLANK_BOX = '⬜'
CHECKED_BOX = '✅'
current_time = datetime.datetime.now().strftime("%Y-%m-%d")
scurrent_time = str(current_time)

def reset_page4_state(window):
    # Clear the table and reset the state on Page 4
    window['-TABLE_EDITOR-'].update(values=[])  # Clear table data
    window['Dropdown2'].update(value='MicroAOI')  # Reset dropdown to default
    window['-BROWSE2-'].update('Browse')  # Reset browse button text
    window['-BROWSE2-'].update(disabled=False)  # Enable browse button

    # Reset buttons and visibility
    window['-ADD_NEW-'].update(visible=False)
    window['-EDIT_GOLDEN_FILE-'].update(visible=False, disabled=True)
    window['-DELETE_FROM_GOLDEN_FILE-'].update(visible=False, disabled=True)
    window['-SAVE_GOLDEN_FILE-'].update(visible=False, disabled=True)

    # Reset metadata to indicate no unsaved changes
    window['-SAVE_GOLDEN_FILE-'].metadata = False
    global table_changed
    table_changed = False


def get_next_revision_number(backup_folder, machine_type):
    # Get all files in the backup folder
    files = os.listdir(backup_folder)

    # Filter files that match the naming convention for the golden file backups
    revision_numbers = []
    for file_name in files:
        if file_name.startswith(machine_type) and "_rev" in file_name:
            rev_number = file_name.split("_rev")[1].split("_")[0]
            if rev_number.isdigit():
                revision_numbers.append(int(rev_number))

    # Return the next revision number
    if revision_numbers:
        return max(revision_numbers) + 1
    return 1


def compare_golden_and_rev_json(golden_file_path, restore_file_path, action_temp_path):
    # Load data from both JSON files
    golden_data = load_registry_from_json2(golden_file_path)
    restore_data = load_registry_from_json2(restore_file_path)

    if not golden_data or not restore_data:
        raise FileNotFoundError("Failed to load one of the files for comparison.")

    added_deleted_entries = []
    edited_entries = []

    # Comparison logic
    for entry in restore_data:
        path, name, data, reg_type = entry
        matching_entry = next((e for e in golden_data if e[0] == path and e[1] == name), None)

        if matching_entry:
            previous_data = matching_entry[2]
            previous_type = matching_entry[3]

            # If the data or type differs, it's an edit
            if data != previous_data or reg_type != previous_type:
                edited_entries.append({
                    'Registry Key/Subkey Path': path,
                    'Registry Name': name,
                    'Previous Data': previous_data,
                    'Current Data': data,
                    'Previous Type': previous_type,
                    'Current Type': reg_type,
                    'Action': 'Edit'
                })
        else:
            # If no matching entry in the golden file, it's a new addition
            added_deleted_entries.append({
                'Registry Key/Subkey Path': path,
                'Registry Name': name,
                'Data': data,
                'Type': reg_type,
                'Action': 'Add'
            })

    for entry in golden_data:
        path, name, data, reg_type = entry
        matching_entry = next((e for e in restore_data if e[0] == path and e[1] == name), None)

        # If an entry exists in the golden file but not in the restore file, it was deleted
        if not matching_entry:
            added_deleted_entries.append({
                'Registry Key/Subkey Path': path,
                'Registry Name': name,
                'Data': data,
                'Type': reg_type,
                'Action': 'Delete'
            })

    # Write to action temp file
    action_temp_data = added_deleted_entries + edited_entries
    write_to_json(action_temp_data, action_temp_path)
    print(f"Action temp file created: {action_temp_path}")

    return added_deleted_entries, edited_entries


def compare_files_and_get_changes(golden_file_path, edit_file_path):
    # Load golden file data
    golden_data = load_registry_from_json2(golden_file_path)
    edit_data = load_registry_from_json2(edit_file_path)

    if not golden_data:
        print("Golden file data not found.")
        return [], []

    if not edit_data:
        print("Edit file data not found.")
        return [], []

    added_deleted_entries = []
    edited_entries = []

    # Iterate over the entries in the edit temp file and categorize them
    for entry in edit_data:
        action = entry.get('Action')
        if action == 'Add' or action == 'Delete':
            added_deleted_entries.append(entry)
        elif action == 'Edit':
            edited_entries.append(entry)

    print(f"Added/Deleted Data for Table: {added_deleted_entries}")
    print(f"Edited Data for Table: {edited_entries}")

    return added_deleted_entries, edited_entries


def makeWinSave(title):
    file_path = get_file_path()
    file_name = Path(file_path).stem
    selected_file_output = sg.Text(file_name, font=("Helvetica", 11, "bold"))

    edit_data=None

    # Load the temp files (edit.temp and golden_file.temp)
    golden_temp_path = os.path.join("data", f"{file_name}_golden_file.temp")
    edit_temp_path = os.path.join("data", f"{file_name}_edit.temp")

    # Get the changes
    added_deleted_entries, edited_entries = compare_files_and_get_changes(golden_temp_path, edit_temp_path)

    # Prepare data for internal data structure (editor_data)
    editor_data = [
        [entry['Registry Key/Subkey Path'], entry['Registry Name'], entry['Data'], entry['Type']]
        for entry in added_deleted_entries
    ]

    # Prepare data for display (displayed_editor_data)
    displayed_editor_data = [
        [entry['Registry Key/Subkey Path'], entry['Registry Name'], entry['Type'], entry['Data'], entry['Action']]
        for entry in added_deleted_entries
    ]

    # Prepare data for Edited table
    edited_data = [
        [entry['Registry Key/Subkey Path'], entry['Registry Name'], entry['Previous Data'], entry['Current Data'], entry['Previous Type'], entry['Current Type']]
        for entry in edited_entries
    ]

    displayed_editor_edited_data = [
        [ entry['Registry Key/Subkey Path'], entry['Registry Name'], entry['Previous Type'], entry['Current Type'], entry['Previous Data'], entry['Current Data']]
        for entry in edited_entries
    ]

    layoutSave = [
        [sg.Push(), sg.pin(sg.Text("Golden List Selected:", font=("Helvetica", 12, "bold"))), selected_file_output,
         sg.Push()],

        # Frame for Registry Data Added/Deleted
        [sg.Frame('Table for Registry Data Added/Deleted', font=("Helvetica", 12, "bold"), layout=[
            [
                sg.Table(
                    values=displayed_editor_data,  # Populate with displayed_editor_data
                    headings=["Registry Key/Subkey Path", "Registry Name", "Type", "Data", "Action Taken"],
                    auto_size_columns=False,
                    vertical_scroll_only=False,
                    justification="left",
                    num_rows=20,
                    key="-TABLE_REG_ADDED_DELETED-",
                    col_widths=[40, 20, 30, 15, 15],  # Adjust column widths
                    background_color="#045D5D",
                    text_color="white",
                    bind_return_key=True,
                    row_colors=[("white", "#045D5D"), ("white", "yellow"), ("white", "red")],
                    enable_events=True,
                    enable_click_events=True,
                    select_mode=sg.TABLE_SELECT_MODE_EXTENDED,
                    tooltip=None,
                    expand_x=True,
                    expand_y=True
                ),
            ],
        ], element_justification="left", size=(450, 200), expand_x=True, expand_y=True)],

        # Frame for Registry Data Edited
        [sg.Frame('Table for Registry Data Edited', font=("Helvetica", 12, "bold"), layout=[
            [
                sg.Table(
                    values=displayed_editor_edited_data,
                    headings=["Registry Key/Subkey Path", "Registry Name", "Previous Type", "Current Type",
                              "Previous Data", "Current Data"],
                    auto_size_columns=False,
                    vertical_scroll_only=False,
                    justification="left",
                    num_rows=10,
                    key="-TABLE_REG_EDITED-",
                    col_widths=[40, 20, 15, 15, 30, 30],  # Adjust column widths
                    background_color="#045D5D",
                    text_color="white",
                    bind_return_key=True,
                    row_colors=[("white", "#045D5D"), ("white", "yellow"), ("white", "red")],
                    enable_events=True,
                    enable_click_events=True,
                    select_mode=sg.TABLE_SELECT_MODE_EXTENDED,
                    tooltip=None,
                    expand_x=True,
                    expand_y=True
                ),
            ],
        ], element_justification="left", size=(450, 200), expand_x=True, expand_y=True)
         ],

        [sg.Push(), sg.Button("Save", key="-SAVE_REGISTRY_CHANGES-"),
         sg.Button("Discard", key="-DISCARD_REGISTRY_CHANGES-")]
    ]

    save_window = sg.Window(title, layoutSave, finalize=True, size=(1000, 750), resizable=True)

    # Maximize the window
    save_window.maximize()

    while True:
        event_save, values_save = save_window.read()

        if event_save == sg.WIN_CLOSED:
            save_window.close()
            break

        if event_save == "-SAVE_REGISTRY_CHANGES-":
            confirmation = sg.popup_yes_no("Do you want to save the changes made to the registry?", title="Confirmation")
            if confirmation == "Yes":
                # Get current timestamp with the provided format
                current_time = datetime.datetime.now().strftime("%Y-%m-%d")
                scurrent_time = str(current_time)

                # Step 1: Load the temp files
                golden_file_path = os.path.join("data", f"{file_name}_golden_file.temp")
                edit_temp_path = os.path.join("data", f"{file_name}_edit.temp")
                original_golden_file_path = f"Golden File/{file_name}.json"

                # Load the data from golden file, edit temp, and current PC registry file
                golden_data = load_registry_from_json2(golden_file_path)
                edit_data = load_registry_from_json2(edit_temp_path)

                # Step 2: Create a backup of the original golden file before applying changes
                machine_type = file_name  # Assuming machine_type is represented by file_name
                saved_rev_folder = os.path.join("Backup", "Backup(Golden File)", "Saved Rev", machine_type)

                # Check if the backup folder for the machine type exists, if not, create it
                if not os.path.exists(saved_rev_folder):
                    os.makedirs(saved_rev_folder)
                    print(f"Backup folder created for {machine_type}: {saved_rev_folder}")

                # Generate backup file name with version number and timestamp
                backup_file_name = f"{file_name}_rev{get_next_revision_number(saved_rev_folder, file_name)}_{scurrent_time}.json"
                backup_file_path = os.path.join(saved_rev_folder, backup_file_name)

                # Copy the original golden file to the backup location
                shutil.copy2(original_golden_file_path, backup_file_path)
                print(f"Backup created: {backup_file_path}")

                # Step 3: Apply changes from the edit temp file to both the golden data and current PC registry data
                for change in edit_data:
                    action = change.get("Action")
                    registry_path = change['Registry Key/Subkey Path']
                    registry_name = change['Registry Name']
                    registry_type = change.get('Type', change.get('Current Type', 'Unknown'))
                    registry_data = change.get('Data', change.get('Current Data', 'Unknown'))

                    if action == "Add":
                        # Update Golden Data
                        golden_data.append([registry_path, registry_name, registry_data, registry_type])
                        # Update Current PC Registry Data
                        write_into_event_log(
                            f"User confirmed added registry key: Path='{registry_path}', Name='{registry_name}', "
                            f"Type='{registry_type}', Data='{registry_data}'.")

                    elif action == "Delete":
                        # Update Golden Data
                        golden_data = [entry for entry in golden_data if not (
                                entry[0] == registry_path and entry[1] == registry_name)]
                        # Log the confirmed deletion
                        write_into_event_log(
                            f"User confirmed deleted registry key: Path='{registry_path}', Name='{registry_name}', "
                            f"Type='{registry_type}', Data='{registry_data}'.")

                    elif action == "Edit":
                        for entry in golden_data:
                            if entry[0] == registry_path and entry[1] == registry_name:
                                # Log the original data and the updated data
                                previous_data = entry[2]
                                previous_type = entry[3]

                                entry[3] = change['Current Type']
                                entry[2] = change['Current Data']

                                write_into_event_log(
                                    f"User confirmed edited registry key: Path='{registry_path}', Name='{registry_name}', "
                                    f"Previous Type='{previous_type}', New Type='{entry[3]}', "
                                    f"Previous Data='{previous_data}', New Data='{entry[2]}'.")

                # Step 4: Write the updated data back to the golden file and current PC registry file
                write_new_data_to_json(golden_data, original_golden_file_path)
                sg.popup_ok("Changes saved successfully.", title="Information")

                # Step 5: Remove temp files after applying changes
                if os.path.exists(golden_file_path):
                    os.remove(golden_file_path)
                if os.path.exists(edit_temp_path):
                    os.remove(edit_temp_path)

                # Step 6: Update the main window GUI (if necessary)
                update_editor_gui(original_golden_file_path)
                reset_page4_state(window)
                save_window.close()

                return "saved"

            else:
                sg.popup_ok("Save action cancelled.", title="Information")

        if event_save == "-DISCARD_REGISTRY_CHANGES-":
            confirmation = sg.popup_yes_no("Are you sure you want to discard the changes? This cannot be undone.", title="Confirmation")
            if confirmation == "Yes":
                # Step 1: Remove temp files without applying changes
                golden_file_path = os.path.join("data", f"{file_name}_golden_file.temp")
                edit_temp_path = os.path.join("data", f"{file_name}_edit.temp")

                if os.path.exists(golden_file_path):
                    os.remove(golden_file_path)
                if os.path.exists(edit_temp_path):
                    os.remove(edit_temp_path)

                sg.popup_ok("Changes discarded successfully.", title="Information")

                if 'edit_data' in locals() and edit_data:
                    for change in edit_data:
                        write_into_event_log(
                            f"Registry data {change['Registry Name']} at {change['Registry Key/Subkey Path']} was discarded.")

                else:
                    write_into_event_log("No changes to discard.")

                # Step 2: Clear the table and reset the state on Page 4
                window['-TABLE_EDITOR-'].update(values=[])  # Set the table to empty

                # Clear the machine type selection and browse path
                window['Dropdown2'].update(value='MicroAOI')  # Reset dropdown value to default
                window['-BROWSE2-'].update('Browse')  # Clear browse button path
                window['-BROWSE2-'].update(disabled=False)

                reset_page4_state(window)
                save_window.close()

                return "discarded"

            else:
                sg.popup_ok("Discard action cancelled.")

    return save_window


def makeWin4Restore(added_deleted_entries, edited_entries):
    file_path = get_file_path()
    file_name = Path(file_path).stem
    selected_file_output = sg.Text(file_name, font=("Helvetica", 11, "bold"))

    # Prepare data for Add/Delete table
    displayed_added_deleted_data = [
        [entry['Registry Key/Subkey Path'], entry['Registry Name'], entry['Type'], entry['Data'], entry['Action']]
        for entry in added_deleted_entries
    ]

    # Prepare data for Edited table
    displayed_edited_data = [
        [entry['Registry Key/Subkey Path'], entry['Registry Name'], entry['Previous Type'], entry['Current Type'],
         entry['Previous Data'], entry['Current Data']]
        for entry in edited_entries
    ]

    layoutRestore = [
        [sg.Push(), sg.Text(f"Restore Registry - {file_name}", font=("Helvetica", 14, "bold")), sg.Push()],
        [sg.Frame('Add/Delete Actions', font=("Helvetica", 12, "bold"), layout=[
            [
                sg.Table(
                    values=displayed_added_deleted_data,
                    headings=["Path", "Name", "Type", "Data", "Action"],
                    auto_size_columns=False,
                    justification="left",
                    num_rows=15,
                    key="-TABLE_REG_ADDED_DELETED-",
                    col_widths=[40, 20, 15, 30, 15],
                    background_color="#045D5D",
                    text_color="white",
                    bind_return_key=True,
                    row_colors=[("white", "#045D5D"), ("white", "yellow"), ("white", "red")],
                    enable_events=True,
                    enable_click_events=True,
                    select_mode=sg.TABLE_SELECT_MODE_EXTENDED,
                    tooltip=None,
                    expand_x=True,
                    expand_y=True
                ),
            ],
        ], expand_x=True, expand_y=True)],
        [sg.Frame('Edited Registry Data', font=("Helvetica", 12, "bold"), layout=[
            [
                sg.Table(
                    values=displayed_edited_data,
                    headings=["Path", "Name", "Previous Type", "Current Type", "Previous Data", "Current Data"],
                    auto_size_columns=False,
                    justification="left",
                    num_rows=10,
                    key="-TABLE_REG_EDITED-",
                    col_widths=[40, 20, 15, 30, 15],
                    background_color="#045D5D",
                    text_color="white",
                    bind_return_key=True,
                    row_colors=[("white", "#045D5D"), ("white", "yellow"), ("white", "red")],
                    enable_events=True,
                    enable_click_events=True,
                    select_mode=sg.TABLE_SELECT_MODE_EXTENDED,
                    tooltip=None,
                    expand_x=True,
                    expand_y=True
                ),
            ],
        ], expand_x=True, expand_y=True)],
        [sg.Push(), sg.Button("Save", key="-SAVE_RESTORE-"),
         sg.Button("Discard", key="-DISCARD_RESTORE-")]
    ]

    restore_window = sg.Window(f"Restore Registry - {file_name}", layoutRestore, finalize=True, resizable=True)
    restore_window.maximize()

    while True:
        event_restore, values_restore = restore_window.read()

        if event_restore == sg.WIN_CLOSED:
            restore_window.close()
            break

        if event_restore == "-SAVE_RESTORE-":
            confirmation = sg.popup_yes_no("Do you want to save the changes and restore the selected backup?", title="Confirmation")
            if confirmation == "Yes":
                # Create a backup of the current golden file in the "Replaced Rev" folder
                current_time = datetime.datetime.now().strftime("%Y-%m-%d")
                scurrent_time = str(current_time)
                replaced_rev_folder = os.path.join("Backup", "Backup(Golden File)", "Replaced Rev", file_name)

                if not os.path.exists(replaced_rev_folder):
                    os.makedirs(replaced_rev_folder)

                backup_file_name = f"{file_name}_rev{get_next_revision_number(replaced_rev_folder, file_name)}_{scurrent_time}.json"
                backup_file_path = os.path.join(replaced_rev_folder, backup_file_name)

                # Copy current golden file to the "Replaced Rev" folder
                shutil.copy2(golden_file_path, backup_file_path)

                # Replace the current golden file with the selected backup
                shutil.copy2(restore_file_path, golden_file_path)
                sg.popup_ok("Backup successfully restored.")
                os.remove(action_temp_path)

                # Clear the table and reset the state on Page 4 (similar to makeWinSave)
                window['-TABLE_EDITOR-'].update(values=[])  # Clear the table
                window['Dropdown2'].update(value='MicroAOI')  # Reset dropdown
                window['-BROWSE2-'].update('Browse')  # Clear browse button path
                window['-BROWSE2-'].update(disabled=False)  # Re-enable browse button

                # Reset buttons
                window['-ADD_NEW-'].update(visible=False)
                window['-EDIT_GOLDEN_FILE-'].update(visible=False, disabled=True)
                window['-DELETE_FROM_GOLDEN_FILE-'].update(visible=False, disabled=True)
                window['-SAVE_GOLDEN_FILE-'].update(visible=False, disabled=True)

                restore_window.close()
            else:
                sg.popup_ok("Restore action cancelled.")

        if event_restore == "-DISCARD_RESTORE-":
            confirmation = sg.popup_yes_no("Are you sure you want to discard the changes?")
            if confirmation == "Yes":
                os.remove(action_temp_path)  # Remove the action temp file
                sg.popup_ok("Changes discarded.")

                # Clear the table and reset the state on Page 4 (similar to makeWinSave)
                window['-TABLE_EDITOR-'].update(values=[])  # Clear the table
                window['Dropdown2'].update(value='MicroAOI')  # Reset dropdown
                window['-BROWSE2-'].update('Browse')  # Clear browse button path
                window['-BROWSE2-'].update(disabled=False)  # Re-enable browse button

                # Reset buttons
                window['-ADD_NEW-'].update(visible=False)
                window['-EDIT_GOLDEN_FILE-'].update(visible=False, disabled=True)
                window['-DELETE_FROM_GOLDEN_FILE-'].update(visible=False, disabled=True)
                window['-SAVE_GOLDEN_FILE-'].update(visible=False, disabled=True)

                restore_window.close()
            else:
                sg.popup_ok("Discard action cancelled.")


# Function to write changes to the {machine_type}_edit.temp file
def write_edit_temp_file(machine_type, changes, action):
    # Get the original golden file name to create the edit temp file
    golden_file_name = f"{machine_type}"
    edit_temp_file_name = f"{golden_file_name}_edit.temp"
    edit_temp_file_path = os.path.join("data", edit_temp_file_name)

    # Load the current edit temp file contents if they exist
    edit_temp_data = load_registry_from_json2(edit_temp_file_path)

    # If the file doesn't exist or is empty, initialize an empty list
    if edit_temp_data is None:
        edit_temp_data = []

    # Process each change
    for change in changes:
        conflict_found = False
        print(f"Processing change: {change}")  # Debugging

        # Convert list to dict if needed
        if isinstance(change, list):
            change = {
                'Registry Key/Subkey Path': change[0],  # path
                'Registry Name': change[1],  # name
                'Data': change[2],  # current data
                'Type': change[3],  # current type
                'Action': action
            }

        # Check for blank data and replace with "-"
        if change['Data'] == "":
            change['Data'] = "-"

        # Check for existing path and name in edit.temp
        for temp_entry in edit_temp_data:
            if temp_entry['Registry Key/Subkey Path'].lower() == change['Registry Key/Subkey Path'].lower() and \
                    temp_entry['Registry Name'] == change['Registry Name']:

                # If the entry was added and now being deleted, change the action to "Delete"
                if temp_entry['Action'] == 'Add' and action == 'Delete':
                    print(f"Found existing 'Add' entry, changing action to 'Delete': {temp_entry}")
                    temp_entry['Action'] = 'Delete'
                    conflict_found = True
                    break  # Exit once the match is found and updated

                # If the entry was added (action "Add"), update type/data and leave action as "Add"
                elif temp_entry['Action'] == 'Add' and action == 'Edit':
                    print(f"Found existing 'Add' entry, updating data and type: {temp_entry}")
                    temp_entry['Data'] = change['Data']
                    temp_entry['Type'] = change['Type']
                    temp_entry['Action'] = 'Add'  # Keep it as "Add"
                    conflict_found = True
                    break

        # If no conflict found, handle as a new "Edit" or "Add"
        if not conflict_found:
            print(f"No match in edit.temp, processing as new {action}: {change}")
            change['Action'] = action
            edit_temp_data.append(change)

    # Write the updated data back to the edit temp file
    print(f"Final edit.temp contents before writing: {edit_temp_data}")
    write_to_json(edit_temp_data, edit_temp_file_path)

    print(f"{edit_temp_file_name} updated with new changes: {edit_temp_data}")


def preserve_table_state(table_values):
    """Preserve the checkbox and sorting state before updating the table."""
    # Save the checkbox states
    checkbox_states = [row[0] for row in table_values]  # Store the checkbox state (Checked or Blank)

    # Save sorting order if necessary (e.g., which column was sorted and in which direction)
    # Assuming the sorting order is already stored in the handle_table_selection_page4
    return checkbox_states


def reapply_table_state(window, table_key, table_values, checkbox_states):
    """Reapply the checkbox and sorting state after updating the table."""
    for i, row in enumerate(table_values):
        # Restore the checkbox state
        row[0] = checkbox_states[i] if i < len(checkbox_states) else BLANK_BOX  # Reapply checkbox states

    # Update the table with the restored state
    window[table_key].update(values=table_values)


def log_deleted_registry_key(path, name, reg_type, data):
    log_message = f"User removed registry key: Path='{path}', Name='{name}', Type='{reg_type}', Data='{data}' from the table."
    write_into_event_log(log_message)


def perform_redundant_search(search_text, window, event, values):
    table_data = window["-TABLE_REDUNDANT-"].get()
    # Filter the table_data to include only rows where the registry key or name contains the search_text
    filtered_data = [row for row in table_data if search_text in row[1].lower() or search_text in row[2].lower()]
    window["-TABLE_REDUNDANT-"].update(values=filtered_data)


def log_deleted_items_page4(deleted_items):
    deleted_data_file = "data/deleted_registry_data_page4.json"
    try:
        if os.path.exists(deleted_data_file):
            with open(deleted_data_file, 'r') as file:
                existing_data = json.load(file)
        else:
            existing_data = []

        # Add the new deleted items to the existing ones
        existing_data.extend(deleted_items)

        with open(deleted_data_file, 'w') as file:
            json.dump(existing_data, file, indent=4)
        print("Deleted items logged successfully.")

    except Exception as e:
        print(f"Error logging deleted items: {e}")


def load_deleted_items_page4():
    deleted_data_file = "data/deleted_registry_data_page4.json"
    try:
        if os.path.exists(deleted_data_file):
            with open(deleted_data_file, 'r') as file:
                return json.load(file)
        else:
            return []
    except Exception as e:
        print(f"Error loading deleted items: {e}")
        return []


'''
def initialize_restore_table_page4():
    try:
        # Load the deleted data from the JSON file
        deleted_data_file = "data/deleted_registry_data_page4.json"
        if os.path.isfile(deleted_data_file):
            with open(deleted_data_file, 'r') as file:
                restored_table_data = json.load(file)
        else:
            restored_table_data = []

        # Only update the table and show the window if there is data to restore
        if restored_table_data:
            window_restore_page4['-TABLE_RESTORE_PAGE4-'].update(values=restored_table_data)
            window['-RESTORE_PAGE4-'].update(disabled=False)  # Enable the restore button if data exists
        else:
            window['-RESTORE_PAGE4-'].update(disabled=True)  # Disable restore button if no data exists

    except Exception as e:
        sg.popup_error(f"Error initializing restore table: {e}")


def check_and_enable_restore_button_page4():
    deleted_items = load_deleted_items_page4()  # This loads items from the file
    if deleted_items:
        window['-RESTORE_PAGE4-'].update(disabled=False)  # Enable the restore button if items exist
    else:
        window['-RESTORE_PAGE4-'].update(disabled=True)  # Disable if no items


def write_restored_data_to_json_page4(restored_data, file_path):
    """
    Appends restored data to the existing JSON file without overwriting the current content.

    :param restored_data: List of restored registry entries.
    :param file_path: Path to the JSON file.
    """
    try:
        # Load the existing data from the JSON file
        with open(file_path, 'r') as json_file:
            try:
                current_data = json.load(json_file)
            except json.JSONDecodeError:
                sg.popup_error("Error: JSON data in the file is invalid.")
                current_data = []

        # Append the restored data to the existing data
        updated_data = current_data + restored_data

        # Write the updated data back to the JSON file
        with open(file_path, 'w') as json_file:
            json.dump(updated_data, json_file, indent=4)

        print(f"Data successfully written to {file_path}")

    except IOError as e:
        sg.popup_error(f"Error writing to file {file_path}: {e}")
        write_into_error_log(content=f"Error writing to file {file_path}: {e}")

'''

def write_to_json_for_page4(data, file_path):
    """
    Write reordered registry data to a JSON file for Page 4.

    :param data: Reordered data (list of registry entries)
    :param file_path: The path to the JSON file where data should be saved
    """
    try:
        # Write the correct_ordered_data into the specified file
        with open(file_path, 'w') as file:
            json.dump(data, file, indent=4)
        print(f"Data successfully written to {file_path}")

    except IOError as e:
        print(f"Error writing to file {file_path}: {e}")
        write_into_error_log(content=f"Error writing to file {file_path}: {e}")


def backup_deleted_registry_page4(exported_file, backup_folder):
    """Backup deleted registry data specifically for Page 4."""
    os.makedirs(backup_folder, exist_ok=True)

    # Generate a unique backup file name using a timestamp
    timestamp = datetime.datetime.now().strftime('%d-%m-%Y_%H%M%S')
    backup_filename = f"deleted_registry_data_page4_{timestamp}.json"
    backup_path = os.path.join(backup_folder, backup_filename)

    try:
        # Read the content of the exported (selected) file for deleted data
        with open(exported_file, 'r') as file:
            exported_data = json.load(file)

        # Check if the exported data is not empty
        if not exported_data:
            print("Warning: No data found to backup for Page 4: Registry Editor.")
            write_into_error_log(content=f"No data found to backup for Page 4: Registry Editor in file: {exported_file}")
            return

        # Ensure the correct order of the fields before writing to JSON
        correct_ordered_data = []
        for entry in exported_data:
            # Ensure that the data follows the format [Path, Name, Type, Data]
            reordered_entry = [
                entry[0],  # Registry Path
                entry[1],  # Registry Name
                entry[3],  # Registry Type (should be at 3rd index)
                entry[2]   # Registry Data (should be at 4th index)
            ]
            correct_ordered_data.append(reordered_entry)

        # Debug print to check exported data before writing
        print("Data after reordering:", correct_ordered_data)

        # Use the new write function specifically for Page 4
        write_to_json_for_page4(correct_ordered_data, backup_path)

        print(f"Deleted item backup created for Page 4: Registry Editor: {backup_path}")
        write_into_event_log(content=f"Deleted item backup for Page 4: Registry Editor has been saved to: {backup_path}")

    except FileNotFoundError:
        print(f"Error: File '{exported_file}' not found.")
        write_into_error_log(content=f"Error: File '{exported_file}' not found.")

    except IOError as e:
        print(f"Error writing file for Page 4: {e}")
        write_into_error_log(content=f"Error writing file for Page 4: Registry Editor: {str(e)}")

'''
def update_json_file(file_path, key, name, reg_type, reg_data):
    try:
        # Load the existing file
        with open(file_path, 'r') as json_file:
            data = json.load(json_file)

        # Add or update the restored key data
        data.append([key, name, reg_type, reg_data])

        # Write it back to the file
        with open(file_path, 'w') as json_file:
            json.dump(data, json_file, indent=4)

        print(f"Filtered data written to {file_path}")
    except Exception as e:
        print(f"Error writing to file: {e}")
        write_into_error_log(f"Error writing to file {file_path}: {e}")


def handle_table_selection_page4(window, event, values):
    selected_list_index = []
    result = []  # Holds the selected row data

    if isinstance(event, tuple) and event[0] == '-TABLE_RESTORE_PAGE4-':
        row_index = event[2][0]
        column_index = event[2][1]

        # Log header click event for debugging
        print(f"Header clicked: row_index={row_index}, column_index={column_index}")

        # Retrieve the table values
        table_values = window["-TABLE_RESTORE_PAGE4-"].get()

        # Handle header click: Select/Deselect all checkboxes
        if column_index == 0 and row_index == -1:
            if not hasattr(handle_table_selection_page4, 'click_time'):
                handle_table_selection_page4.click_time = 0

            # Toggle between select all and deselect all
            if handle_table_selection_page4.click_time == 0:
                handle_table_selection_page4.click_time = 1
                for row in table_values:
                    row[0] = CHECKED_BOX  # Set all checkboxes to checked
            else:
                handle_table_selection_page4.click_time = 0
                for row in table_values:
                    row[0] = BLANK_BOX  # Set all checkboxes to unchecked

            # Update the table with the new values
            window["-TABLE_RESTORE_PAGE4-"].update(values=table_values)

            # Enable/Disable restore button based on selection
            selected_list_index = [i for i, row in enumerate(table_values) if row[0] == CHECKED_BOX]
            window.find_element("-restoreSelectedPage4-").update(disabled=not bool(selected_list_index))

        # Handle sorting for other columns
        elif row_index == -1 and column_index > 0:
            print(f"Sorting column: {column_index}")  # Log sorting event
            # Implement sorting logic here if needed

        else:
            # Handle row clicks (non-header clicks)
            if isinstance(row_index, int) and row_index >= 0:
                if 0 <= row_index < len(table_values):
                    clicked_row_data = table_values[row_index]
                    if clicked_row_data[0] == BLANK_BOX:
                        clicked_row_data[0] = CHECKED_BOX  # Check row checkbox
                    else:
                        clicked_row_data[0] = BLANK_BOX  # Uncheck row checkbox

                    # Update the table with the new values
                    table_values[row_index] = clicked_row_data
                    window["-TABLE_RESTORE_PAGE4-"].update(values=table_values)

                    # Enable/Disable restore button based on selection
                    selected_list_index = [i for i, row in enumerate(table_values) if row[0] == CHECKED_BOX]
                    window.find_element("-restoreSelectedPage4-").update(disabled=not bool(selected_list_index))

    return result


def restore_selected_data(golden_file, deleted_data):
    """
    Restores the deleted registry keys back into the golden file.

    :param golden_file: Path to the golden file where data is stored.
    :param deleted_data: The list of deleted registry keys to restore.
    """
    try:
        # Load the current data in the golden file
        current_data = load_registry_from_json2(golden_file)

        # Append the deleted data back into the current data
        for item in deleted_data:
            current_data.append(item)

        # Write the updated data back to the golden file
        write_new_data_to_json(current_data, golden_file)

        print("Restored the deleted data back into the golden file.")

    except Exception as e:
        print(f"Error restoring data: {e}")
        write_into_error_log(f"Error restoring data: {e}")


def generate_current_registry_json_page4(output_file):
    """Generate the current registry data for Page 4 and save it as a JSON file."""
    # Assuming this is a mockup of the registry data for Page 4
    registry_data_page4 = [
        {
            'Registry Key/Subkey Path': 'Software\\WOW6432Node\\MV Technology',
            'Registry Name': '111',
            'Type': 'REG_SZ',
            'Data': 'Some data here'
        },
        # Add actual registry data here
    ]

    try:
        with open(output_file, 'w') as json_file:
            json.dump(registry_data_page4, json_file, indent=4)
        print(f"Current registry data for Page : Registry Editor saved to {output_file}")
    except IOError as e:
        print(f"Error writing to {output_file}: {e}")

# Conversion from JSON to .txt
def generate_txt_from_json(json_file_path, txt_file_path):
    """Generate a .txt file from the registry data in a .json file."""
    try:
        with open(json_file_path, 'r') as json_file:
            data = json.load(json_file)

        # Write data to a .txt file in a readable format
        with open(txt_file_path, 'w') as txt_file:
            for entry in data:
                txt_file.write(json.dumps(entry, indent=4) + '\n')

        print(f".txt file created from {json_file_path} at {txt_file_path}")

    except FileNotFoundError:
        print(f"Error: File '{json_file_path}' not found.")
    except IOError as e:
        print(f"Error creating .txt file: {e}")

def delete_selected_keys_page4(selected_data_page4):
    """ Delete registry keys for Page 4 based on the given data. """
    try:
        # Connect to the registry (e.g., HKEY_LOCAL_MACHINE in this case)
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as hkey:
            for key in selected_data_page4:
                key_path = key['Registry Key/Subkey Path']
                value_name = key['Registry Name']

                # Open the key and delete the specific value
                try:
                    with winreg.OpenKey(hkey, key_path, 0, winreg.KEY_ALL_ACCESS) as key:
                        winreg.DeleteValue(key, value_name)
                        print(f"Deleted value '{value_name}' from key '{key_path}' successfully.")
                        write_into_event_log(f"Deleted value '{value_name}' from key '{key_path}' successfully.")
                except OSError as e:
                    print(f"Error deleting value '{value_name}' from key '{key_path}': {e}")
                    write_into_error_log(f"Error deleting value '{value_name}' from key '{key_path}': {e}")

    except Exception as e:
        print(f"Unexpected error: {e}")
        write_into_error_log(f"Unexpected error: {e}")

def backup_registry_page4(exported_file3, backup_folder3):
    """Backup function specifically for Page 4, backing up the full registry data before deletion."""
    join_path = os.path.join("data", exported_file3)
    src_file_path = get_current_file_path(join_path)

    # Ensure the backup folder exists
    os.makedirs(backup_folder3, exist_ok=True)

    # Generate a unique timestamped backup filename for the .txt file
    timestamp = datetime.datetime.now().strftime('%d-%m-%Y_%H%M%S')
    backup_filename_txt = f"{exported_file3[:-4]}_{timestamp}.txt"
    backup_path_txt = os.path.join(backup_folder3, backup_filename_txt)

    # JSON backup filename
    backup_filename_json = "current_pc_registry_data_page4.json"
    json_path = os.path.join('data', backup_filename_json)

    try:
        # Fetch the table data and remove any unwanted symbols
        full_registry_data = window['-TABLE_EDITOR-'].get()  # Fetch the data from the table (Page 4)

        # Cleaned registry data (removing checkbox markers)
        cleaned_registry_data = [
            entry[1:] for entry in full_registry_data  # Strip the first element (checkbox)
        ]

        # Save the cleaned data into the JSON file in 'data' folder
        with open(json_path, 'w', encoding='utf-8') as json_file:
            json.dump(cleaned_registry_data, json_file, indent=4)

        print(f"Updated JSON file: {json_path}")
        write_into_event_log(content=f"Updated JSON file: {json_path}")

        # Now also write the cleaned data to the .txt backup file
        with open(backup_path_txt, 'w', encoding='utf-8') as backup_file:
            for entry in cleaned_registry_data:
                backup_file.write(f"{json.dumps(entry, indent=4)}\n")

        print(f"Full registry backup created for Page 4: Registry Editor in .txt format: {backup_path_txt}")
        write_into_event_log(content=f"The full registry backup file for Page 4: Registry Editor has been saved to: {backup_path_txt}")

    except FileNotFoundError:
        print(f"Error: File '{exported_file3}' not found for Page 4: Registry Editor backup.")
        write_into_error_log(content=f"Error: File '{exported_file3}' not found for Page 4: Registry Editor backup.")

    except IOError as e:
        print(f"Error writing file for Page 4: Registry Editor backup: {e}")
        write_into_error_log(content=f"Error writing file for Page 4: Registry Editor backup: {str(e)}")
'''

def load_registry_from_json2(file_path):
    """Load registry data from the specified JSON file."""
    try:
        with open(file_path, 'r') as json_file:
            data = json.load(json_file)
        return data
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Error loading registry data from {file_path}: {e}")
        return []


def get_current_directory():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def get_current_file_path(filename):
    return os.path.join(get_current_directory(), filename)


def write_to_json(data, filename):
    file_path = get_current_file_path(filename)
    try:
        with open(file_path, 'w') as file:
            json.dump(data, file, indent=4)
        # print(f"File written successfully to {file_path}")

    except IOError as e:
        print(f"Error writing to file {file_path}: {e}")


# Globale defined the log txt file naming
event_log_file_path = get_current_file_path('log//[EventLog]_' + scurrent_time + '.txt')
error_log_file_path = get_current_file_path('log//[ErrorLog]_' + scurrent_time + '.txt')


def write_into_event_log(content):
    # inside the content
    current_time_content = datetime.datetime.now().strftime("%m/%d/%Y %H:%M:%S %p: ")
    # scurrent_time_content=str(current_time_content)

    os.makedirs(os.path.dirname(event_log_file_path), exist_ok=True)
    # Read the last timestamp from the log file if it exists
    last_timestamp = None
    if os.path.exists(event_log_file_path):
        with open(event_log_file_path, "r") as f:
            lines = f.readlines()
            if lines:
                last_timestamp = lines[-3].strip()

    # Write to the log file if the current timestamp is different from the last timestamp
    if last_timestamp != current_time_content:
        with open(event_log_file_path, "a") as f:
            f.write(current_time_content + "\n")
            f.write("\t" + content + "\n\n")
    else:
        f.write("\t" + content + "\n\n")


def write_into_error_log(content):
    # inside the content
    current_time_content = datetime.datetime.now().strftime("%m/%d/%Y %H:%M:%S %p: ")
    # current_time_content=str(current_time_content)

    os.makedirs(os.path.dirname(error_log_file_path), exist_ok=True)
    # Read the last timestamp from the log file if it exists
    last_timestamp = None
    if os.path.exists(event_log_file_path):
        with open(event_log_file_path, "r") as f:
            lines = f.readlines()
            if lines:
                last_timestamp = lines[-3].strip()

    # Write to the log file if the current timestamp is different from the last timestamp
    if last_timestamp != current_time_content:
        with open(event_log_file_path, "a") as f:
            f.write(current_time_content + "\n")
            f.write("\t" + content + "\n\n")
    else:
        f.write("\t" + content + "\n\n")


def get_sample_file():
    # to choose the sample file (got 4 type, semicon, smt, microAOI, side cam)
    file_exist = False
    # to get the sample file user want to use for compaere the registry
    while file_exist == False:
        file_path = sg.popup_get_file("Select 'REGISTRY' file for Registry Checking", title="File selector",
                                      no_window=True, file_types=(('JSON', '*.json'),))
        if not Path(file_path).is_file():
            if file_path == '':
                sg.popup_ok('Plese select a file ', title="Error")
                file_exist = False
            else:
                sg.popup_ok('File not exist', title="Error")
                file_exist = False
        else:
            file_exist = True

    return file_path


def run_registry_main(file_path):
    # Generate the current PC registry data
    join_path = os.path.join("data", "current_pc_registry_data.json")
    os.makedirs(os.path.dirname(join_path), exist_ok=True)
    # current_pc_reg_output_file = merged_current_registry_data("current_pc_registry_data.json")
    current_pc_reg_output_file = merged_current_registry_data(join_path)

    # Load the software PC data and current PC registry data
    # software_pc_data = load_registry_from_json('C:/Users/xiau-yen.kelly-teo/Desktop/MyRegistryChecker/sample-registry.json')
    software_pc_data = load_registry_from_json(file_path)
    current_pc_reg_data = load_registry_from_json(current_pc_reg_output_file)

    # Generate the registry data results
    results_data = generate_reg_data(software_pc_data, current_pc_reg_data)

    # Define the result file path
    result_reg_data_file = "data/result_reg_data.json"

    # result_reg_data_file_path= get_current_file_path('result_reg_data.json')
    write_to_json(results_data, result_reg_data_file)

    # Update the GUI with the latest data
    update_reg_gui()


# for list of fail/missing keys result in summary page(upper table)
def run_registry_main2(file_path):
    # Generate the current PC registry data
    join_path = os.path.join("data", "current_pc_registry_data.json")
    os.makedirs(os.path.dirname(join_path), exist_ok=True)
    current_pc_reg_output_file = merged_current_registry_data(join_path)

    # Load the software PC data and current PC registry data
    # software_pc_data = load_registry_from_json('C:/Users/xiau-yen.kelly-teo/Desktop/MyRegistryChecker/sample-registry.json')
    sample_reg_pc_data = load_registry_from_json(file_path)
    current_pc_reg_data = load_registry_from_json(current_pc_reg_output_file)

    # Generate the registry data results
    results_data = generate_compared_results(sample_reg_pc_data, current_pc_reg_data)

    # Define the result file path
    result_reg_data_file = "data/compared_result_reg_data.json"

    write_to_json(results_data, result_reg_data_file)


# for redundant registry key data in summary page (bottom table)
def run_registry_main3(file_path):
    # Generate the current PC registry data
    join_path = os.path.join("data", "current_pc_registry_data.json")
    os.makedirs(os.path.dirname(join_path), exist_ok=True)
    current_pc_reg_output_file = merged_current_registry_data(join_path)

    # Load the software PC data and current PC registry data
    # software_pc_data = load_registry_from_json('C:/Users/xiau-yen.kelly-teo/Desktop/MyRegistryChecker/sample-registry.json')
    sample_reg_pc_data = load_registry_from_json(file_path)
    current_pc_reg_data = load_registry_from_json(current_pc_reg_output_file)

    # Generate the registry data results
    results_data = generate_redundant_data(sample_reg_pc_data, current_pc_reg_data)

    # Define the result file path
    result_reg_data_file = "data/redundant_data.json"

    write_to_json(results_data, result_reg_data_file)
    # Save the results to the JSON file


def get_redundant_key_count():
    json_file_path = get_current_file_path('data/redundant_data.json')
    os.makedirs(os.path.dirname(json_file_path), exist_ok=True)

    try:
        # Read and load the JSON file
        if os.path.isfile(json_file_path) and os.path.getsize(json_file_path) > 0:
            with open(json_file_path, 'r') as json_file:
                results = json.load(json_file)
            # Count entries with "Missing" status
            num_reg_redundant = sum(1 for result_data in results.values() if result_data.get("Status") == "Missing")
            return num_reg_redundant
    except Exception as e:
        write_into_error_log(content="An error occurred while retrieving redundant key count: " + str(e))
        return 0  # Return 0 in case of any error

# update the registry table in first page
def update_reg_gui():
    try:
        json_file_path = get_current_file_path('data/result_reg_data.json')
        os.makedirs(os.path.dirname(json_file_path), exist_ok=True)

        # Initialize counters
        num_reg_fail = num_reg_not_found = num_not_exist_keys = num_exist_keys = num_reg_pass = 0

        # Check if the JSON file exists and is not empty
        if not os.path.isfile(json_file_path):
            write_into_error_log(content="The file '" + json_file_path + "' does not exist.")
            raise FileNotFoundError(f"The file {json_file_path} does not exist.")

        if os.path.getsize(json_file_path) == 0:
            write_into_error_log(content="The file '" + json_file_path + "' is empty.")
            raise ValueError(f"The file {json_file_path} is empty.")

        # Read and load the JSON file
        with open(json_file_path, 'r') as json_file:
            try:
                results = json.load(json_file)
            except json.JSONDecodeError:
                write_into_error_log(content="The file '" + json_file_path + "' contains invalid JSON data.")
                raise ValueError(f"The file {json_file_path} contains invalid JSON data.")

        # Process main data from result_reg_data.json
        num_reg_required = len(results)
        num_reg_found = sum(1 for result_data in results.values() if result_data.get("Status") in ["Pass", "Fail", "Exist", "Not Exist"])
        num_reg_not_found = sum(1 for result_data in results.values() if result_data.get("Status") == "Missing")
        num_not_exist_keys = sum(1 for result_data in results.values() if result_data.get("Status") == "Not Exist")
        num_reg_fail = sum(1 for result_data in results.values() if result_data.get("Status") == "Fail")
        num_exist_keys = sum(1 for result_data in results.values() if result_data.get("Status") == "Exist")
        num_reg_pass = sum(1 for result_data in results.values() if result_data.get("Status") == "Pass")

        # Calculate the Total Required Registry Keys
        total_required_registry_keys = num_reg_not_found + num_not_exist_keys + num_reg_found

        # Total Registry Keys Found (Failed + Exist + Matched)
        total_registry_keys_found = num_reg_fail + num_exist_keys + num_reg_pass

        def calculate_percentage(count):
            return (count / num_reg_required) * 100 if num_reg_required > 0 else 0

        # Update counters in the UI
        window["-TOTAL_REQUIRED_REG-"].update(f"{total_required_registry_keys}")
        window["-MISSING_REG-"].update(f"{num_reg_not_found} ({calculate_percentage(num_reg_not_found):.2f}%)")
        window["-NOT_EXIST_REG-"].update(f"{num_not_exist_keys} ({calculate_percentage(num_not_exist_keys):.2f}%)")

        window["-TOTAL_FOUND_REG-"].update(f"{total_registry_keys_found}")
        window["-FAIL_REG-"].update(f"{num_reg_fail} ({calculate_percentage(num_reg_fail):.2f}%)")
        window["-EXIST_REG-"].update(f"{num_exist_keys} ({calculate_percentage(num_exist_keys):.2f}%)")
        window["-MATCHED_REG-"].update(f"{num_reg_pass} ({calculate_percentage(num_reg_pass):.2f}%)")

        # Prepare the table data and update the table
        updated_reg_data = [
            [
                None,  # Placeholder for the "List" index, which we'll add after sorting
                result_reg_data["Registry Key/Subkey Path"],
                result_reg_data["Registry Name"],
                result_reg_data["Registry Type"],
                result_reg_data["Data"],
                result_reg_data["Status"]
            ]
            for result_reg_data in results.values()
        ]

        # Define sorting order for the "Status" field: Fail, Missing, Exist, Not Exist, Pass
        status_order = {"Fail": 1, "Missing": 2, "Exist": 3, "Not Exist": 4, "Pass": 5}

        # Sort updated_reg_data based on the "Status" using our custom order
        sorted_data = sorted(updated_reg_data, key=lambda x: status_order.get(x[5], 6))

        # Assign fixed sequential indices starting from 1 after sorting
        final_data = [
            [index + 1] + row[1:]  # Adding sequential index to the "List" column
            for index, row in enumerate(sorted_data)
        ]

        new_color = set_status_color(final_data)
        window["-TABLE_REG-"].update(values=final_data, row_colors=new_color)

    except Exception as e:
        sg.popup_error(f"An error occurred while updating the table: {e}", title="Error")
        write_into_error_log(content="An error occurred while updating the table:" + str(e))

    return final_data


# Set the row color in the table based on the status
def set_status_color(table_data):
    row_colors = []
    for i, row in enumerate(table_data):
        status = row[5]

        if "Pass" in status:
            row_colors.append((i, "white", "#045D5D"))
        elif "Fail" in status:
            row_colors.append((i, "black", "red"))
        elif "Missing" in status:
            row_colors.append((i, "black", "yellow"))
        elif "Not Exist" in status:
            row_colors.append((i, "black", "orange"))
        elif "Exist" in status:
            row_colors.append((i, "white","purple"))

    return (row_colors)


# update the compared registry table in summary page (upper page)
def update_reg_compared_gui(window):
    # json_file_path = 'C:/Users/xiau-yen.kelly-teo/Desktop/SW & Size checker - Copy/compared_result_reg_data.json'
    json_file_path = get_current_file_path('data/compared_result_reg_data.json')
    os.makedirs(os.path.dirname(json_file_path), exist_ok=True)
    try:

        # Check if the JSON file exists and is not empty
        if not os.path.isfile(json_file_path):
            write_into_error_log(content="The file '" + json_file_path + "' does not exist.")
            raise FileNotFoundError(f"The file {json_file_path} does not exist.")
        if os.path.getsize(json_file_path) == 0:
            write_into_error_log(content="The file '" + json_file_path + "' is empty.")
            raise ValueError(f"The file {json_file_path} is empty.")

        # Read and load the JSON file
        with open(json_file_path, 'r') as json_file:
            try:
                results = json.load(json_file)
            except json.JSONDecodeError:
                write_into_error_log(content="The file '" + json_file_path + "' contains invalid JSON data.")
                raise ValueError(f"The file {json_file_path} contains invalid JSON data.")

        updated_reg_compare_data = [
            [
                BLANK_BOX,
                compared_result_reg_data["Registry Key/Subkey Path"],
                compared_result_reg_data["Registry Name"],
                compared_result_reg_data["Current Type"],
                compared_result_reg_data["Expected Type"],
                compared_result_reg_data["Current Data"],
                compared_result_reg_data["Expected Data"],
                compared_result_reg_data["Status"]
            ]
            for i, compared_result_reg_data in enumerate(results.values())
        ]

        filtered_data = []
        for row in updated_reg_compare_data:
            _, _, _, _, _, _, _, status = row
            if status == "Fail" or status == "Missing":
                filtered_data.append(row)

        if filtered_data:

            window["-TABLE_REG_COMPARED-"].update(values=filtered_data)

        else:
            window["-TABLE_REG_COMPARED-"].update(values=filtered_data)

        num_reg_fail = sum(1 for result_data in results.values() if result_data.get("Status") == "Fail")
        num_reg_not_found = sum(
            1 for result_data in results.values() if result_data.get("Status") in ["Missing", "Not Exist"])

        window["-FAIL_REG-"].update(num_reg_fail)
        window["-NOT_FOUND_REG-"].update(num_reg_not_found)

    except Exception as e:
        write_into_error_log(content="An error occurred while updating the table: " + e)
        sg.popup_error(f"An error occurred while updating the table: {e}", title="Error")


# update the redundant key table in summary page (bottom table)
def update_redundant_gui(window, search_text=None):
    json_file_path = get_current_file_path('data/redundant_data.json')
    os.makedirs(os.path.dirname(json_file_path), exist_ok=True)

    try:
        # Check if the JSON file exists and is not empty
        if not os.path.isfile(json_file_path):
            write_into_error_log(content="The file '" + json_file_path + "' does not exist.")
            raise FileNotFoundError(f"The file {json_file_path} does not exist.")
        if os.path.getsize(json_file_path) == 0:
            write_into_error_log(content="The file '" + json_file_path + "' is empty.")
            raise ValueError(f"The file {json_file_path} is empty.")

        # Read and load the JSON file
        with open(json_file_path, 'r') as json_file:
            try:
                results = json.load(json_file)
            except json.JSONDecodeError:
                write_into_error_log(content="The file '" + json_file_path + "' contains invalid JSON data.")
                raise ValueError(f"The file {json_file_path} contains invalid JSON data.")

        # Count the redundant keys (independent of search filter)
        num_reg_redundant = sum(1 for result_data in results.values() if result_data.get("Status") == "Missing")
        window["-REDUNDANT_REG-"].update(num_reg_redundant)

        # Filter data by "Missing" status and exclude "Default" keys
        redundant_data = [
            [
                BLANK_BOX,
                redundant_data["Registry Key/Subkey Path"],
                redundant_data["Registry Name"],
                redundant_data["Registry Type"],
                redundant_data["Data"],
                redundant_data["Status"]
            ]
            for redundant_data in results.values()
            if redundant_data.get("Status") == "Missing" and redundant_data.get("Registry Name") != "Default"
        ]

        # Apply search filter only for display, if search_text is provided
        if search_text:
            search_text = search_text.lower()
            redundant_data = [row for row in redundant_data if search_text in row[1].lower()]

        # Update the table with the filtered data
        window["-TABLE_REDUNDANT-"].update(values=redundant_data)

    except Exception as e:
        write_into_error_log(content="An error occurred while updating the table: " + str(e))
        sg.popup_error(f"An error occurred while updating the table: {e}", title="Error")


# update the restore/deleted keys in the
def update_restore_gui(window):
    # json_file_path = 'C:/Users/xiau-yen.kelly-teo/Desktop/SW & Size checker - Copy/redundant_data.json'
    get_deleted_redundant_path = os.path.join("data", "list_of_deleted_keys.json")
    json_file_path = get_current_file_path(get_deleted_redundant_path)
    # os.makedirs(os.path.dirname(json_file_path), exist_ok=True)

    try:

        # Check if the JSON file exists and is not empty
        if not os.path.isfile(json_file_path):
            write_into_error_log(content="The file '" + json_file_path + "' does not exist.")
            raise FileNotFoundError(f"The file {json_file_path} does not exist.")
        if os.path.getsize(json_file_path) == 0:
            write_into_error_log(content="The file '" + json_file_path + "' is empty.")
            raise ValueError(f"The file {json_file_path} is empty.")

        # Read and load the JSON file
        with open(json_file_path, 'r') as json_file:
            try:
                results = json.load(json_file)
            except json.JSONDecodeError:
                write_into_error_log(content="The file '" + json_file_path + "' contains invalid JSON data.")
                raise ValueError(f"The file {json_file_path} contains invalid JSON data.")

        updated_delete_data = [
            [
                BLANK_BOX,
                delete_data.get("Registry Key/Subkey Path", ""),
                delete_data.get("Registry Name", ""),
                delete_data.get("Type", ""),
                delete_data.get("Data", "")
            ]
            for delete_data in results
        ]

        if updated_delete_data:

            window["-TABLE_RESTORE-"].update(values=updated_delete_data)
        else:

            window["-TABLE_RESTORE-"].update(values=updated_delete_data)
            window.find_element("-restoreSelected-").update(disabled=True)
        num_reg_delete = len(results)
        #window["-DEL_REG-"].update(num_reg_delete)

    except Exception as e:
        write_into_error_log(content="An error occurred while updating the table." + e)
        sg.popup_error(f"An error occurred while updating the table: {e}", title="Error")

'''
def update_restore_gui_page4(window):
    get_deleted_page4_path = os.path.join("data", "deleted_registry_data_page4.json")
    json_file_path = get_current_file_path(get_deleted_page4_path)

    try:
        if not os.path.isfile(json_file_path):
            write_into_error_log(content="The file '" + json_file_path + "' does not exist.")
            raise FileNotFoundError(f"The file {json_file_path} does not exist.")
        if os.path.getsize(json_file_path) == 0:
            write_into_error_log(content="The file '" + json_file_path + "' is empty.")
            raise ValueError(f"The file {json_file_path} is empty.")

        # Read and load the JSON file
        with open(json_file_path, 'r') as json_file:
            try:
                results = json.load(json_file)
            except json.JSONDecodeError:
                write_into_error_log(content="The file '" + json_file_path + "' contains invalid JSON data.")
                raise ValueError(f"The file {json_file_path} contains invalid JSON data.")

        # Retrieve the current state of the table (checkbox states)
        table_values = window["-TABLE_RESTORE_PAGE4-"].get()
        checkbox_states = [row[0] for row in table_values] if table_values else []

        # Force each row to have a checkbox in the first column
        updated_delete_data = [
            [
                BLANK_BOX,  # This will be overwritten by the preserved state
                delete_data[0],  # 'Registry Key/Subkey Path'
                delete_data[1],  # 'Registry Name'
                delete_data[2],  # 'Type'
                delete_data[3]   # 'Data'
            ]
            for delete_data in results
        ]

        # Update the table in the restore window with new data
        if updated_delete_data:
            # Reapply checkbox states
            for i, row in enumerate(updated_delete_data):
                if i < len(checkbox_states):
                    row[0] = checkbox_states[i]  # Restore the checkbox state

            window["-TABLE_RESTORE_PAGE4-"].update(values=updated_delete_data)
        else:
            window["-TABLE_RESTORE_PAGE4-"].update(values=[])
            window.find_element("-restoreSelectedPage4-").update(disabled=True)

    except Exception as e:
        write_into_error_log(content="An error occurred while updating the table: " + str(e))
        sg.popup_error(f"An error occurred while updating the table: {e}", title="Error")
'''

# compare the registry key between the current pc with the golden file chose (use in first page)
def match_reg_item(registry_path, registry_name, registry_type, registry_data, registry):
    try:
        current_registry_path, current_registry_name, current_registry_type, current_registry_data = registry
    except ValueError:
        write_into_error_log(content="Registry item has unexpected structure: " + registry)
        raise ValueError(f"Registry item has unexpected structure: {registry}")

    # Check name similarity
    if current_registry_name == "":
        registry_name = current_registry_name

        # Check if both path and name fully match
    if current_registry_path.lower() == registry_path.lower() and current_registry_name.lower() == registry_name.lower():
        # Check if either type is 'N/A'
        if registry_type == 'N/A' or current_registry_type == 'N/A':
            return "Pass"

        if registry_name in ['Password', 'calib3dpassword'] or current_registry_name in ['Password', 'calib3dpassword']:
            return "Exist"

        # Convert data to string for comparison if necessary
        if isinstance(current_registry_data, bytes):
            current_registry_data = " ".join([f"{byte:02X}" for byte in current_registry_data])
        if isinstance(registry_data, bytes):
            registry_data = " ".join([f"{byte:02X}" for byte in registry_data])

        # Compare type and data
        if current_registry_type == registry_type and current_registry_data == registry_data:
            return "Pass"
        else:
            return "Fail"

    elif current_registry_path.lower() != registry_path.lower() and current_registry_name != registry_name and registry_name not in [
        'Password', 'calib3dpassword']:

        return "Missing"  # If path or name does not fully match, return missing

    elif current_registry_path.lower() != registry_path.lower() and current_registry_name != registry_name and registry_name in [
        'Password', 'calib3dpassword']:

        return "Not Exist"


# Compare the registry keys between the golden file with the current pc to get the redundant keys
# Special case: current pc got, golden file dun have
def redundant_reg_keys(registry_path, registry_name, registry_type, registry_data, registry):
    try:
        current_registry_path, current_registry_name, current_registry_type, current_registry_data = registry
    except ValueError:
        write_into_error_log(content="Registry item has unexpected structure: " + str(registry))
        raise ValueError(f"Registry item has unexpected structure: {registry}")

    # Step 1: Check if the path and name do not fully match
    if current_registry_path.lower() != registry_path.lower() or current_registry_name.lower() != registry_name.lower():
        # Condition preserved from previous code
        if registry_name not in ['Password', 'calib3dpassword']:
            return "Missing"  # Indicates it's present on current PC but absent in the golden file
        else:
            return "Not Exist"  # Special case for 'Password' or 'calib3dpassword'

    # Step 2: Proceed with existing checks if the path and name fully match
    # Check if both path and name fully match
    if current_registry_path.lower() == registry_path.lower() and current_registry_name.lower() == registry_name.lower():
        # Check if either type is 'N/A'
        if registry_type == 'N/A' or current_registry_type == 'N/A':
            return "Pass"

        if registry_name in ['Password', 'calib3dpassword'] or current_registry_name in ['Password', 'calib3dpassword']:
            return "Exist"

        # Convert data to string for comparison if necessary
        if isinstance(current_registry_data, bytes):
            current_registry_data = " ".join([f"{byte:02X}" for byte in current_registry_data])
        if isinstance(registry_data, bytes):
            registry_data = " ".join([f"{byte:02X}" for byte in registry_data])

        # Compare type and data
        if current_registry_type == registry_type and current_registry_data == registry_data:
            return "Pass"
        else:
            return "Fail"


# generate the registry keys data and save to the result_reg_data.json
# show the data in the first page table
def generate_reg_data(software_pc_data, current_pc_reg_data):
    results = {}

    # Step 1: Create a lookup set for quick presence checks in current data
    current_keys = {(registry[0].lower(), registry[1].lower()) for registry in current_pc_reg_data}

    # Step 2: Process each entry in software data
    for registry_path, registry_name, registry_type, registry_data in software_pc_data:
        registry_key = (registry_path.lower(), registry_name.lower())

        # Initialize matched_registry to store potential matches for Pass
        matched_registry = []

        # Step 3: Check for special cases of Exist and Not Exist
        if any(match_reg_item(registry_path, registry_name, registry_type, registry_data, registry) == 'Exist'
               for registry in current_pc_reg_data):
            status = 'Exist'
            matched_registry_dict = {
                'Registry Key/Subkey Path': registry_path,
                'Registry Name': registry_name,
                'Registry Type': registry_data,
                'Data': registry_type
            }

        elif any(match_reg_item(registry_path, registry_name, registry_type, registry_data, registry) == 'Not Exist'
                 for registry in current_pc_reg_data):
            status = 'Not Exist'
            matched_registry_dict = {
                'Registry Key/Subkey Path': registry_path,
                'Registry Name': registry_name,
                'Registry Type': registry_data,
                'Data': registry_type
            }

        # Step 4: Handle Missing status if the key does not exist in current data
        elif registry_key not in current_keys:
            status = 'Missing'
            matched_registry_dict = {
                'Registry Key/Subkey Path': registry_path,
                'Registry Name': registry_name,
                'Registry Type': registry_data,
                'Data': registry_type
            }

        else:
            # Step 5: Process Pass and Fail statuses by checking current data entries
            for registry in current_pc_reg_data:
                try:
                    # Use the match function to identify Pass or Fail
                    match_status = match_reg_item(registry_path, registry_name, registry_type, registry_data, registry)
                    if match_status == 'Pass':
                        matched_registry.append(registry)  # Found a matching entry with Pass
                    elif match_status == 'Fail':
                        continue  # Skip to the next item if Fail
                except ValueError as ve:
                    print(ve)
                    write_into_error_log(content="Value error: " + str(ve))
                    continue

            # Finalize status based on matches found
            if matched_registry:
                status = 'Pass'
                matched_registry_item = matched_registry[0]  # Use first Pass match found
                matched_registry_dict = {
                    'Registry Key/Subkey Path': matched_registry_item[0],
                    'Registry Name': matched_registry_item[1],
                    'Registry Type': matched_registry_item[3],
                    'Data': matched_registry_item[2]
                }
            else:
                status = 'Fail'
                matched_registry_dict = {
                    'Registry Key/Subkey Path': registry_path,
                    'Registry Name': registry_name,
                    'Registry Type': registry_data,
                    'Data': registry_type
                }

        # Step 6: Add the result to the results dictionary
        results[f"{registry_path}\\{registry_name}"] = {
            'Registry Key/Subkey Path': matched_registry_dict['Registry Key/Subkey Path'],
            'Registry Name': matched_registry_dict['Registry Name'],
            'Registry Type': matched_registry_dict['Registry Type'],
            'Data': matched_registry_dict['Data'],
            'Status': status
        }

    return results


# Generate the redundant keys data show in the bottom table of the summary page
def generate_redundant_data(software_pc_data, current_pc_reg_data):
    results = {}
    # sample registry
    for registry_path, registry_name, registry_type, registry_data in current_pc_reg_data:
        matched_registry = []
        for registry in software_pc_data:
            try:
                match_status = redundant_reg_keys(registry_path, registry_name, registry_type, registry_data, registry)
                if match_status == 'Pass':
                    matched_registry.append(registry)
                elif match_status == 'Fail':
                    continue  # Do nothing, continue checking other items
                elif match_status == 'Missing':
                    continue  # Do nothing, continue checking other items
            except ValueError as ve:
                print(ve)
                write_into_error_log(content="Value error: " + ve)
                continue

        # Determine the status based on whether a match was found
        if matched_registry:
            status = 'Pass'
            matched_registry_item = matched_registry[0]  # Assume first match is used
            matched_registry_dict = {
                'Registry Key/Subkey Path': matched_registry_item[0],
                'Registry Name': matched_registry_item[1],
                'Registry Type': matched_registry_item[3],
                'Data': matched_registry_item[2]
            }
        else:
            if any(redundant_reg_keys(registry_path, registry_name, registry_type, registry_data, registry) == 'Fail'
                   for registry in software_pc_data):
                status = 'Fail'

            elif any(redundant_reg_keys(registry_path, registry_name, registry_type, registry_data, registry) == 'Exist'
                     for registry in software_pc_data):
                status = 'Exist'

            elif any(match_reg_item(registry_path, registry_name, registry_type, registry_data, registry) == 'Not Exist'
                     for registry in current_pc_reg_data):
                status = 'Not Exist'

            else:
                status = 'Missing'

            matched_registry_dict = {
                'Registry Key/Subkey Path': registry_path,
                'Registry Name': registry_name,
                'Registry Type': registry_data,
                'Data': registry_type
            }

        # Add the result to the results dictionary
        results[f"{registry_path}\\{registry_name}"] = {
            'Registry Key/Subkey Path': matched_registry_dict['Registry Key/Subkey Path'],
            'Registry Name': matched_registry_dict['Registry Name'],
            'Registry Type': matched_registry_dict['Registry Type'],
            'Data': matched_registry_dict['Data'],
            'Status': status
        }

    return results


# Compare the registry keys and get the fail and missing keys data only
# The data will be used to show in the upper table in the summary page
def match_reg_compared_item(registry_path, registry_name, registry_type, registry_data, registry):
    try:
        current_registry_path, current_registry_name, current_registry_type, current_registry_data = registry
    except ValueError:
        write_into_error_log(content="Registry item has unexpected structure: " + registry)
        raise ValueError(f"Registry item has unexpected structure: {registry}")

    # Check name similarity
    if current_registry_name == "":
        registry_name = current_registry_name

        # Check if both path and name fully match
    # if registry_path.lower() == current_registry_path.lower() and name_similarity >= 100:
    if current_registry_path.lower() == registry_path.lower() and current_registry_name.lower() == registry_name.lower():
        # Check if either type is 'N/A'
        if registry_type == 'N/A' or current_registry_type == 'N/A':
            return "Pass"
        # special case for password
        if registry_name in ['Password', 'calib3dpassword'] or current_registry_name in ['Password', 'calib3dpassword']:
            return "Exist"
        # Convert data to string for comparison if necessary
        if isinstance(current_registry_data, bytes):
            current_registry_data = " ".join([f"{byte:02X}" for byte in current_registry_data])
        if isinstance(registry_data, bytes):
            registry_data = " ".join([f"{byte:02X}" for byte in registry_data])

        # Compare type and data
        if current_registry_type == registry_type and current_registry_data == registry_data:
            return "Pass"
        else:
            return "Fail"

    elif current_registry_path.lower() != registry_path.lower() and current_registry_name.lower() != registry_name.lower() and registry_name not in [
        'Password', 'calib3dpassword']:

        return "Missing"  # If path or name does not fully match, return missing

    elif current_registry_path.lower() != registry_path.lower() and current_registry_name != registry_name and registry_name in [
        'Password', 'calib3dpassword']:

        return "Not Exist"


# Generate the compared registry data for upper table in summary page
def generate_compared_results(sample_reg_pc_data, current_pc_reg_data):
    results = {}

    # Step 1: Create a set of identifiers for fast lookup
    current_keys = {
        (registry[0].lower(), registry[1].lower()): registry for registry in current_pc_reg_data
    }

    # Step 2: Process each sample registry entry
    for registry_path, registry_name, registry_type, registry_data in sample_reg_pc_data:
        registry_key = (registry_path.lower(), registry_name.lower())

        # Step 3: Check if this key exists in the current data set
        if registry_key in current_keys:
            # Retrieve the matching registry entry from current data
            current_registry = current_keys[registry_key]

            # Determine specific match status with helper function
            status = match_reg_compared_item(
                registry_path, registry_name, registry_type, registry_data, current_registry
            )

            # Build registry dictionary based on identified status
            matched_registry_dict = {
                'Registry Key/Subkey Path': current_registry[0],
                'Registry Name': current_registry[1] or registry_name,
                'Current Type': current_registry[3],
                'Expected Type': registry_data,
                'Current Data': current_registry[2],
                'Expected Data': registry_type
            }

            # Differentiation for `Pass` and `Fail`
            if status == 'Pass':
                status = 'Pass'  # All details match
            elif status == 'Fail':
                status = 'Fail'  # Mismatched type or data

        else:
            # Item is not found in current data; determine if it’s `Missing`, `Exist`, or `Not Exist`
            if registry_name.lower() in ['password', 'calib3dpassword']:
                status = 'Exist'  # Treated as `Exist` if it's a password or similar sensitive key
                matched_registry_dict = {
                    'Registry Key/Subkey Path': registry_path,
                    'Registry Name': registry_name,
                    'Current Type': '-',
                    'Expected Type': registry_data,
                    'Current Data': '-',
                    'Expected Data': registry_type
                }
            else:
                # Default to `Missing` when no match is found in current data
                status = 'Missing'
                matched_registry_dict = {
                    'Registry Key/Subkey Path': registry_path,
                    'Registry Name': registry_name,
                    'Current Type': '-',
                    'Expected Type': registry_data,
                    'Current Data': '-',
                    'Expected Data': registry_type
                }

        # Step 4: Add each processed entry to the results dictionary
        results[f"{registry_path}\\{registry_name}"] = {
            **matched_registry_dict,
            'Status': status
        }

    return results

# Update the read_registry_recursive function to include the registry type
def read_registry_recursive(root_key, path, default_name='Default'):
    result = []
    try:
        with winreg.OpenKey(root_key, path, 0, winreg.KEY_READ) as key:
            index = 0
            while True:
                try:
                    subkey_name = winreg.EnumKey(key, index)
                    subkey_path = os.path.join(path, subkey_name)
                    result.extend(read_registry_recursive(root_key, subkey_path))  # Recursively read subkeys
                    # result.append(subkey_path)
                    result.append(
                        (subkey_path, default_name, '', 'REG_SZ'))  # Add default entry, default type as REG_SZ(string)
                    # result.append((subkey_path, default_name, '', 'Default'))  # Add default entry
                    index += 1
                except OSError:
                    break
            # Read values of the current key
            try:
                value_index = 0
                while True:
                    try:
                        value_name, value_data, value_type = winreg.EnumValue(key, value_index)
                        # debugging part
                        # value_name=value_name.casefold()
                        # print(f"{path},{value_name},{value_data} ({value_type})")
                        if value_type == winreg.REG_DWORD:
                            # Convert DWORD values to hexadecimal strings
                            value_data = f'0x{value_data:08x}'
                        elif value_type == winreg.REG_QWORD:
                            # Convert QWORD values to hexadecimal strings
                            value_data = f'0x{value_data:016x}'
                        elif value_type == winreg.REG_BINARY:
                            value_data = print_reg_binary(value_data)
                            # value_data = f'{value_data}'
                        '''elif reg_type == winreg.REG_SZ or reg_type == winreg.REG_EXPAND_SZ:
                            value_data=value_data
                        elif reg_type == winreg.REG_MULTI_SZ:
                            # Multi-string data should be a list of strings
                            if isinstance(value_data, str):
                                value_data = value_data.split('\0')  # Split the data into a list of strings'''

                        # Map registry type integer value to its string representation
                        registry_type_str = REG_TYPE_MAP.get(value_type, str(value_type))
                        # Include registry type string in the result
                        result.append((path, value_name, value_data, registry_type_str))
                        value_index += 1
                    except OSError:
                        break
            except Exception as e:
                '''print(f"Error reading values: {e}")
                write_into_error_log(f"Error reading values: {e}")'''
                pass
    except Exception as e:
        print(f"Error reading key: {e}")
        write_into_error_log(f"Error reading key: {e}")
    return result


# print reg_binary data
def print_reg_binary(binary_data):
    hex_string = " ".join("{:02x}".format(byte) for byte in binary_data)
    return hex_string


# read the installed registry key in current pc
def read_installed_registry():
    result = []
    try:
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as hkey:
            result = read_registry_recursive(hkey, r'Software\WOW6432Node\MV Technology')
            # result = read_registry_recursive(hkey, r'SOFTWARE\WOW6432Node\MV Technology')
    except Exception as e:
        write_into_error_log(f"Error: {e}")
        print(f"Error: {e}")
    return result


# load the registry data in the json file
def load_registry_from_json(file_path):
    with open(file_path, 'r') as file:
        return json.load(file)


# merge the current existing registry in pc and merge with the data inside the json file that is use to collect the complete data of the registry list
def merged_current_registry_data(output_file):
    existing_registry = read_installed_registry()

    output_file = get_current_file_path(output_file)

    # clear the data in file first so can save the latest data into it
    # save the latest updated data into it
    write_to_json([], output_file)
    updated_reg_data = existing_registry

    write_to_json(updated_reg_data, output_file)
    # save_to_jsonfile(updated_reg_data, output_file)#for compare purpose

    return output_file


'''
def save_to_jsonfile(updated_reg_data, output_file):
    with open(output_file, 'w')as json_file:
        json.dump(updated_reg_data,json_file,indent=4)
        #json_file.write(updated_reg_data.toJson(), indent=4)'''


# backup the registry data into backup folder
def backup_registry(exported_file, backup_folder):
    join_path = os.path.join("data", exported_file)
    src_file_path = get_current_file_path(join_path)
    # Create the backup folder if it doesn't exist
    os.makedirs(backup_folder, exist_ok=True)

    # Generate a unique backup file name using current timestamp
    timestamp = datetime.datetime.now().strftime('%d-%m-%Y_%H%M%S')
    backup_filename = f"{exported_file[:-4]}_{timestamp}.txt"  # Appending timestamp before extension

    # Construct the full path for backup file
    # backup_path = os.path.join(backup_folder, backup_filename)
    backup_path = os.path.join(backup_folder, backup_filename)

    try:
        # Copy the exported file to the backup folder
        shutil.copy(src_file_path, backup_path)
        print(f"Backup created: {backup_path}")
        backup_full_path = get_current_file_path(backup_path)
        write_into_event_log(content="The backup file has been save into: " + backup_full_path)

    except FileNotFoundError:
        print(f"Error: File '{exported_file}' not found.")
        with open(error_log_file_path, 'a') as log_file:
            log_file.write("Error: File '" + exported_file + "' not found.\n")

    except IOError as e:
        print(f"Error copying file: {e}")
        with open(error_log_file_path, 'a') as log_file:
            log_file.write("Error copying file: " + str(e) + "\n")


# backup the registry keys before deletion
'''def backup_deleted_registry(exported_file, backup_folder):
    # Create the backup folder if it doesn't exist

    os.makedirs(backup_folder, exist_ok=True)

    # Generate a unique backup file name using current timestamp
    #timestamp = datetime.datetime.now().strftime('%d-%m-%Y')
    #backup_filename = f"{exported_file[:-5]}_{timestamp}.json"  # Use only the date in the filename
    backup_filename = f"list_of_deleted_keys.json"
    # Construct the full path for the backup file
    backup_path = os.path.join(backup_folder, backup_filename)

    try:
        # Read the content of the exported file
        with open(exported_file, 'r') as file:
            exported_data = json.load(file)

        # Check if the backup file already exists
        if os.path.exists(backup_path):
            # Read existing data from the backup file
            with open(backup_path, 'r') as backup_file:
                backup_data = json.load(backup_file)

            # Append the new data to the existing data
            backup_data.extend(exported_data)

            # Write the combined data back to the backup file
            with open(backup_path, 'w') as backup_file:
                json.dump(backup_data, backup_file, indent=4)

            backup_full_path= get_current_file_path(backup_path)
            print(f"Data appended to backup file: {backup_full_path}")
            write_into_event_log(content="Data appended to backup file: " + backup_full_path)

        else:
            # Copy the exported file to the backup folder as a new file
            shutil.copy(exported_file, backup_path)
            print(f"Backup created: {backup_full_path}")
            write_into_event_log(content="The backup file has been saved into: " + backup_full_path)

    except FileNotFoundError:
        print(f"Error: File '{exported_file}' not found.")
        write_into_error_log(content="Error: File '" + exported_file + "' not found.")

    except IOError as e:
        print(f"Error copying file: {e}")
        write_into_error_log(content="Error copying file: " + str(e))

'''


def backup_deleted_registry(exported_file, backup_folder):
    # Create the backup folder if it doesn't exist
    os.makedirs(backup_folder, exist_ok=True)

    # Generate a unique backup file name
    backup_filename = "list_of_deleted_keys.json"
    backup_path = os.path.join(backup_folder, backup_filename)

    try:
        # Read the content of the exported file
        with open(exported_file, 'r') as file:
            exported_data = json.load(file)

        # Read existing data from the backup file if it exists
        if os.path.exists(backup_path):
            with open(backup_path, 'r') as backup_file:
                backup_data = json.load(backup_file)

            # Convert lists of dictionaries to sets of tuples for comparison
            backup_data_set = {tuple(sorted(item.items())) for item in backup_data}
            exported_data_set = {tuple(sorted(item.items())) for item in exported_data}

            # Find new data that is not already in the backup file
            new_data_set = exported_data_set - backup_data_set
            new_data = [dict(item) for item in new_data_set]

            # Append new data to the backup file
            if new_data:
                backup_data.extend(new_data)

                # Write the combined data back to the backup file
                with open(backup_path, 'w') as backup_file:
                    json.dump(backup_data, backup_file, indent=4)

                backup_full_path = get_current_file_path(backup_path)
                print(f"Data appended to backup file: {backup_full_path}")
                write_into_event_log(content="Data appended to backup file: " + backup_full_path)
            else:
                print("No new data to append.")
                write_into_event_log(content="No new data to append.")

        else:
            # Copy the exported file to the backup folder as a new file
            shutil.copy(exported_file, backup_path)
            print(f"Backup created: {backup_path}")
            write_into_event_log(content="The backup file has been saved into: " + backup_path)

    except FileNotFoundError:
        print(f"Error: File '{exported_file}' not found.")
        write_into_error_log(content="Error: File '" + exported_file + "' not found.")

    except IOError as e:
        print(f"Error copying file: {e}")
        write_into_error_log(content="Error copying file: " + str(e))


def export_file(exported_file):
    # Create the backup folder if it doesn't exist
    # os.makedirs(backup_folder, exist_ok=True)

    # Generate a unique backup file name using current timestamp
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_filename = f"{exported_file[:-5]}_{timestamp}.json"  # Appending timestamp before extension

    # Construct the full path for backup file
    backup_path = os.path.join(export_file_path, backup_filename)

    try:
        # Copy the exported file to the backup folder
        shutil.move(exported_file, backup_path)
        print(f"Backup created: {backup_path}")
        sg.popup_ok(f"The file has been exported to '{export_file_path}\' ", title='Export')
        write_into_event_log(content="The file has been exported to '" + export_file_path)

    except FileNotFoundError:
        print(f"Error: File '{exported_file}' not found.")
        write_into_error_log(content="Error: File  '" + exported_file + "' not found.")

    except IOError as e:
        print(f"Error copying file: {e}")
        write_into_error_log(f"Error copying file: {e}")


def load_registry_from_json2(json_file):
    if os.path.exists(json_file):
        with open(json_file, 'r') as f:
            registry_data = json.load(f)
        return registry_data
    else:
        # Handle the case when the file does not exist
        print(f"File {json_file} not found. Returning empty registry data.")
        return None

'''Golden File Editor Section'''


###############################################################################################################################################
# sorting function
def sort_editor_table(window, event, col_count, current_sort_order):
    if isinstance(event, tuple):
        if event[0] == '-TABLE_EDITOR-':
            # event[2][0] is the row
            # event[2][1] is the column
            if event[2][0] == -1 and event[2][1] != -1:
                col_num_clicked = event[2][1]
                table_values = window["-TABLE_EDITOR-"].get()

                # disable the first column click event
                if col_num_clicked == 0:
                    return
                table_values = window["-TABLE_EDITOR-"].get()

                # Adjust the number of columns based on the provided col_count
                data = [row[:col_count] for row in table_values]

                new_table = sort_order_table(data, col_num_clicked, current_sort_order)

                window["-TABLE_EDITOR-"].update(values=new_table)


# Searching function
def perform_editor_search(search_text, table_data):
    # Filter the table_data to include only rows where the software name contains the search_text
    filtered_data = [row for row in table_data if search_text in row[2].lower()]
    window["-TABLE_EDITOR-"].update(values=filtered_data)


# Initialize current sort order with ascending for each column
current_sort_order_table = [False, False, False, False, False]


# get the data from the selected golden file
def get_json_data(selected_json_file):
    with open(selected_json_file) as f:
        data = json.load(f)
    return data


# append/add new data into the selected json file
def write_new_data_to_json(new_data, file_path):
    # file_path = get_current_file_path(filename)
    try:
        with open(file_path, 'w') as file:
            json.dump(new_data, file, indent=4)
        # print(f"File written successfully to {file_path}")

    except IOError as e:
        print(f"Error writing to file {file_path}: {e}")


# get the new added data
def get_new_data(window):
    # Software\WOW6432Node\MV Technology
    updated_data = []
    new_path = window.find_element("-ADD_PATH-").get()
    new_name = window.find_element("-ADD_NAME-").get()
    new_type = window.find_element("-ADD_TYPE_DROPDOWN-").get()
    decimal_selected = window.find_element("-ADD_FORMAT_DECIMAL-").get()
    hex_selected = window.find_element("-ADD_FORMAT_HEX-").get()

    if new_type == "String":
        updated_type = "REG_SZ"
        add_data = window.find_element("-ADD_DATA-").get()

    elif new_type == "Binary":
        updated_type = "REG_BINARY"
        add_data = window.find_element("-ADD_DATA-").get()

        add_data = ' '.join(format(ord(c), '02x') for c in add_data)

        '''#add_data = bin(int(add_data))
        add_data = add_data.strip().replace(" ", "").upper()  # Clean up input

        try:
            # Convert to bytes
            if len(add_data) % 2 != 0:
                raise ValueError("Hex string must have an even length.")
            binary_data = bytes.fromhex(add_data)

            formatted_data = ' '.join(f'{byte:02X}' for byte in binary_data)
        except ValueError as e:
            print(f"Error in converting hex input to binary: {e}")
            return []

        add_data = formatted_data'''

    elif new_type == "DWORD(32-bit)":
        updated_type = "REG_DWORD_LITTLE_ENDIAN"
        add_data = window.find_element("-ADD_DATA-").get()

        # Ensure edit_data is in hexadecimal format
        if decimal_selected:
            add_data = get_valid_integer_input2(window)
            if add_data is not None:
                add_data = f'0x{int(add_data):08X}'  # Convert from decimal to hexadecimal
            else:
                return []  # Exit if user cancels or closes the window

        elif hex_selected:
            # Ensure the hex string is formatted correctly
            # add_data = f'0x{add_data.upper().lstrip("08X")}'
            add_data = f'0x{add_data.upper().zfill(8)}'

    elif new_type == "QWORD(64-bit)":
        updated_type = "REG_QWORD_LITTLE_ENDIAN"
        add_data = window.find_element("-ADD_DATA-").get()

        # Ensure edit_data is in hexadecimal format
        if decimal_selected:
            add_data = get_valid_integer_input2(window)
            if add_data is not None:
                add_data = f'0x{int(add_data):016X}'  # Convert from decimal to hexadecimal
            else:
                return []  # Exit if user cancels or closes the window
        elif hex_selected:
            # add_data = f'0x{add_data.upper().lstrip("016X")}'
            add_data = f'0x{add_data.upper().zfill(16)}'

    elif new_type == "Multi-String":
        updated_type = "REG_MULTI_SZ"
        add_data = window.find_element("-ADD_DATA-").get().split('\0')  # Split by null characters
        add_data = str(add_data)
        add_data = add_data.split('\0')

    elif new_type == "Expandable String":
        updated_type = "REG_EXPAND_SZ"
        add_data = window.find_element("-ADD_DATA-").get()
        add_data = str(add_data)

    updated_data.append([new_path, new_name, add_data, updated_type])
    print(updated_data)

    return updated_data

def get_valid_integer_input2(window):
    while True:
        add_data = window.find_element("-ADD_DATA-").get()
        if is_decimal(add_data):
            return add_data
        else:
            sg.popup_error("Please enter a valid integer.", title="Error")
            window.find_element("-ADD_DATA-").update('')  # Optionally clear the field
            # window.find_element("-SAVE-").update(disabled=True)  # Disable save button until valid input is entered
            event, _ = window.read()  # Wait for user input
            if event == sg.WIN_CLOSED:
                return None  # Exit if the window is closed


# get the  updated data
def edit_selected_data(selected_edit_data, window):
    edited_data = []
    data = [
        [
            row[0], row[1], row[2], row[3]
        ]
        for row in selected_edit_data
    ]

    for path, name, reg_type, data in data:
        edit_path = window.find_element("-PAGE4_EDIT_PATH-").get()
        edit_name = window.find_element("-PAGE4_EDIT_NAME-").get()
        edit_type = window.find_element("-PAGE4_EDIT_TYPE_DROPDOWN-").get()
        decimal_selected = window.find_element("-PAGE4_FORMAT_DECIMAL-").get()
        hex_selected = window.find_element("-PAGE4_FORMAT_HEX-").get()

        if edit_type == "String":
            updated_edit_type = "REG_SZ"
            edit_data = window.find_element("-PAGE4_EDIT_DATA-").get()

        elif edit_type == "Binary":
            updated_edit_type = "REG_BINARY"
            edit_data = window.find_element("-PAGE4_EDIT_DATA-").get()

            # Convert ASCII characters to hexadecimal representation
            edit_data = ' '.join(format(ord(c), '02x') for c in edit_data)

        elif edit_type == "DWORD(32-bit)":
            updated_edit_type = "REG_DWORD_LITTLE_ENDIAN"
            edit_data = window.find_element("-PAGE4_EDIT_DATA-").get()

            # Ensure edit_data is in hexadecimal format
            if decimal_selected:
                edit_data = get_valid_integer_input3(window)
                if edit_data is not None:
                    edit_data = f'0x{int(edit_data):08X}'  # Convert from decimal to hexadecimal
                else:
                    return []  # Exit if user cancels or closes the window

            elif hex_selected:
                # Ensure the hex string is formatted correctly
                edit_data = f'0x{edit_data.upper().zfill(8)}'

        elif edit_type == "QWORD(64-bit)":
            updated_edit_type = "REG_QWORD_LITTLE_ENDIAN"
            edit_data = window.find_element("-PAGE4_EDIT_DATA-").get()

            # Ensure edit_data is in hexadecimal format
            if decimal_selected:
                edit_data = get_valid_integer_input3(window)
                if edit_data is not None:
                    edit_data = f'0x{int(edit_data):016X}'  # Convert from decimal to hexadecimal
                else:
                    return []  # Exit if user cancels or closes the window
            elif hex_selected:
                edit_data = f'0x{edit_data.upper().zfill(16)}'

        elif edit_type == "Multi-String":
            updated_edit_type = "REG_MULTI_SZ"
            edit_data = window.find_element("-PAGE4_EDIT_DATA-").get().split('\0')  # Split by null characters
            edit_data = str(edit_data)

        elif edit_type == "Expandable String":
            updated_edit_type = "REG_EXPAND_SZ"
            edit_data = window.find_element("-PAGE4_EDIT_DATA-").get()
            data = str(data)

        edited_data.append([edit_path, edit_name, updated_edit_type, edit_data])
        print(f"Edited registry data: {edited_data}")

    return edited_data


def get_valid_integer_input3(window):
    while True:
        edited_data = window.find_element("-PAGE4_EDIT_DATA-").get()
        if is_decimal(edited_data):
            return edited_data
        else:
            sg.popup_error("Please enter a valid integer.", title="Error")
            window.find_element("-PAGE4_EDIT_DATA-").update('')  # Optionally clear the field
            # window.find_element("-SAVE-").update(disabled=True)  # Disable save button until valid input is entered
            event, _ = window.read()  # Wait for user input
            if event == sg.WIN_CLOSED:
                return None  # Exit if the window is closed


def handle_table_selection4(window, event, values):
    selected_list_index = []
    result = []  # Holds the selected row data

    # Incorporate the logic for handling table events
    if isinstance(event, tuple) and event[0] == '-TABLE_EDITOR-':
        row_index = event[2][0]
        column_index = event[2][1]

        # Retrieve the table values
        table_values = window["-TABLE_EDITOR-"].get()

        # Check if the click is on the first column (the checkbox column)
        if column_index == 0 and row_index == -1:
            if not hasattr(handle_table_selection4, 'click_time'):
                handle_table_selection4.click_time = 0

            if handle_table_selection4.click_time == 0:
                handle_table_selection4.click_time = 1
                for row in table_values:
                    if row[0] == BLANK_BOX:
                        row[0] = CHECKED_BOX
            elif handle_table_selection4.click_time == 1:
                handle_table_selection4.click_time = 0
                for row in table_values:
                    if row[0] == CHECKED_BOX:
                        row[0] = BLANK_BOX

            # Update the table with the new checkbox states
            window["-TABLE_EDITOR-"].update(values=table_values)

            # Update the selected rows list
            selected_list_index = [i for i, row in enumerate(table_values) if row[0] == CHECKED_BOX]
            print(f"Selected indices after toggle: {selected_list_index}")

            # Enable or disable buttons based on selection
            if len(selected_list_index) != 1:
                window.find_element("-EDIT_GOLDEN_FILE-").update(disabled=True)
                window.find_element("-DELETE_FROM_GOLDEN_FILE-").update(disabled=False)
            else:
                window.find_element("-EDIT_GOLDEN_FILE-").update(disabled=False)
                window.find_element("-DELETE_FROM_GOLDEN_FILE-").update(disabled=False)

            if not selected_list_index:
                # Clear the selection if no rows are checked
                file_path = 'data/editor/selected_data.json'
                output_file = get_current_file_path(file_path)
                write_to_json([], output_file)

                window.find_element("-DELETE_FROM_GOLDEN_FILE-").update(disabled=True)
            else:
                window.find_element("-DELETE_FROM_GOLDEN_FILE-").update(disabled=False)
                for i in selected_list_index:
                    row = table_values[i]
                    matched_registry_dict = [
                        row[1],  # Registry Key/Subkey Path
                        row[2],  # Registry Name
                        row[4],  # Registry Data
                        row[3]   # Registry Type
                    ]
                    result.append(matched_registry_dict)

        else:
            # Handle row clicks (non-header clicks)
            if isinstance(row_index, int) and row_index >= 0:
                if 0 <= row_index < len(table_values):
                    clicked_row_data = table_values[row_index]
                    if clicked_row_data[0] == BLANK_BOX:
                        # If checkbox is unchecked, check it
                        clicked_row_data[0] = CHECKED_BOX
                    else:
                        # If checkbox is checked, uncheck it
                        clicked_row_data[0] = BLANK_BOX

                    # Update the table with the new checkbox state
                    table_values[row_index] = clicked_row_data
                    window["-TABLE_EDITOR-"].update(values=table_values)

                    # Update the selected rows list
                    selected_list_index = [i for i, row in enumerate(table_values) if row[0] == CHECKED_BOX]
                    print(f"Selected indices: {selected_list_index}")

                    # Enable or disable buttons based on selection
                    if len(selected_list_index) != 1:
                        window.find_element("-EDIT_GOLDEN_FILE-").update(disabled=True)
                        window.find_element("-DELETE_FROM_GOLDEN_FILE-").update(disabled=False)
                    else:
                        window.find_element("-EDIT_GOLDEN_FILE-").update(disabled=False)
                        window.find_element("-DELETE_FROM_GOLDEN_FILE-").update(disabled=False)

                    if not selected_list_index:
                        # Clear the selection if no rows are checked
                        file_path = 'data/editor/selected_data.json'
                        output_file = get_current_file_path(file_path)
                        write_to_json([], output_file)

                        window.find_element("-DELETE_FROM_GOLDEN_FILE-").update(disabled=True)
                    else:
                        window.find_element("-DELETE_FROM_GOLDEN_FILE-").update(disabled=False)
                        for i in selected_list_index:
                            row = table_values[i]
                            matched_registry_dict = [
                                row[1],  # Registry Key/Subkey Path
                                row[2],  # Registry Name
                                row[4],  # Registry Data
                                row[3]   # Registry Type
                            ]
                            result.append(matched_registry_dict)

                        print(f"Result: {result}")

    return result


def delete_selected_data(golden_file, select_file, output_file):
    original_data = load_registry_from_json2(golden_file)  # golden file
    delete_data = load_registry_from_json2(select_file)  # file that stores delete data - data/editor/selected_delete_data.json

    # Convert each entry to a tuple, ensuring any nested lists are converted to tuples
    def convert_to_tuple(entry):
        return tuple(
            tuple(item) if isinstance(item, list) else item
            for item in entry
        )

    # Convert filter data to a set of tuples for easy comparison
    filter_set = set(
        convert_to_tuple(entry)  # Ensure each entry is fully converted to a tuple
        for entry in delete_data
    )

    # Filter out matching entries
    filtered_data = [
        entry for entry in original_data
        if convert_to_tuple(entry) not in filter_set  # Compare fully converted tuples
    ]

    # Write the filtered data back to the JSON file
    write_to_json(filtered_data, output_file)
    sg.popup_ok("The selected data has been deleted.")
    print(f"Filtered data written to {output_file}")
    write_into_event_log(f"The list of deleted registry keys has been updated in {output_file}")

    # Indicate that the deletion was successful
    return True


############################################################################################################################
# Refresh the main page
def refresh_table():
    current_path = get_file_path()
    # file_name= os.path.basename(current_path)
    print(f"Using file path: {current_path}")
    installed_registry = read_installed_registry()
    if installed_registry:

        run_registry_main(current_path)
        update_reg_gui()

    else:
        sg.popup_error("Fail to retrieve registry keys", title="Error")
        write_into_error_log(f"Fail to retrieve registry keys")


# Refresh the main and summary page
def refresh_table2(window, event, values):
    current_path = get_file_path()
    print(f"Using file path: {current_path}")

    installed_registry = read_installed_registry()
    if installed_registry:
        run_registry_main(current_path)
        update_reg_gui()
        run_registry_main2(current_path)
        update_reg_compared_gui(window)
        run_registry_main3(current_path)
        update_redundant_gui(window)
    else:
        sg.popup_error("Fail to retrieve registry keys", title="Error")
        write_into_error_log(f"Fail to retrieve registry keys")

    # For "List of Fail/Missing Registry Keys" search
    if event in ("-SEARCH_BUTTON_REG2-", "\r", "-SEARCH_REG2-"):
        search_text_page4 = values["-SEARCH_REG2-"].strip().lower()
        perform_reg_search2(search_text_page4, window)

        if not search_text_page4:
            update_reg_compared_gui(window)

    # For "List of Redundant Keys" search
    if event in ("-SEARCH_BUTTON_REDUNDANT-", "\r", "-SEARCH_REDUNDANT-"):
        search_text_redundant = values["-SEARCH_REDUNDANT-"].strip().lower()

        # Perform the search for redundant keys
        if search_text_redundant:
            update_redundant_gui(window, search_text_redundant)
        else:
            update_redundant_gui(window)


# Update/Import the registry key into the current pc
def import_registry(registry_keys):
    with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as hkey:
        for key_path, value_name, value_data, registry_type in registry_keys:
            full_key_path = rf'{key_path}'
            try:
                with winreg.CreateKey(hkey, full_key_path) as new_key:
                    reg_type = None

                    # Determine the registry type based on REG_TYPE_MAP
                    for reg_key, reg_str in REG_TYPE_MAP.items():
                        if reg_str == registry_type:
                            reg_type = reg_key
                            break

                    # Special case: registry name contains "password" will be bypassed
                    if value_name.lower() in ['password', 'calib3dpassword']:
                        print(f"Password will not be replaced: '{registry_type}' for key '{key_path}\\{value_name}'")
                        write_into_event_log(
                            content="Password will not be replaced for key '" + key_path + "\\" + value_name)
                        continue

                    # If registry type is 'Default', set as REG_BINARY
                    if registry_type == "Default":
                        reg_type = winreg.REG_BINARY
                        value_data = value_data.encode()  # Convert value data to bytes

                    if reg_type is None:
                        print(f"Unsupported registry type '{registry_type}' for key '{key_path}\\{value_name}'")
                        write_into_event_log(
                            content="Unsupported registry type '" + registry_type + "' for key '" + key_path + "\\" + value_name + "\n")
                        continue

                    # Handle different registry types
                    if reg_type == winreg.REG_DWORD:
                        # Convert hexadecimal string to integer
                        value_data = int(value_data, 16)
                        winreg.SetValueEx(new_key, value_name, 0, reg_type, value_data)
                    elif reg_type == winreg.REG_SZ or reg_type == winreg.REG_EXPAND_SZ:
                        winreg.SetValueEx(new_key, value_name, 0, reg_type, str(value_data))
                    elif reg_type == winreg.REG_BINARY:
                        # Ensure value_data is in bytes
                        if isinstance(value_data, str):
                            try:
                                value_data = bytes.fromhex(value_data.replace(" ", ""))  # Convert hex string to bytes
                            except ValueError:
                                print(f"Failed to convert value data '{value_data}' to bytes")
                                write_into_error_log(f"Failed to convert value data '{value_data}' to bytes")
                                continue
                        winreg.SetValueEx(new_key, value_name, 0, reg_type, value_data)
                    elif reg_type == winreg.REG_MULTI_SZ:
                        # Multi-string data should be a list of strings
                        if isinstance(value_data, str):
                            value_data = value_data.split('\0')  # Split the data into a list of strings
                        winreg.SetValueEx(new_key, value_name, 0, reg_type, value_data)
                    elif reg_type == winreg.REG_QWORD:
                        # Convert hexadecimal string to integer for QWORD
                        value_data = int(value_data, 16)
                        winreg.SetValueEx(new_key, value_name, 0, reg_type, value_data)
                    else:
                        print(f"Unsupported registry type '{registry_type}' for key '{key_path}\\{value_name}'")
                        write_into_event_log(
                            content="Unsupported registry type '" + registry_type + "' for key '" + key_path + "\\" + value_name)

                print(f"Imported key '{key_path}\\{value_name}' successfully.")
                # write_into_log(content = "Imported key  '"+ key_path + "\\" + value_name + " successfully."+"\n")
            except OSError as e:
                print(f"Error creating key '{full_key_path}\\{value_name}': {e}")
                write_into_error_log(f"Error creating key '{full_key_path}\\{value_name}': {e}")


# Import the selected registry keys only
def import_selected_registry_result(registry_keys):
    with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as hkey:
        for key in registry_keys:
            key_path = key['Registry Key/Subkey Path']
            value_name = key['Registry Name']
            value_data = key['Expected Data']
            registry_type = key['Expected Type']  # sample type
            current_type = key['Current Type']
            current_data = key['Current Data']

            try:
                # Create or open the registry key
                with winreg.CreateKey(hkey, key_path) as new_key:
                    reg_type = None

                    # Determine the registry type based on REG_TYPE_MAP
                    for reg_key, reg_str in REG_TYPE_MAP.items():
                        if reg_str == registry_type:
                            reg_type = reg_key
                            break

                    # If registry type is 'Default', set as REG_BINARY
                    if registry_type == "Default":
                        reg_type = winreg.REG_BINARY
                        value_data = value_data.encode()  # Convert value data to bytes

                    if reg_type is None:
                        print(f"Unsupported registry type '{registry_type}' for key '{key_path}\\{value_name}'")
                        write_into_event_log(
                            content="Unsupported registry type '" + registry_type + "' for key '" + key_path + "\\" + value_name)
                        continue

                    # Handle different registry types
                    if reg_type == winreg.REG_DWORD:
                        # Convert hexadecimal string to integer
                        value_data = int(value_data, 16)
                        winreg.SetValueEx(new_key, value_name, 0, reg_type, value_data)
                        current_data = f'0x{value_data:08x}'
                    elif reg_type == winreg.REG_SZ or reg_type == winreg.REG_EXPAND_SZ:
                        winreg.SetValueEx(new_key, value_name, 0, reg_type, str(value_data))
                    elif reg_type == winreg.REG_BINARY:
                        # Ensure value_data is in bytes
                        if isinstance(value_data, str):
                            value_data = bytes.fromhex(value_data.replace(" ", ""))  # Convert the hex string to bytes
                        winreg.SetValueEx(new_key, value_name, 0, reg_type, value_data)

                        # Ensure current_data is also treated as bytes for comparison and use
                        if isinstance(current_data, str) and current_data != "-":
                            try:
                                current_data = bytes.fromhex(current_data.replace(" ", ""))
                                current_data = print_reg_binary(current_data)
                            except ValueError:
                                print(f"Failed to convert current data '{current_data}' to bytes")
                    elif reg_type == winreg.REG_MULTI_SZ:
                        # Multi-string data should be a list of strings
                        if isinstance(value_data, str):
                            value_data = value_data.split('\0')  # Split the data into a list of strings
                        winreg.SetValueEx(new_key, value_name, 0, reg_type, value_data)
                    elif reg_type == winreg.REG_QWORD:
                        # Convert hexadecimal string to integer for QWORD
                        value_data = int(value_data, 16)
                        winreg.SetValueEx(new_key, value_name, 0, reg_type, value_data)
                        current_data = f'0x{value_data:016x}'
                    else:
                        print(f"Unsupported registry type '{registry_type}' for key '{key_path}\\{value_name}'")
                        write_into_event_log(
                            content="Unsupported registry type '" + registry_type + "' for key '" + key_path + "\\" + value_name)

                print(f"Imported key '{key_path}\\{value_name}' successfully.")
                write_into_event_log(
                    content="Imported key '" + key_path + "\\" + value_name + " successfully.\n\tRegistry key type has been changed from  " + current_type + " to " + registry_type + "\n\tRegistry key data has been changed from " + str(
                        current_data) + " to " + str(value_data))
            except OSError as e:
                print(f"Error creating key '{key_path}\\{value_name}': {e}")

def restore_selected_registry_results_page4(registry_keys):
    with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as hkey:
        for key in registry_keys:
            key_path = key['Registry Key/Subkey Path']  # Correct field
            value_data = key['Data']  # Correct field
            registry_type = key['Type']  # Correct field

            try:
                # Create or open the registry key
                with winreg.CreateKey(hkey, key_path) as new_key:
                    reg_type = None

                    # Determine the registry type based on REG_TYPE_MAP
                    for reg_key, reg_str in REG_TYPE_MAP.items():
                        if reg_str == registry_type:
                            reg_type = reg_key
                            break

                    # If registry type is missing, log an error
                    if reg_type is None:
                        write_into_error_log(f"Unknown registry type for {key_path}: {registry_type}")
                        continue

                    # Write the value to the registry key
                    winreg.SetValueEx(new_key, "", 0, reg_type, value_data)
                    write_into_event_log(f"Restored {key_path} with value {value_data} and type {registry_type}")

            except OSError as e:
                write_into_error_log(f"Failed to restore {key_path}: {str(e)}")
                print(f"Failed to restore {key_path}: {str(e)}")


# import back / restore the deleted keys into the current pc
def restore_selected_registry_result(registry_keys):
    with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as hkey:
        for key in registry_keys:
            key_path = key['Registry Key/Subkey Path']
            value_name = key['Registry Name']
            value_data = key['Data']
            registry_type = key['Type']  # sample type

            try:
                # Create or open the registry key
                with winreg.CreateKey(hkey, key_path) as new_key:
                    reg_type = None

                    # Determine the registry type based on REG_TYPE_MAP
                    for reg_key, reg_str in REG_TYPE_MAP.items():
                        if reg_str == registry_type:
                            reg_type = reg_key
                            break
                    # If registry type is 'Default', set as REG_BINARY
                    if registry_type == "Default":
                        reg_type = winreg.REG_BINARY
                        value_data = value_data.encode()  # Convert value data to bytes

                    if reg_type is None:
                        print(f"Unsupported registry type '{registry_type}' for key '{key_path}\\{value_name}'")
                        write_into_event_log(
                            content="Unsupported registry type '" + registry_type + "' for key '" + key_path + "\\" + value_name)
                        continue

                    # Handle different registry types
                    if reg_type == winreg.REG_DWORD:
                        # Convert hexadecimal string to integer
                        value_data = int(value_data, 16)
                        winreg.SetValueEx(new_key, value_name, 0, reg_type, value_data)

                    elif reg_type == winreg.REG_SZ or reg_type == winreg.REG_EXPAND_SZ:
                        winreg.SetValueEx(new_key, value_name, 0, reg_type, str(value_data))
                    elif reg_type == winreg.REG_BINARY:
                        # Ensure value_data is in bytes
                        if isinstance(value_data, str):
                            value_data = value_data.encode()  # Convert to bytes if it's a string
                        winreg.SetValueEx(new_key, value_name, 0, reg_type, value_data)

                    elif reg_type == winreg.REG_MULTI_SZ:
                        # Multi-string data should be a list of strings
                        if isinstance(value_data, str):
                            value_data = value_data.split('\0')  # Split the data into a list of strings
                        winreg.SetValueEx(new_key, value_name, 0, reg_type, value_data)
                    elif reg_type == winreg.REG_QWORD:
                        # Convert hexadecimal string to integer for QWORD
                        value_data = int(value_data, 16)
                        winreg.SetValueEx(new_key, value_name, 0, reg_type, value_data)

                    else:
                        print(f"Unsupported registry type '{registry_type}' for key '{key_path}\\{value_name}'")
                        write_into_event_log(
                            content="Unsupported registry type '" + registry_type + "' for key '" + key_path + "\\" + value_name)

                print(f"Imported key '{key_path}\\{value_name}' successfully.")
                write_into_event_log(content="Imported key '" + key_path + "\\" + value_name + " successfully.")
            except OSError as e:
                print(f"Error creating key '{key_path}\\{value_name}': {e}")


# Compare the restore keys data and all the key in "List_of_deleted_keys",
# update the list of deleted keys after the restore keys has  been restore
# using filter function
def filter_out_matching_entries(source_file, filter_file, output_file):
    """
    Removes entries from source_file that are present in filter_file.
    """
    source_data = load_registry_from_json2(source_file)  # list_of_deleted_keys.json
    filter_data = load_registry_from_json2(filter_file)  # selected_restore_data

    # Convert filter data to a set of tuples for easy comparison
    filter_set = set(
        (entry["Registry Key/Subkey Path"], entry["Registry Name"], entry["Type"], entry["Data"])
        for entry in filter_data
    )

    # Filter out matching entries
    filtered_data = [
        entry for entry in source_data
        if (entry["Registry Key/Subkey Path"], entry["Registry Name"], entry["Type"], entry["Data"]) not in filter_set
    ]

    # Write the filtered data back to the JSON file
    write_to_json(filtered_data, output_file)
    print(f"Filtered data written to {output_file}")
    write_into_event_log(f"The list of deleted registry keys has been updated in {output_file}")


# compare the registry key
# for fail and missing keys
def compare_registries(current_pc_reg_data, software_pc_data):
    failed_missing_keys = []
    for sample_entry in software_pc_data:
        registry_path, registry_name, registry_type, registry_data = sample_entry
        match_found = False

        for current_registry in current_pc_reg_data:
            current_registry_path, current_registry_name, current_registry_type, current_registry_data = current_registry

            if (registry_path.lower() == current_registry_path.lower()) and (
                    registry_name.lower() == current_registry_name.lower()):
                match_found = True  # Path and name match
                if (registry_type != current_registry_type) or (registry_data != current_registry_data):
                    # Type or data does not match
                    # failed keys
                    failed_missing_keys.append((registry_path, registry_name, registry_type, registry_data))
                    # for print the failed_missing_key into the log
                    write_into_event_log(
                        content="Imported failed key '" + registry_path + "\\" + registry_name + " successfully.\n\tRegistry key data has been changed from " + str(
                            current_registry_type) + " to " + str(registry_type))

                break  # No need to check further if a match is found
        # missing keys
        # If no match was found for the current sample_entry
        if not match_found:
            # Optionally, you could choose to append or handle cases where path and name do not match
            failed_missing_keys.append((registry_path, registry_name, registry_type, registry_data))
            write_into_event_log(
                content="Imported missing key '" + registry_path + "\\" + registry_name + " successfully.")

            # pass  # Do nothing if you only care about mismatches in type/data

    print(f"Fail and missing keys: {failed_missing_keys}")
    return failed_missing_keys


# compare and get the fail registry keys only
def compare_fail_registries(current_pc_reg_data, software_pc_data):
    fail_keys = []

    for sample_entry in software_pc_data:
        registry_path, registry_name, registry_type, registry_data = sample_entry
        match_found = False

        for current_registry in current_pc_reg_data:
            current_registry_path, current_registry_name, current_registry_type, current_registry_data = current_registry

            if (registry_path.lower() == current_registry_path.lower()) and (
                    registry_name.lower() == current_registry_name.lower()):
                match_found = True  # Path and name match
                if registry_name in ['Password', 'calib3dpassword'] or current_registry_name in ['Password',
                                                                                                 'calib3dpassword']:
                    match_found = True
                if ((registry_type != current_registry_type) or (registry_data != current_registry_data)) and (
                        registry_name not in ['Password', 'calib3dpassword'] or current_registry_name not in [
                    'Password', 'calib3dpassword']):
                    # Type or data does not match
                    fail_keys.append((registry_path, registry_name, registry_type, registry_data))
                    # print into log file
                    write_into_event_log(
                        content="Imported failed key '" + registry_path + "\\" + registry_name + " successfully.\n\tRegistry key data has been changed from " + str(
                            current_registry_type) + " to " + str(registry_type))

                break  # No need to check further if a match is found

        # If no match was found for the current sample_entry
        if not match_found:
            # Optionally, you could choose to append or handle cases where path and name do not match
            # fail_keys.append((registry_path, registry_name, registry_type, registry_data))
            pass  # Do nothing if you only care about mismatches in type/data

    print(f"Fail keys: {fail_keys}")
    return fail_keys


# compare and get the missing keys only
def compare_missing_registries(current_pc_reg_data, software_pc_data):
    missing_keys = []

    for sample_entry in software_pc_data:
        registry_path, registry_name, registry_type, registry_data = sample_entry
        match_found = False

        for current_registry in current_pc_reg_data:
            current_registry_path, current_registry_name, current_registry_type, current_registry_data = current_registry

            if (registry_path.lower() == current_registry_path.lower()) and (
                    registry_name.lower() == current_registry_name.lower()):
                match_found = True
                break

        if not match_found:
            missing_keys.append((registry_path, registry_name, registry_type, registry_data))
            write_into_event_log(
                content="Imported missing key '" + registry_path + "\\" + registry_name + " successfully.")

    print(f"Missing keys: {missing_keys}")
    return missing_keys


#########################################################################################################################
# functions for page 4

def update_editor_gui(file_path):
    editor_data = []  # Initialize as an empty list
    try:
        json_file_path = get_current_file_path(file_path)
        os.makedirs(os.path.dirname(json_file_path), exist_ok=True)

        if not os.path.isfile(json_file_path):
            write_into_error_log(content="The file '" + json_file_path + "' does not exist.")
            raise FileNotFoundError(f"The file {json_file_path} does not exist.")

        if os.path.getsize(json_file_path) == 0:
            write_into_error_log(content="The file '" + json_file_path + "' is empty.")
            raise ValueError(f"The file {json_file_path} is empty.")

        with open(json_file_path, 'r') as json_file:
            try:
                base_data = json.load(json_file)
            except json.JSONDecodeError:
                write_into_error_log(content="The file '" + json_file_path + "' contains invalid JSON data.")
                raise ValueError(f"The file {json_file_path} contains invalid JSON data.")

        golden_file_name = os.path.basename(file_path).replace(".json", "")
        edit_temp_file_path = os.path.join("data", f"{golden_file_name}_edit.temp")

        if os.path.exists(edit_temp_file_path) and os.path.getsize(edit_temp_file_path) > 0:
            with open(edit_temp_file_path, 'r') as edit_temp_file:
                changes = json.load(edit_temp_file)
        else:
            changes = []

        for change in changes:
            action = change.get("Action")
            key_path = change.get("Registry Key/Subkey Path")
            name = change.get("Registry Name")

            if action == "Add":
                if not any(entry[0] == key_path and entry[1] == name for entry in base_data):
                    base_data.append([key_path, name, change['Data'], change['Type']])
            elif action == "Edit":
                for entry in base_data:
                    if entry[0] == key_path and entry[1] == name:
                        entry[2] = change['Current Data']
                        entry[3] = change['Current Type']

            elif action == "Delete":
                base_data = [entry for entry in base_data if not (entry[0] == key_path and entry[1] == name)]

        editor_data = [
            [row[0], row[1], row[2], row[3]]
            for row in base_data
        ]

        displayed_editor_data = [
            [BLANK_BOX, row[0], row[1], row[3], row[2]]
            for row in base_data
        ]

        window["-TABLE_EDITOR-"].update(values=displayed_editor_data)

        # Disable the delete button if the table is empty
        if len(displayed_editor_data) == 0:
            window['-DELETE_FROM_GOLDEN_FILE-'].update(disabled=True)

    except Exception as e:
        sg.popup_error(f"An error occurred while updating the table: {e}", title="Error")
        write_into_error_log(content="An error occurred while updating the table: " + str(e))

    return editor_data

def write_restored_data_to_json(updated_data, file_path):
    """
    Write the restored data back to the JSON file.
    """
    try:
        with open(file_path, 'w') as file:
            json.dump(updated_data, file, indent=4)
        print(f"Restored data successfully written to {file_path}")
    except IOError as e:
        print(f"Error writing restored data to {file_path}: {e}")
        write_into_error_log(content=f"Error writing restored data to {file_path}: {e}")


def restore_registry_key(registry_path, registry_name, registry_type, registry_data):
    try:
        # Connect to the registry
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as hkey:
            with winreg.OpenKey(hkey, registry_path, 0, winreg.KEY_ALL_ACCESS) as reg_key:
                # Restore the registry value based on type
                if registry_type == 'REG_SZ':
                    winreg.SetValueEx(reg_key, registry_name, 0, winreg.REG_SZ, registry_data)
                elif registry_type == 'REG_DWORD':
                    winreg.SetValueEx(reg_key, registry_name, 0, winreg.REG_DWORD, int(registry_data, 16))

        print(f"Successfully restored {registry_name} in {registry_path}")
        write_into_event_log(f"Restored {registry_name} in {registry_path}")

    except Exception as e:
        print(f"Error restoring registry key: {e}")
        write_into_error_log(f"Error restoring {registry_name} in {registry_path}: {e}")

'''
def perform_restore(selected_rows, window_restore_page4):
    try:
        # Get the full table values from the window
        table_values = window_restore_page4['-TABLE_RESTORE_PAGE4-'].get()

        # Iterate through the selected rows
        restored_data = []
        for row_index in selected_rows:
            # Fetch the row data using the row index
            row = table_values[row_index]

            # Ensure you're fetching the correct columns in the right order:
            registry_key = row[1]  # Registry Key/Subkey Path
            registry_name = row[2]  # Registry Name
            registry_type = row[3]  # Registry Type
            registry_data = row[4]  # Registry Data

            # Validate registry type and data before restoring
            valid_registry_types = ['REG_SZ', 'REG_BINARY', 'REG_DWORD', 'REG_QWORD', 'REG_MULTI_SZ', 'REG_EXPAND_SZ']
            if registry_type not in valid_registry_types or registry_type == '':
                sg.popup_error(f"Error restoring key '{registry_key}': Unsupported or missing registry type: {registry_type}")
                continue

            try:
                # Perform restore
                restore_registry_key(registry_key, registry_name, registry_type, registry_data)
                # Collect restored data
                restored_data.append([registry_key, registry_name, registry_data, registry_type])
            except Exception as e:
                sg.popup_error(f"Error restoring key '{registry_key}': {str(e)}")
                continue

        # Write restored data back to the Golden File (sample-MicroAOI.json)
        golden_file = "C:/Users/zi-kai.soon/Documents/zk/Source code/Golden File/sample-MicroAOI.json"
        write_to_json_for_page4(restored_data, golden_file)

        # Refresh the table after restoration to show the restored data
        update_editor_gui(file_path=golden_file)

        sg.popup_auto_close('Successfully restored selected keys.')
        window_restore_page4['-restoreSelectedPage4-'].update(disabled=True)
        window_restore_page4.close()

    except Exception as e:
        sg.popup_error(f"An error occurred during restore: {e}")
'''

####################################################################################################################################
def compare_sizes(selected_folder, item_sizes):
    result_data = []
    missing_apps = set(item_sizes.keys())

    for root, dirs, files in os.walk(selected_folder):
        for file in files:  # for key values (file) in files)
            file_path = os.path.join(root, file)  # joining 'os.path' and file
            if os.path.basename(file_path) in item_sizes:
                size = os.path.getsize(file_path)  # get the current size of the current file in system
                expected_size = item_sizes[
                    os.path.basename(file_path)]  # Reference value # size we want listed in json file
                missing_apps.discard(os.path.basename(file_path))
                result_data.append((os.path.basename(file_path), expected_size, size,
                                    "Pass" if int(size) == int(expected_size) else "Failed"))

    for missing_app in missing_apps:
        result_data.append((missing_app, item_sizes.get(missing_app, "N/A"), "N/A", "Missing"))

    return result_data


# Searching function for Software checker in page 1
# Function to handle the search logic
def perform_search(search_text, table_data):
    # Filter the table_data to include only rows where the software name contains the search_text
    filtered_data = [row for row in table_data if search_text in row[1].lower()]

    # Update the table with the filtered data for Page 1
    # window["-TABLE-"].update(values=filtered_data)
    update_table(window, filtered_data)


# Searching function for size checker in  page 2
def perform_size_search(search_text, table_data):
    # Assuming you want to search in the "Application" column (index 1)
    filtered_data = [row for row in table_data if search_text in str(row[1]).lower()]

    # Update the table with the filtered data for Page 2
    update_size_table(window, filtered_data)


# Searching function for Registry Checker in Page 3
def perform_reg_search(search_text, table_data):
    # Filter the table_data to include only rows where the software name contains the search_text
    filtered_data = [row for row in table_data if search_text in row[2].lower()]

    for i, row in enumerate(filtered_data):
        row[0] = str(i + 1)
    new_color = set_status_color(filtered_data)
    window["-TABLE_REG-"].update(values=filtered_data, row_colors=new_color)


def perform_reg_search2(search_text, window, event, values):
    table_data = window["-TABLE_REG_COMPARED-"].get()
    # Filter the table_data to include only rows where the software name contains the search_text
    filtered_data = [row for row in table_data if search_text in row[2].lower()]
    handle_table_selection(window, event, values)
    window["-TABLE_REG_COMPARED-"].update(values=filtered_data)


def perform_reg_search3(search_text, window, event, values):
    table_data = window["-TABLE_RESTORE-"].get()
    # Filter the table_data to include only rows where the software name contains the search_text
    filtered_data = [row for row in table_data if search_text in row[2].lower()]
    handle_table_selection3(window, event, values)
    window["-TABLE_RESTORE-"].update(values=filtered_data)


# Function to create the layout for Page 2: Software Checker
def create_page1_layout():
    layout = [
        [
            sg.Button("Check", key="-CHECK-"),
        ],
        [
            sg.InputText(key="-SEARCH-", size=(20, 1), do_not_clear=True, enable_events=True),
            sg.Button("Search", key="-SEARCH_BUTTON-"),
        ],
        [
            sg.Table(
                values=[],
                headings=["List", "Software", "Required Version", "Installed Version", "Status"],
                auto_size_columns=False,
                justification="left",
                num_rows=20,
                key="-TABLE-",
                col_widths=[3, 40, 20, 20, 20],
                background_color="#045D5D",
                text_color="white",
                bind_return_key=True,
                row_colors=[("white", "#045D5D"), ("white", "yellow"), ("white", "red")],
                enable_events=True,
                enable_click_events=True,
                expand_x=True,
                expand_y=True
            ),
        ],
        [
            sg.Column(layout=[
                [sg.Frame('Software Found Info', font=("Helvetica", 12, "bold"), layout=[
                    [sg.Text("Required Software:"),sg.Push(),  sg.Text("0", font=("Helvetica", 12, "bold"), key="-REQUIRED-")],
                    [sg.Text("Matched Version Software:"), sg.Push(), sg.Text("0", font=("Helvetica", 12, "bold"), key="-MATCHED_VERSION-")],
                    [sg.Text("Software Found:"), sg.Push(), sg.Text("0", font=("Helvetica", 12, "bold"), key="-FOUND-")],
                    [sg.Text("Software Not Found:"), sg.Push(), sg.Text("0", font=("Helvetica", 12, "bold"), key="-NOT_FOUND-")],
                ], element_justification="left", size=(400, 250), expand_x=True, expand_y=True)],
            ], expand_y=True),

            sg.Column(layout=[
                [sg.Frame('List of Missing/Failed Status', font=("Helvetica", 12, "bold"), layout=[
                    [sg.Multiline("", key="-MISSING_FAILED-", size=(80, 10), disabled=True, background_color="#045D5D",
                                  text_color="white", expand_x=True, expand_y=True)],  # Enable multiline expansion
                ], element_justification="left", size=(1505, 250), expand_x=True, expand_y=True)],  # Enable frame expansion
            ], expand_x=True, expand_y=True),  # Enable column expansion
        ],
    ]

    return layout


def create_page2_layout():
    # Determine the display path based on the availability of 'Tools' and 'GNU' folders
    paths_to_check = ["C:\\Tools\\GNU", "C:\\Tools"]

    for path in paths_to_check:
        if os.path.exists(path):
            display_path = path
            break
    else:
        display_path = "C"

    layout = [
        [
            sg.Text("GNU Folder Directory Path:"),
            sg.InputText(default_text=display_path, key="-FOLDER-", readonly=True, text_color="#045D5D",
                         background_color="White", disabled_readonly_background_color="White"),
            sg.Button("Check", key="-CHECK_GNU_PATH-"),
            sg.Button("Browse", key="-CHECK_SIZE-")
        ],
        [
            sg.InputText(key="-SEARCH_SIZE-", size=(20, 1), do_not_clear=True, enable_events=True),
            sg.Button("Search", key="-SEARCH_BUTTON_SIZE-"),
        ],
        [
            sg.Table(
                values=[],  # Data will be updated when comparing
                headings=["List", "Application", "Expected Size (Bytes)", "Retrieved Size (Bytes)", "Status"],
                auto_size_columns=False,
                justification="left",
                num_rows=20,
                key="-SIZE_TABLE-",
                col_widths=[3, 60, 20, 20, 20],
                background_color="#045D5D",
                text_color="white",
                bind_return_key=True,
                row_colors=[("white", "#045D5D"), ("white", "yellow"), ("white", "red")],
                enable_events=True,
                enable_click_events=True,
                expand_x=True,  # Allow horizontal expansion
                expand_y=True   # Allow vertical expansion
            ),
        ],
        [
            sg.Column(layout=[
                [sg.Frame('Size Found Info', font=("Helvetica", 12, "bold"), layout=[
                    [sg.Text("Number of Required Sizes:"), sg.Push(), sg.Text("98", font=("Helvetica", 12, "bold"), key="-REQUIRED_SIZE-")],
                    [sg.Text("Number of Matched Sizes Found:"), sg.Push(), sg.Text("0", font=("Helvetica", 12, "bold"), key="-MATCHED_SIZE-")],
                    [sg.Text("Number of Sizes Found:"), sg.Push(), sg.Text("0", font=("Helvetica", 12, "bold"), key="-FOUND_SIZE-")],
                    [sg.Text("Number of Sizes Not Found:"), sg.Push(), sg.Text("0", font=("Helvetica", 12, "bold"), key="-NOT_FOUND_SIZE-")],
                ], element_justification="left", size=(400, 250), expand_x=True, expand_y=True)],  # Match layout size with Page 1
            ], expand_y=True),

            sg.Column(layout=[
                [sg.Frame('List of Incorrect Sizes Applications', font=("Helvetica", 12, "bold"), layout=[
                    [sg.Multiline("", key="-INCORRECT_SIZES-", size=(80, 10), disabled=True, background_color="#045D5D",
                                  text_color="white", expand_x=True, expand_y=True)],  # Enable multiline expansion
                ], element_justification="left", size=(1505, 250), expand_x=True, expand_y=True)],  # Match layout size with Page 1
            ], expand_x=True, expand_y=True),
        ],
    ]

    return layout


# page 3 layout
selected_file_output = sg.Text("")

# select the machine type ( to select the golden file based on the type)
dropdown_values = ["MicroAOI", "Semicon", "SideCam", "SMT", "Other"]


def create_page3_layout(table_data, new_color):
    # Original page layout content, now inside a scrollable Column
    inner_layout = [
        [sg.Text('Select a machine type:'),
         sg.Combo(dropdown_values, default_value='MicroAOI', key="Dropdown", readonly=True),
         sg.Button("Browse"),
         sg.Push(),
         sg.Button('Export')],

        [
            sg.Text("Search:", key="-SEARCH_TEXT-", visible=True),
            sg.InputText(key="-SEARCH_REG-", size=(20, 1), do_not_clear=True, enable_events=True, visible=True),
        ],

        [
            sg.pin(sg.Button("Compare Registry", key='_compare_', visible=False)),
            sg.Push()
        ],

        [
            sg.Table(
                values=table_data,
                headings=["List", "Registry Key/Subkey Path", "Registry Name", "Type", "Data", "Status"],
                auto_size_columns=False,
                vertical_scroll_only=False,
                justification="left",
                num_rows=30,
                key="-TABLE_REG-",
                col_widths=[8, 60, 30, 40, 50, 20],
                background_color="#045D5D",
                text_color="white",
                bind_return_key=True,
                enable_events=True,
                enable_click_events=True,
                row_colors=new_color,
                tooltip=None,
                expand_x=False,
                expand_y=True
            ),
        ],

        [sg.Column(layout=[
            [sg.Push(), sg.Button("View Details", key="ViewSummary", visible=False, size=(12, 1))]
        ], justification="right")],

        # New layout for counters arranged horizontally
        [
            sg.Frame('Registry Found Info', font=("Helvetica", 12, "bold"), layout=[
                [
                    sg.Frame('', layout=[
                        [sg.Text("Total Required Registry Keys:"), sg.Push(),
                         sg.Text("0", key="-TOTAL_REQUIRED_REG-", font=("Helvetica", 12, "bold"))],
                        [sg.Text('●', text_color='yellow'), sg.Text("Missing Registry Keys:"), sg.Push(),
                         sg.Text("0", key="-MISSING_REG-", font=("Helvetica", 12, "bold"))],
                        [sg.Text('●', text_color='orange'), sg.Text("Not Exist Registry Keys:"), sg.Push(),
                         sg.Text("0", key="-NOT_EXIST_REG-", font=("Helvetica", 12, "bold"))]
                    ], element_justification="left", size=(500, 150)),

                    sg.Frame('', layout=[
                        [sg.Text("Total Registry Keys Found:"), sg.Push(),
                         sg.Text("0", key="-TOTAL_FOUND_REG-", font=("Helvetica", 12, "bold"))],
                        [sg.Text('●', text_color='red'), sg.Text("Failed Registry Keys:"), sg.Push(),
                         sg.Text("0", key="-FAIL_REG-", font=("Helvetica", 12, "bold"))],
                        [sg.Text('●', text_color='purple'), sg.Text("Exist Registry Keys:"), sg.Push(),
                         sg.Text("0", key="-EXIST_REG-", font=("Helvetica", 12, "bold"))],
                        [sg.Text('●', text_color='green'), sg.Text("Matched Registry Keys:"), sg.Push(),
                         sg.Text("0", key="-MATCHED_REG-", font=("Helvetica", 12, "bold"))]
                    ], element_justification="left", size=(500, 150))
                ]
            ], element_justification="center")
        ]
    ]

    # Wrap the entire inner layout in a scrollable Column
    layout = [
        [sg.Column(inner_layout, size=(2000, 1200), scrollable=True, vertical_scroll_only=False)]
    ]

    return layout


def create_page4_layout():
    # Define dropdown values without "Other"
    dropdown_values_filtered = ["MicroAOI", "Semicon", "SideCam", "SMT"]

    layout = [
        [sg.Text('Select a machine type:'),
         sg.Combo(dropdown_values_filtered, default_value='MicroAOI', key="Dropdown2", readonly=True),
         sg.Button("Browse", key="-BROWSE2-"),
         sg.Push(),  # Push Browse to the right
         sg.Button("Restore", key="-RESTORE_REV-", disabled=True)],

        [
            sg.pin(sg.Text("Search:")),
            sg.pin(
                sg.InputText(key="-SEARCH_EDITOR-", size=(20, 1), do_not_clear=True, enable_events=True, visible=True)),
            sg.Push(),
            sg.pin(sg.Button("Add", key="-ADD_NEW-", visible=False)),
            sg.pin(sg.Button("Edit", key="-EDIT_GOLDEN_FILE-", visible=False, disabled=True)),
            sg.pin(sg.Button("Delete", key="-DELETE_FROM_GOLDEN_FILE-", visible=False, disabled=True)),
            sg.pin(sg.Button("Save", key="-SAVE_GOLDEN_FILE-", visible=False, disabled=True))
        ],
        [
            sg.Table(
                values=[],
                headings=["", "Registry Key/Subkey Path", "Registry Name", "Type", "Data"],
                auto_size_columns=False,
                vertical_scroll_only=False,
                justification="left",
                num_rows=35,
                key="-TABLE_EDITOR-",
                col_widths=[5, 60, 25, 30, 100],
                background_color="#045D5D",
                text_color="white",
                bind_return_key=True,
                enable_events=True,
                enable_click_events=True,
                row_colors=[("white", "#045D5D"), ("white", "yellow"), ("white", "red")],
                tooltip=None,
                expand_x=False,
                expand_y=True
            ),
        ],
    ]

    return layout


# size checker
def run_size_check(folder_path):
    try:
        item_sizes_path = get_current_file_path("item_sizes.json")
        # Load item_sizes.json
        try:
            with open(item_sizes_path, 'r') as json_file:
                item_sizes = json.load(json_file)
        except FileNotFoundError:
            sg.popup_error("item_sizes.json not found. Please make sure it's in the same directory as this script.", title="Error")
            item_sizes = {}

        if folder_path:
            # Compare the sizes and update the table
            result_data = compare_sizes(folder_path, item_sizes)

            # Add numbering column to the result_data
            result_data_with_numbering = [
                [i + 1] + list(row) for i, row in enumerate(result_data)
            ]

            # Update the table with the numbered data
            window["-SIZE_TABLE-"].update(values=result_data_with_numbering)

            # Calculate the number of matched sizes found (Pass status)
            num_matched_sizes = sum(1 for _, _, _, status in result_data if status == "Pass")

            # Calculate the number of sizes found (total items in the table)
            num_sizes_found = len(result_data)

            # Calculate the number of sizes not found (Number of Required Size - Number of Sizes Found)
            num_required_size = int(window["-REQUIRED_SIZE-"].get())

            # Calculate the number of "Missing" statuses
            num_missing_sizes = sum(1 for _, _, _, status in result_data if status == "Missing")

            # Update the corresponding GUI elements
            window["-MATCHED_SIZE-"].update(num_matched_sizes)
            window["-FOUND_SIZE-"].update(num_sizes_found)
            window["-NOT_FOUND_SIZE-"].update(num_missing_sizes)

            # Collect names of incorrect sizes applications
            incorrect_sizes_list = [
                (app, status, item_sizes.get(app, "N/A")) for app, _, _, status in result_data if
                status in ("Failed", "Missing")
            ]

            # Update the "List of Incorrect Sizes Applications"
            window["-INCORRECT_SIZES-"].update("\n".join(
                [f"{app}: {status}, Expected Size: {expected_size}" for app, status, expected_size in
                 incorrect_sizes_list]))
    except Exception as e:
        sg.popup_error(f"An error occurred while checking folder size: {e}", title="Error")


# software checker
# Function to run "Main_21.py" when the "Check" button is clicked
def run_main_21():
    try:
        # Merge the existing data with the new data for current PC software
        current_pc_output_file = merge_current_pc_data("currentpc_software.json")

        # Load software data
        # software_data = load_software_data('C:/Users/hai-kent.kok/Desktop/My/software.json')
        # get_file_path_sw ='C:/Users/xiau-yen.kelly-teo/Desktop/SW & Size checker - Copy/software.json'
        get_file_path_sw = get_current_file_path('software.json')
        os.makedirs(os.path.dirname(get_file_path_sw), exist_ok=True)
        software_data = load_software_data(get_file_path_sw)

        # Load current PC data
        current_pc_data = load_current_pc_data(current_pc_output_file)  # load the appended data in this file

        # Generate results
        results = generate_results(software_data, current_pc_data)

        # Define the output file path for the results
        results_output_file = "results.json"

        # Combine the script directory and relative output file path for results
        '''script_directory = os.path.dirname(os.path.abspath(__file__))#get the currrent working path first
        results_output_file = os.path.join(script_directory, results_output_file)# combine the working path and the file we want to get the complete file path
'''
        # Save the results to the JSON file
        # write the data into the JSON file
        write_to_json(results, results_output_file)
        '''with open(results_output_file, 'w') as results_file:
            json.dump(results, results_file, indent=4)
'''
        # After running Main_21.py, update the table with the latest data
        update_gui()
    except Exception as e:
        sg.popup_error(f"An unexpected error occurred: {e}", title="Error")


# Function to load software data
def load_software_data(file_path):
    with open(file_path, 'r') as software_file:
        return json.load(software_file)['software_list']


# Function to load current PC data
def load_current_pc_data(file_path):
    with open(file_path, 'r') as current_pc_file:
        return json.load(current_pc_file)


# Function to update the table with the latest data and row colors
def update_gui():
    try:
        # Load the latest data from "results.json"
        # with open('C:/Users/hai-kent.kok/Desktop/My/results.json', 'r') as json_file:
        json_path = get_current_file_path('results.json')
        os.makedirs(os.path.dirname(json_path), exist_ok=True)
        with open(json_path, 'r') as json_file:
            results = json.load(json_file)

        updated_data = [
            [i + 1, software, result_data["Required Version"], result_data["Installed Version"], result_data["Status"]]
            for i, (software, result_data) in enumerate(results.items())
        ]

        # Calculate other statistics and update labels
        num_required = len(results)
        num_matched_version = len([result_data for result_data in results.values() if result_data["Status"] == "Pass"])
        num_found = len(
            [result_data for result_data in results.values() if result_data["Status"] in ["Pass", "Failed"]])
        num_not_found = len([result_data for result_data in results.values() if result_data["Status"] == "Missing"])

        window["-REQUIRED-"].update(num_required)
        window["-MATCHED_VERSION-"].update(num_matched_version)
        window["-FOUND-"].update(num_found)
        window["-NOT_FOUND-"].update(num_not_found)

        # Create a list of row colors based on the "Status" column
        row_colors = []
        for _, _, _, _, status in updated_data:
            if status == "Failed":
                row_colors.append(("white", "red"))  # White text on red background for "Failed"
            elif status == "Missing":
                row_colors.append(("white", "yellow"))  # White text on yellow background for "Missing"
            else:
                row_colors.append(("white", "#045D5D"))  # White text on default background

        # Update the table with the latest data and row colors
        window["-TABLE-"].update(values=updated_data, row_colors=row_colors)

        # Update the "List of Missing/Failed Status"
        missing_failed_list = [
            f"{software}: {result_data['Status']}"
            for software, result_data in results.items()
            if result_data['Status'] in ["Failed", "Missing"]
        ]
        window["-MISSING_FAILED-"].update("\n".join(missing_failed_list))
    except Exception as e:
        sg.popup_error(f"An error occurred while updating the table: {e}", title="Error")


# Function to merge existing data with new data for current PC software
def merge_current_pc_data(output_file):
    installed_programs = get_installed_programs()

    '''# Get the absolute path of the script's directory
    script_directory = os.path.dirname(os.path.abspath(__file__))

    # Combine the script directory and relative output file path for current PC software
    output_file = os.path.join(script_directory, output_file)'''

    output_file = get_current_file_path(output_file)
    write_to_json([], output_file)
    '''# Clear the contents of the existing file before saving
    with open(output_file, 'w') as json_file:
        json.dump([], json_file)'''

    # Merge the existing data with the new data
    updated_data = installed_programs
    write_to_json(updated_data, output_file)
    # Save the updated data to the JSON file for current PC software
    save_to_json(updated_data, output_file)

    return output_file


# Function to get a list of installed programs on the current PC
def get_installed_programs():
    program_list = []

    # Define the flags for 32-bit and 64-bit views
    flags = [win32con.KEY_WOW64_32KEY, win32con.KEY_WOW64_64KEY]

    # Connect to the Uninstall registry keys for each flag and HKEY_CURRENT_USER
    reg_hives = [win32con.HKEY_LOCAL_MACHINE, win32con.HKEY_CURRENT_USER]

    for reg_hive in reg_hives:
        for flag in flags:
            try:
                aReg = winreg.ConnectRegistry(None, reg_hive)
                aKey = winreg.OpenKey(aReg, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", 0,
                                      win32con.KEY_READ | flag)
                count_subkey = winreg.QueryInfoKey(aKey)[0]

                for i in range(count_subkey):
                    try:
                        asubkey_name = winreg.EnumKey(aKey, i)
                        asubkey = winreg.OpenKey(aKey, asubkey_name)
                        program = {}

                        program['Name'] = winreg.QueryValueEx(asubkey, "DisplayName")[0]
                        program['Version'] = winreg.QueryValueEx(asubkey, "DisplayVersion")[0]
                        program['Publisher'] = winreg.QueryValueEx(asubkey, "Publisher")[0]

                        # Try to retrieve the installation directory
                        try:
                            program['InstallLocation'] = winreg.QueryValueEx(asubkey, "InstallLocation")[0]
                        except EnvironmentError:
                            program['InstallLocation'] = None

                        program_list.append(program)
                    except EnvironmentError:
                        continue

            except Exception as e:
                print(f"Error accessing registry: {e}")

    return program_list


# Function to save data to a JSON file
def save_to_json(data, output_file):
    with open(output_file, 'w') as json_file:
        json.dump(data, json_file, indent=4)


# Function to match installed items with required items
def match_item(required_name, required_version, installed_item):
    name_similarity = fuzz.token_set_ratio(required_name.lower(), installed_item['Name'].lower())

    # Check if either version is 'N/A'
    if required_version == 'N/A' or installed_item['Version'] == 'N/A':
        return name_similarity >= 70

    # Special case: Handle Google Chrome
    if 'Google Chrome' in required_name.lower() or 'chrome' in required_name.lower():
        return name_similarity >= 80 and fuzz.ratio('Google Chrome', installed_item['Name']) >= 80

    # Special case for "Microsoft Visual C++"
    if 'microsoft visual c++' in required_name.lower():
        # Extract the version from the name using regular expressions
        version_match_required = re.search(r'(\d+(\.\d+)+)', required_name, re.IGNORECASE)
        version_match_installed = re.search(r'(\d+(\.\d+)+)', installed_item['Name'], re.IGNORECASE)

        if version_match_required and version_match_installed:
            required_version = version_match_required.group(0)
            installed_version = version_match_installed.group(0)

            # Compare version similarity
            version_similarity = fuzz.ratio(required_version, installed_version)

            # Check if both name and version meet the similarity threshold
            return name_similarity >= 70 and version_similarity >= 70

    # Special case: Handle WinMerge
    if 'winmerge' in required_name.lower():
        # Extract the version from the name using regular expressions
        version_match = re.search(r'WinMerge (\d+(\.\d+)+)', installed_item['Name'], re.IGNORECASE)
        if version_match:
            installed_version = version_match.group(1)
            # Compare the extracted version with the required version
            if version.parse(installed_version) >= version.parse(required_version):
                return True

    # Special case for "NVIDIA Graphics Driver"
    if 'NVIDIA Graphics Driver' in required_name.lower():
        # Extract the version number from the name
        version_match_required = re.search(r'(\d+(\.\d+)+)', required_name, re.IGNORECASE)

        if version_match_required:
            required_version = version_match_required.group(0)

            # Assume the number behind the name is the installed version
            installed_version_match = re.search(r'(\d+(\.\d+)+)', installed_item['Name'])

            if installed_version_match:
                installed_version = installed_version_match.group(0)

                # Compare version similarity
                version_similarity = fuzz.ratio(required_version, installed_version)

                # Check if both name and version meet the similarity threshold
                return name_similarity >= 70 and version_similarity >= 70

    # For other cases, split the name into words and check if all words appear in the installed software name
    required_name_words = required_name.lower().split()
    installed_name_words = installed_item['Name'].lower().split()
    name_similarity_words = fuzz.token_set_ratio(required_name_words, installed_name_words)

    # Parse the version strings and compare
    version_similarity = fuzz.ratio(required_version, installed_item['Version'])
    if 'microsoft visual c++' in required_name.lower():
        return name_similarity >= 70 and version_similarity >= 70
    else:
        return name_similarity >= 70 and name_similarity_words >= 70


def generate_results(software_data, current_pc_data):
    results = {}

    # Create a dictionary to store the highest version number for each software name
    highest_versions = {}

    for software_item in software_data:
        software_name = software_item['Name']
        required_version = software_item['Required Version']

        matched_items = []
        for current_pc_item in current_pc_data:
            if match_item(software_name, required_version, current_pc_item):
                matched_items.append(current_pc_item)

        if matched_items:
            # Sort matched items by version number (in descending order)
            matched_items.sort(key=lambda x: get_version_number(x['Version']), reverse=True)

            # Select the item with the highest version number
            matched_item = matched_items[0]

            installed_version = matched_item['Version']
            if installed_version >= required_version:
                status = 'Pass'
            else:
                status = 'Failed'

            # Update the highest version for this software
            if software_name not in highest_versions:
                highest_versions[software_name] = installed_version
            else:
                # Check if the current version is higher than the stored highest version
                if get_version_number(installed_version) > get_version_number(highest_versions[software_name]):
                    highest_versions[software_name] = installed_version
        else:
            status = 'Missing'
            installed_version = 'N/A'

        # Special case: If the software name contains "WinMerge," display it as "WinMerge"
        if 'WinMerge' in software_name:
            software_name = 'WinMerge'

        results[software_name] = {
            'Required Version': required_version,
            'Installed Version': installed_version,
            'Status': status
        }

    return results


def get_version_number(version_str):
    # Extract the version number from a string using regular expressions
    version_match = re.search(r'(\d+(\.\d+)*)', version_str)
    if version_match:
        return version_match.group(1)
    else:
        return '0'  # Return '0' if no version number is found


# Function to retrieve .NET Framework versions and save them to a JSON file
def retrieve_dotnet_versions():
    def get_dotnet_version(registry_path):
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, registry_path) as key:
                version = winreg.QueryValueEx(key, "Version")[0]
                return version
        except Exception as e:
            print(f"Error reading registry: {e}")
            return "N/A"

    # Get the versions of .NET Framework 3.5 and 4.8
    dotnet_35_version = get_dotnet_version(r"SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5")
    dotnet_48_version = get_dotnet_version(r"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full")

    # Create a dictionary to store the versions
    dotnet_versions = {
        ".NET Framework 3.5": dotnet_35_version,
        ".NET Framework 4.8": dotnet_48_version
    }

    # Save the versions to a JSON file
    with open('.network_frame.json', 'w') as json_file:
        json.dump(dotnet_versions, json_file, indent=4)

    print(".NET Framework 3.5 Version:", dotnet_35_version)
    print(".NET Framework 4.8 Version:", dotnet_48_version)


# Update Data in C:\\Tools\\GNU before Browse
def update_gnu_data():
    tools_path = "C:\\Tools"
    gnu_path = "C:\\Tools\\GNU"

    if not os.path.exists(tools_path) and not os.path.exists(gnu_path):
        sg.popup_error("Error: Both 'Tools' folder and 'GNU' folder are missing.", title="Error")
        return None

    if not os.path.exists(gnu_path):
        sg.popup_error("Error: 'GNU' folder is missing.", title="Error")
        return None

    run_size_check(gnu_path)
    return gnu_path

def sort_and_display_table(data):
    # Step 1: Sort the data, with abnormal status rows coming first.
    sorted_data = sorted(data, key=lambda x: x['status'] != 'abnormal')  # Adjust condition as needed

    # Step 2: Assign new indices in ascending order starting from 1.
    for index, row in enumerate(sorted_data, start=1):
        row['index'] = index

    # Step 3: Use this sorted and indexed data to load the table.
    load_table_data(sorted_data)


def load_table_data(json_file_path):
    if not os.path.isfile(json_file_path):
        return []

    if os.path.getsize(json_file_path) == 0:
        sg.popup_error(f"File {json_file_path} is empty.", title="Error")
        return []

    with open(json_file_path, 'r') as json_file:
        try:
            results = json.load(json_file)
            # Prepare the table data
            table_data = [
                [
                    i + 1,  # List column index
                    data["Registry Key/Subkey Path"],
                    data["Registry Name"],
                    data["Registry Type"],
                    data["Data"],
                    data["Status"]
                ]
                for i, (key, data) in enumerate(results.items())
            ]
            return table_data
        except json.JSONDecodeError:
            sg.popup_error(f"Error decoding JSON from file {json_file_path}.", title="Error")
            return []


# Create the layout for Page 1: Software Checker
page1_layout = create_page1_layout()
# Create the layout for Page 2: Size Checker
page2_layout = create_page2_layout()
# Create the layout for Page 3: Registry Checker
json_file_path = ''
table_data = load_table_data(json_file_path)
new_color = set_status_color(table_data)
page3_layout = create_page3_layout(table_data, new_color)
# Create the layout for Page 4: Registry Editor
page4_layout = create_page4_layout()
# Create the PySimpleGUI window with the default tab content
default_tutorial_layout = [
    [sg.Text("Welcome to 'Checker Tool'. This is the default page and a simple tutorial for users.\n\n"
             "1. For 'Page 1: Software Checker,' please kindly click on the 'Check' button to run version checking.\n\n"
             "2. For 'Page 2: Size Checker,' please select the 'GNU' folder manually by clicking on the 'Browse' button.\n\n"
             "3. For 'Page 3: Registry Checker,' please choose the machine type file then click 'Browse' button.\n"
             "   To compare the registry key, click 'Compare Registry' button.\n"
             "   To view more detail of the compared registry keys, click 'View Details' button in right bottom.\n"
             "   In Summary page, click 'Import' button to import the desired registry keys.\n"
             "   You can edit, delete and restore the redundant registry keys.\n\n"
             "4. For 'Page 4: Registry Editor,' please choose the machine type file then click 'Browse' button.\n"
             "   You can add, edit, and delete the registry keys in the golden file you selected.\n\n"
             "Thank you and have a nice day! ;)")],
    [sg.Checkbox("Disable Page 4 Password for testing mode", key="-DISABLE_PASSWORD-", enable_events=True)],
    [sg.Text("Status: Page 4 password is required.", key="-PASSWORD_STATUS-", size=(30, 1))]
]



def main_window():
    layout = [
        [
            sg.TabGroup(
                [
                    [
                        sg.Tab("Default: Tutorial", default_tutorial_layout, key="-TAB1-"),
                        sg.Tab("Page 1: Software Checker", page1_layout, key="-TAB2-"),
                        sg.Tab("Page 2: Size Checker", page2_layout, key="-TAB3-"),
                        sg.Tab("Page 3: Registry Checker", page3_layout, key="-TAB4-"),
                        sg.Tab("Page 4: Registry Editor", page4_layout, key="-TAB5-")
                    ],
                ],
                key="-TABS-",
                tab_location="top",
                enable_events=True
            ),
        ],
        [sg.Text("Selected Tab:", size=(20, 1)), sg.Text("", key="-STATUS-", size=(30, 1))]
        # Status row for selected tab
    ]

    window = sg.Window("Checker Tool", layout, finalize=True, size=(1000, 850), resizable=True,
                       enable_close_attempted_event=True)
    window.maximize()
    return window


# Create the PySimpleGUI window
window = main_window()
# For import key window in page 3
windowImport_active = False
# For window view more in page 3
window_view_more_active = False
window_edit_active = False
window_restore_active = False
window_add_data_active = False

original_table_data = []
original_size_table_data = []
original_reg_table_data = []

update_page1_data = False
update_page2_data = False
update_page3_data = False
update_page4_upper_data = False  # for table in the summary page(upper part)
update_page4_bottom_data = False  # for table in the summary page(bottom part)
update_restore_page_data = False  # for table in the restore page
update_editor_data = False  # for golden file editor page

folder_path = update_gnu_data()
table_data = window["-SIZE_TABLE-"].get()
original_size_table_data = table_data

update_page2_data = True

table_reg_data = window["-TABLE_REG-"].get()
original_reg_table_data = table_reg_data

update_page3_data = True


def parse_version(version):
    # Replace numbers with a numeric value and keep non-numeric parts as they are
    return tuple(int(x) if x.isdigit() else x for x in re.split(r'([0-9]+)', version))


def get_sort_key(item):
    try:
        # Try to convert the item to a numeric type
        return (0, float(item))
    except (ValueError, TypeError):
        # If conversion fails, return a tuple with a high value for data type and the parsed version
        return (1, parse_version(str(item).lower()))


# Sorting function
# Click the header of the column and sort
def sort_order_table(table, col_clicked, current_sort_order):
    try:
        # This takes the table and sorts everything given the column number (index)
        # Use a custom key function to handle mixed data types
        table = sorted(table, key=lambda row: get_sort_key(row[col_clicked]), reverse=current_sort_order[col_clicked])
        # Toggle the sort order for the next click
        current_sort_order[col_clicked] = not current_sort_order[col_clicked]
    except Exception as e:
        sg.popup_error('Error in sort_table', 'Exception in sort_table', e, title="Error")
    return table


def sort_table(window, event, row_count, col_count, current_sort_order):
    if isinstance(event, tuple):
        if event[0] == '-TABLE-':
            # event[2][0] is the row
            # event[2][1] is the column
            if event[2][0] == -1 and event[2][1] != -1:
                col_num_clicked = event[2][1]
                table_values = window["-TABLE-"].get()

                # Adjust the number of columns based on the provided col_count
                data = [row[:col_count] for row in table_values]

                new_table = sort_order_table(data, col_num_clicked, current_sort_order)
                update_table(window, new_table)


def sort_size_table(window, event, row_count, col_count, current_sort_order):
    if isinstance(event, tuple):
        if event[0] == '-SIZE_TABLE-':
            if event[2][0] == -1 and event[2][1] != -1:
                col_num_clicked = event[2][1]
                table_values = window["-SIZE_TABLE-"].get()

                # Adjust the number of columns based on the provided col_count
                data = [list(map(str, row[:col_count])) for row in table_values]

                new_table = sort_order_table(data, col_num_clicked, current_sort_order)
                update_size_table(window, new_table)


# Sorting functions for the registry checker tables
def sort_reg_table(window, event, col_count, current_sort_order):
    if isinstance(event, tuple):
        if event[0] == '-TABLE_REG-':
            # event[2][0] is the row
            # event[2][1] is the column
            if event[2][0] == -1 and event[2][1] != -1:
                col_num_clicked = event[2][1]
                table_values = window["-TABLE_REG-"].get()

                # disable the first column click event
                if col_num_clicked == 0:
                    return

                table_values = window["-TABLE_REG-"].get()

                # Adjust the number of columns based on the provided col_count
                data = [row[:col_count] for row in table_values]
                new_table = sort_order_table(data, col_num_clicked, current_sort_order)

                for i, row in enumerate(new_table):
                    row[0] = str(i + 1)

                new_color = set_status_color(new_table)
                window["-TABLE_REG-"].update(values=new_table, row_colors=new_color)


# Initialize current sort order with ascending for each column
current_sort_order_table = [False, False, False, False, False, False]


# current_sort_order_table = [False, False, False, False, False, False]

def sort_compared_reg_table(window, event, col_count, current_sort_order):
    if isinstance(event, tuple):
        if event[0] == '-TABLE_REG_COMPARED-':
            # event[2][0] is the row
            # event[2][1] is the column
            if event[2][0] == -1 and event[2][1] != -1:
                col_num_clicked = event[2][1]
                table_values = window["-TABLE_REG_COMPARED-"].get()

                # disable the first column click event
                if col_num_clicked == 0:
                    return

                table_values = window["-TABLE_REG_COMPARED-"].get()

                # Adjust the number of columns based on the provided col_count
                data = [row[:col_count] for row in table_values]

                new_table = sort_order_table(data, col_num_clicked, current_sort_order)
                window["-TABLE_REG_COMPARED-"].update(values=new_table)


# Initialize current sort order with ascending for each column
current_sort_order_table2 = [False, False, False, False, False, False, False, False]


def sort_redundant_reg_table(window, event, col_count, current_sort_order):
    if isinstance(event, tuple):
        if event[0] == '-TABLE_REDUNDANT-':
            # event[2][0] is the row
            # event[2][1] is the column
            if event[2][0] == -1 and event[2][1] != -1:
                col_num_clicked = event[2][1]
                table_values = window["-TABLE_REDUNDANT-"].get()

                # disable the first column click event
                if col_num_clicked == 0:
                    return

                table_values = window["-TABLE_REDUNDANT-"].get()

                # Adjust the number of columns based on the provided col_count
                data = [row[:col_count] for row in table_values]

                new_table = sort_order_table(data, col_num_clicked, current_sort_order)
                window["-TABLE_REDUNDANT-"].update(values=new_table)


# Initialize current sort order with ascending for each column
current_sort_order_table3 = [False, False, False, False, False]


def sort_restore_reg_table(window, event, col_count, current_sort_order):
    if isinstance(event, tuple):
        if event[0] == '-TABLE_RESTORE-':
            # event[2][0] is the row
            # event[2][1] is the column
            if event[2][0] == -1 and event[2][1] != -1:
                col_num_clicked = event[2][1]
                table_values = window["-TABLE_RESTORE-"].get()

                # disable the first column click event
                if col_num_clicked == 0:
                    return

                table_values = window["-TABLE_RESTORE-"].get()

                # Adjust the number of columns based on the provided col_count
                data = [row[:col_count] for row in table_values]

                new_table = sort_order_table(data, col_num_clicked, current_sort_order)
                window["-TABLE_RESTORE-"].update(values=new_table)


# Initialize current sort order with ascending for each column
current_sort_order_table4 = [False, False, False, False, False]


# Handle table selection in registry checker tool
# Allow click and select the rows and get the result from the selected rows
# for the table above in summary page
def handle_table_selection(window, event, values):
    selected_list_index = []
    result = []  # Change to a list to hold multiple row data

    if isinstance(event, tuple) and event[0] == '-TABLE_REG_COMPARED-':
        if event[2] and len(event[2]) > 0:
            row_index = event[2][0]
            column_index = event[2][1]

            # Retrieve the table values
            table_values = window["-TABLE_REG_COMPARED-"].get()

            # Check if the click is on the first column
            if column_index == 0 and row_index == -1:
                if not hasattr(handle_table_selection, 'click_time'):
                    handle_table_selection.click_time = 0

                if handle_table_selection.click_time == 0:
                    # sg.popup_ok("First click: set all to checked")
                    handle_table_selection.click_time = 1

                    # Set all checkboxes to checked
                    for row in table_values:
                        if row[0] == BLANK_BOX:
                            row[0] = CHECKED_BOX

                elif handle_table_selection.click_time == 1:
                    # sg.popup_ok("Second click: set all to unchecked")
                    handle_table_selection.click_time = 0

                    # Set all checkboxes to unchecked
                    for row in table_values:
                        if row[0] == CHECKED_BOX:
                            row[0] = BLANK_BOX

                # Update the table with the new values
                window["-TABLE_REG_COMPARED-"].update(values=table_values)

                # Update the selected rows list
                selected_list_index = [i for i, row in enumerate(table_values) if row[0] == CHECKED_BOX]
                print(f"Selected indices after toggle: {selected_list_index}")

                # Check if any rows are selected
                if not selected_list_index:
                    file_path = get_selected_reg_file_path()
                    '''script_dir = os.path.dirname(os.path.abspath(__file__))
                    output_file = os.path.join(script_dir, file_path)'''
                    output_file = get_current_file_path(file_path)
                    # Clear the data in file first so we can save the latest data into it
                    write_to_json([], output_file)
                    '''with open(output_file, 'w') as json_file:
                        json.dump([], json_file)'''

                    # Show the import button (all, fail, missing, ... button)
                    window.find_element('-update-').update(visible=True)
                    window.find_element("-update-").update(disabled=False)
                    # Hide the import selected reg key button
                    window.find_element('-updateSelected-').update(visible=False)

                else:
                    for i in selected_list_index:
                        row = table_values[i]
                        matched_registry_dict = {
                            'Registry Key/Subkey Path': row[1],
                            'Registry Name': row[2],
                            'Expected Type': row[4],
                            'Current Type': row[3],
                            'Expected Data': row[6],
                            'Current Data': row[5]
                        }
                        result.append(matched_registry_dict)  # Append each row's dictionary to the result list
                    # Show the import button (all, fail, missing, ... button)
                    window.find_element('-update-').update(visible=False)
                    # Hide the import selected reg key button
                    window.find_element('-updateSelected-').update(visible=True)
                    print(f"Result: {result}")

            else:
                # Handle row clicks (non-header clicks)
                if isinstance(row_index, int) and row_index >= 0:
                    if 0 <= row_index < len(table_values):
                        clicked_row_data = table_values[row_index]
                        print(f"Clicked row index: {row_index}")
                        print(f"Clicked row data: {clicked_row_data}")

                        if clicked_row_data[0] == BLANK_BOX:
                            # If checkbox is unchecked, check it
                            clicked_row_data[0] = CHECKED_BOX
                            window.find_element('-update-').update(visible=False)
                            window.find_element('-updateSelected-').update(visible=True)
                        else:
                            # If checkbox is checked, uncheck it
                            clicked_row_data[0] = BLANK_BOX
                            window.find_element('-updateSelected-').update(visible=False)
                            window.find_element('-update-').update(visible=True)
                            window.find_element("-update-").update(disabled=False)

                        # Update the table with the new values
                        table_values[row_index] = clicked_row_data
                        window["-TABLE_REG_COMPARED-"].update(values=table_values)

                        # Update the selected rows list
                        selected_list_index = [i for i, row in enumerate(table_values) if row[0] == CHECKED_BOX]
                        print(f"Selected indices: {selected_list_index}")

                        if not selected_list_index:
                            file_path = get_selected_reg_file_path()
                            '''script_dir = os.path.dirname(os.path.abspath(__file__))
                            output_file = os.path.join(script_dir, file_path)'''
                            output_file = get_current_file_path(file_path)
                            write_to_json([], output_file)
                            # Clear the data in file first so we can save the latest data into it
                            '''with open(output_file, 'w') as json_file:
                                json.dump([], json_file)
'''
                            # Show the import button (all, fail, missing, ... button)
                            window.find_element('-update-').update(visible=True)
                            window.find_element("-update-").update(disabled=False)
                            # Hide the import selected reg key button
                            window.find_element('-updateSelected-').update(visible=False)

                        else:
                            for i in selected_list_index:
                                row = table_values[i]
                                matched_registry_dict = {
                                    'Registry Key/Subkey Path': row[1],
                                    'Registry Name': row[2],
                                    'Expected Type': row[4],
                                    'Current Type': row[3],
                                    'Expected Data': row[6],
                                    'Current Data': row[5]
                                }
                                result.append(matched_registry_dict)  # Append each row's dictionary to the result list

                            # Show the import button (all, fail, missing, ... button)
                            window.find_element('-update-').update(visible=False)
                            # Hide the import selected reg key button
                            window.find_element('-updateSelected-').update(visible=True)
                            print(f"Result: {result}")

    return result


# for the table bottom in summary page
def handle_table_selection2(window, event, values):
    selected_list_index = []
    result = []  # Change to a list to hold multiple row data

    if isinstance(event, tuple) and event[0] == '-TABLE_REDUNDANT-':
        if event[2] and len(event[2]) > 0:
            row_index = event[2][0]
            column_index = event[2][1]

            # Retrieve the table values
            table_values = window["-TABLE_REDUNDANT-"].get()

            # Check if the click is on the first column
            if column_index == 0 and row_index == -1:
                if not hasattr(handle_table_selection2, 'click_time'):
                    handle_table_selection2.click_time = 0

                if handle_table_selection2.click_time == 0:
                    # sg.popup_ok("First click: set all to checked")
                    handle_table_selection2.click_time = 1

                    # Set all checkboxes to checked
                    for row in table_values:
                        if row[0] == BLANK_BOX:
                            row[0] = CHECKED_BOX

                elif handle_table_selection2.click_time == 1:
                    # sg.popup_ok("Second click: set all to unchecked")
                    handle_table_selection2.click_time = 0

                    # Set all checkboxes to unchecked
                    for row in table_values:
                        if row[0] == CHECKED_BOX:
                            row[0] = BLANK_BOX

                # Update the table with the new values
                window["-TABLE_REDUNDANT-"].update(values=table_values)

                # Update the selected rows list
                selected_list_index = [i for i, row in enumerate(table_values) if row[0] == CHECKED_BOX]
                print(f"Selected indices after toggle: {selected_list_index}")

                # Disable the edit function when more than one result select
                # disable the delete function when no result is select
                if len(selected_list_index) != 1:  # more than one
                    window.find_element("-EDIT-").update(disabled=True)
                    window.find_element("-DELETE-").update(disabled=False)
                else:
                    window.find_element("-EDIT-").update(disabled=False)
                    window.find_element("-DELETE-").update(disabled=False)

                # Check if any rows are selected
                if not selected_list_index:
                    file_path = get_selected_redundant_file_path()
                    '''script_dir = os.path.dirname(os.path.abspath(__file__))
                    output_file = os.path.join(script_dir, file_path)'''
                    output_file = get_current_file_path(file_path)
                    write_to_json([], output_file)
                    window.find_element("-DELETE-").update(disabled=True)
                    # Clear the data in file first so we can save the latest data into it
                    '''with open(output_file, 'w') as json_file:
                            json.dump([], json_file)
'''
                else:
                    for i in selected_list_index:
                        row = table_values[i]
                        matched_registry_dict = {
                            'Registry Key/Subkey Path': row[1],
                            'Registry Name': row[2],
                            'Type': row[3],
                            'Data': row[4]
                        }
                        result.append(matched_registry_dict)  # Append each row's dictionary to the result list

            else:
                # Handle row clicks (non-header clicks)
                if isinstance(row_index, int) and row_index >= 0:
                    if 0 <= row_index < len(table_values):
                        clicked_row_data = table_values[row_index]

                        if clicked_row_data[0] == BLANK_BOX:
                            # If checkbox is unchecked, check it
                            clicked_row_data[0] = CHECKED_BOX

                        else:
                            # If checkbox is checked, uncheck it
                            clicked_row_data[0] = BLANK_BOX

                        # Update the table with the new values
                        table_values[row_index] = clicked_row_data
                        window["-TABLE_REDUNDANT-"].update(values=table_values)

                        # Update the selected rows list
                        selected_list_index = [i for i, row in enumerate(table_values) if row[0] == CHECKED_BOX]
                        print(f"Selected indices: {selected_list_index}")

                        if len(selected_list_index) != 1:  # more than one
                            window.find_element("-EDIT-").update(disabled=True)
                            window.find_element("-DELETE-").update(disabled=False)
                        else:
                            window.find_element("-EDIT-").update(disabled=False)
                            window.find_element("-DELETE-").update(disabled=False)

                        if not selected_list_index:
                            file_path = get_selected_redundant_file_path()
                            '''script_dir = os.path.dirname(os.path.abspath(__file__))
                            output_file = os.path.join(script_dir, file_path)'''
                            output_file = get_current_file_path(file_path)
                            write_to_json([], output_file)
                            window.find_element("-DELETE-").update(disabled=True)
                            # Clear the data in file first so we can save the latest data into it
                            '''with open(output_file, 'w') as json_file:
                                json.dump([], json_file)
'''
                        else:
                            for i in selected_list_index:
                                row = table_values[i]
                                matched_registry_dict = {
                                    'Registry Key/Subkey Path': row[1],
                                    'Registry Name': row[2],
                                    'Type': row[3],
                                    'Data': row[4]
                                }
                                result.append(matched_registry_dict)  # Append each row's dictionary to the result list

                            print(f"Result: {result}")

    return result


# for the restore page
# for the table bottom in summary page
def handle_table_selection3(window, event, values):
    selected_list_index = []
    result = []  # Change to a list to hold multiple row data

    if isinstance(event, tuple) and event[0] == '-TABLE_RESTORE-':
        if event[2] and len(event[2]) > 0:
            row_index = event[2][0]
            column_index = event[2][1]

            # Retrieve the table values
            table_values = window["-TABLE_RESTORE-"].get()

            # Check if the click is on the first column
            if column_index == 0 and row_index == -1:
                if not hasattr(handle_table_selection3, 'click_time'):
                    handle_table_selection3.click_time = 0

                if handle_table_selection3.click_time == 0:
                    # sg.popup_ok("First click: set all to checked")
                    handle_table_selection3.click_time = 1

                    # Set all checkboxes to checked
                    for row in table_values:
                        if row[0] == BLANK_BOX:
                            row[0] = CHECKED_BOX

                elif handle_table_selection3.click_time == 1:
                    # sg.popup_ok("Second click: set all to unchecked")
                    handle_table_selection3.click_time = 0

                    # Set all checkboxes to unchecked
                    for row in table_values:
                        if row[0] == CHECKED_BOX:
                            row[0] = BLANK_BOX

                # Update the table with the new values
                window["-TABLE_RESTORE-"].update(values=table_values)

                # Update the selected rows list
                selected_list_index = [i for i, row in enumerate(table_values) if row[0] == CHECKED_BOX]
                print(f"Selected indices after toggle: {selected_list_index}")

                # Check if any rows are selected
                if not selected_list_index:
                    file_path = get_selected_restore_file_path()
                    '''script_dir = os.path.dirname(os.path.abspath(__file__))
                    output_file = os.path.join(script_dir, file_path)'''
                    output_file = get_current_file_path(file_path)
                    write_to_json([], output_file)
                    # Clear the data in file first so we can save the latest data into it
                    '''with open(output_file, 'w') as json_file:
                            json.dump([], json_file)

'''
                    window.find_element("-restoreSelected-").update(disabled=True)
                else:
                    window.find_element("-restoreSelected-").update(disabled=False)
                    for i in selected_list_index:
                        row = table_values[i]
                        matched_registry_dict = {
                            'Registry Key/Subkey Path': row[1],
                            'Registry Name': row[2],
                            'Type': row[3],
                            'Data': row[4]
                        }
                        result.append(matched_registry_dict)  # Append each row's dictionary to the result list
            else:
                # Handle row clicks (non-header clicks)
                if isinstance(row_index, int) and row_index >= 0:
                    if 0 <= row_index < len(table_values):
                        clicked_row_data = table_values[row_index]
                        if clicked_row_data[0] == BLANK_BOX:
                            # If checkbox is unchecked, check it
                            clicked_row_data[0] = CHECKED_BOX

                        else:
                            # If checkbox is checked, uncheck it
                            clicked_row_data[0] = BLANK_BOX

                        # Update the table with the new values
                        table_values[row_index] = clicked_row_data
                        window["-TABLE_RESTORE-"].update(values=table_values)

                        # Update the selected rows list
                        selected_list_index = [i for i, row in enumerate(table_values) if row[0] == CHECKED_BOX]
                        print(f"Selected indices: {selected_list_index}")
                        if not selected_list_index:
                            file_path = get_selected_restore_file_path()

                            '''script_dir = os.path.dirname(os.path.abspath(__file__))
                            output_file = os.path.join(script_dir, file_path)'''
                            output_file = get_current_file_path(file_path)
                            write_to_json([], output_file)
                            # Clear the data in file first so we can save the latest data into it
                            '''with open(output_file, 'w') as json_file:
                                json.dump([], json_file)
'''
                            window.find_element("-restoreSelected-").update(disabled=True)

                        else:
                            window.find_element("-restoreSelected-").update(disabled=False)
                            for i in selected_list_index:
                                row = table_values[i]
                                matched_registry_dict = {
                                    'Registry Key/Subkey Path': row[1],
                                    'Registry Name': row[2],
                                    'Type': row[3],
                                    'Data': row[4]
                                }
                                result.append(matched_registry_dict)  # Append each row's dictionary to the result list

                            print(f"Result: {result}")

    return result
'''
def restore_page4():
    # Path to store deleted registry data for Page 4
    get_deleted_page4_data_path = "data/deleted_registry_data_page4.json"

    # Ensure directory exists and the deleted file exists
    os.makedirs(os.path.dirname(get_deleted_page4_data_path), exist_ok=True)

    # If the file doesn't exist, create an empty one
    if not os.path.isfile(get_deleted_page4_data_path):
        write_to_json([], get_deleted_page4_data_path)

    # Load the deleted data from the file
    get_deleted_page4_data = load_registry_from_json2(get_deleted_page4_data_path)

    # Check if there's any data to restore, enable or disable Restore button accordingly
    if get_deleted_page4_data:
        window.find_element("-RESTORE_PAGE4-").update(disabled=False)
    else:
        window.find_element("-RESTORE_PAGE4-").update(disabled=True)
'''

def get_selected_reg_file_path():
    return 'current_registry_pc.json'


def update_table(window, data):
    # window["-TABLE-"].update(values=data)
    for i, row in enumerate(data):
        row[0] = str(i + 1)
    window["-TABLE-"].update(values=data)


def update_size_table(window, data):
    window["-SIZE_TABLE-"].update(values=data)


# for page 3
# get the selected file
file_path = ""


# Function to set file_path
def set_file_path(path):
    global file_path
    file_path = path


# Function to get file_path
def get_file_path():
    return file_path


def get_selected_reg_file_path():
    # file_path = 'C:/Users/xiau-yen.kelly-teo/Desktop/SW & Size checker - Copy/selected_registry_results.json'
    file_path = get_current_file_path('data/selected_registry_results.json')

    return file_path


def get_selected_redundant_file_path():
    # file_path = 'C:/Users/xiau-yen.kelly-teo/Desktop/SW & Size checker - Copy/selected_redundant_data.json'
    file_path = get_current_file_path('data/selected_redundant_data.json')

    return file_path


def get_selected_restore_file_path():
    # file_path = 'C:/Users/xiau-yen.kelly-teo/Desktop/SW & Size checker - Copy/selected_redundant_data.json'
    file_path = get_current_file_path('data/selected_restore_data.json')

    return file_path


# For redundant registry key table in summary part
# Edit the selected row redundant keys
def edit_selected_row(selected_redundant_data, window):
    edited_key = []
    data = [
        [
            result_reg_data["Registry Key/Subkey Path"],
            result_reg_data["Registry Name"],
            result_reg_data["Type"],
            result_reg_data["Data"]
        ]
        for i, result_reg_data in enumerate(selected_redundant_data)
    ]

    for path, name, reg_type, data in data:
        edit_path = window.find_element("-EDIT_PATH-").get()
        edit_name = window.find_element("-EDIT_NAME-").get()
        edit_type = window.find_element("-EDIT_TYPE_DROPDOWN-").get()
        decimal_selected = window.find_element("-FORMAT_DECIMAL-").get()
        hex_selected = window.find_element("-FORMAT_HEX-").get()

        if edit_type == "String":
            updated_edit_type = "REG_SZ"
            edit_data = window.find_element("-EDIT_DATA-").get()

        elif edit_type == "Binary":
            updated_edit_type = "REG_BINARY"
            edit_data = window.find_element("-EDIT_DATA-").get()

        elif edit_type == "DWORD(32-bit)":
            updated_edit_type = "REG_DWORD_LITTLE_ENDIAN"
            edit_data = window.find_element("-EDIT_DATA-").get()

            # Ensure edit_data is in hexadecimal format
            if decimal_selected:
                edit_data = get_valid_integer_input(window)
                if edit_data is not None:
                    edit_data = f'0x{int(edit_data):08X}'  # Convert from decimal to hexadecimal
                else:
                    return []  # Exit if user cancels or closes the window

            elif hex_selected:
                # Ensure the hex string is formatted correctly
                edit_data = f'0x{edit_data.upper().lstrip("0X")}'

        elif edit_type == "QWORD(64-bit)":
            updated_edit_type = "REG_QWORD_LITTLE_ENDIAN"
            edit_data = window.find_element("-EDIT_DATA-").get()

            # Ensure edit_data is in hexadecimal format
            if decimal_selected:
                edit_data = get_valid_integer_input(window)
                if edit_data is not None:
                    edit_data = f'0x{int(edit_data):016X}'  # Convert from decimal to hexadecimal
                else:
                    return []  # Exit if user cancels or closes the window
            elif hex_selected:
                edit_data = f'0x{edit_data.upper().lstrip("0X")}'

        elif edit_type == "Multi-String":
            updated_edit_type = "REG_MULTI_SZ"
            edit_data = window.find_element("-EDIT_DATA-").get().split('\0')  # Split by null characters
            edit_data = str(edit_data)

        elif edit_type == "Expandable String":
            updated_edit_type = "REG_EXPAND_SZ"
            edit_data = window.find_element("-EDIT_DATA-").get()
            data = str(data)

        edited_key.append((edit_path, edit_name, edit_data, updated_edit_type))
        # Log the change after processing all data
        write_into_event_log(
            "The registry key has been edited from \n\t'" + path + "\\" + name + ", Type: " + reg_type + ", Data: " + data + "' to \n\t'" + edit_path + "\\" + edit_name + ", Type: " + updated_edit_type + ", Data: " + edit_data + "'.")
        print(f"Edited registry data: {edited_key}")

    return edited_key


def get_valid_integer_input(window):
    while True:
        edit_data = window.find_element("-EDIT_DATA-").get()
        if is_decimal(edit_data):
            return edit_data
        else:
            sg.popup_error("Please enter a valid integer.", title="Error")
            window.find_element("-EDIT_DATA-").update('')  # Optionally clear the field
            # window.find_element("-SAVE-").update(disabled=True)  # Disable save button until valid input is entered
            event, _ = window.read()  # Wait for user input
            if event == sg.WIN_CLOSED:
                return None  # Exit if the window is closed


def is_decimal(value):
    try:
        int(value)
        return True
    except ValueError:
        return False


def is_hex(value):
    try:
        int(value, 16)
        return True
    except ValueError:
        return False


# Delete the selected redundant keys in summary page
def delete_redundant_keys(selected_redundant_data):
    """ Delete registry keys and specific values based on the given data. """
    try:
        # Connect to the registry (HKEY_LOCAL_MACHINE in this case)
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as hkey:
            for key in selected_redundant_data:
                key_path = key['Registry Key/Subkey Path']
                value_name = key['Registry Name']
                # specific_key = key_path + "\\" + value_name
                # Call delete_key_recursively for the key path
                delete_key_recursively(hkey, value_name)

                # Open the key and delete the specific value
                try:
                    with winreg.OpenKey(hkey, key_path, 0, winreg.KEY_ALL_ACCESS) as key:
                        winreg.DeleteValue(key, value_name)
                        print(f"Deleted value '{value_name}' from key '{key_path}' successfully.")
                        write_into_event_log(f"Deleted value '{value_name}' from key '{key_path}' successfully.")
                except OSError as e:
                    print(f"Error deleting value '{value_name}' from key '{key_path}': {e}")
                    write_into_error_log(f"Error deleting value '{value_name}' from key '{key_path}': {e}")

    except Exception as e:
        print(f"Unexpected error: {e}")
        write_into_error_log(f"Unexpected error: {e}")


def delete_key_recursively(hkey, value_name):
    """ Recursively delete a registry key and its subkeys and values. """
    try:
        # Open the registry key
        open_key = winreg.OpenKey(hkey, value_name, 0, winreg.KEY_ALL_ACCESS)
        info_key = winreg.QueryInfoKey(open_key)
        for x in range(0, info_key[0]):
            subkey = winreg.EnumKey(open_key, 0)
            try:
                winreg.DeleteKey(open_key, subkey)
            except:
                delete_key_recursively(open_key, subkey)

        winreg.DeleteKey(open_key, "")
        open_key.Close()

    except OSError as e:
        print(f"Error deleting key '{value_name}': {e}")
        write_into_error_log(f"Error deleting key '{value_name}': {e}")


def add_registry_entry(path, name, reg_type, data):
    try:
        # Simulate adding the registry entry to your data source (for example, a JSON file)
        # You can modify this logic to write the data to the appropriate location
        new_entry = {
            "path": path,
            "name": name,
            "type": reg_type,
            "data": data
        }
        # Append this new entry to your registry file or system
        # Example: Append to a JSON file, or a database, etc.
        # Here you would include the logic to store this new entry
        return True  # Simulate success
    except Exception as e:
        print(f"Error adding registry entry: {e}")
        return False

'''
def get_selected_restore_file_path_page4():
    # Assuming you store deleted keys from Page 4 in a specific file
    deleted_registry_file_path = "data/deleted_registry_data_page4.json"

    # Check if the deleted registry data file exists
    if os.path.exists(deleted_registry_file_path):
        # Load the deleted data to check if there is any content
        deleted_registry_data = load_registry_from_json2(deleted_registry_file_path)

        if deleted_registry_data:
            return deleted_registry_file_path  # Return the path if data exists
        else:
            sg.popup_error("No deleted registry data available to restore for Page 4.")
            return None
    else:
        sg.popup_error("No deleted registry data found for Page 4.")
        return None
'''

# Make new window for the update page - update the registry keys
def makeWin(title):
    layout = [
        [sg.Push(), sg.Text('Do you want to update/add the keys into your current pc?'), sg.Push()],
        [
            sg.Push(), sg.Button("Yes (All)", key='_updateAll_', disabled=False),
            sg.Button("Yes (Fail Key)", key='_updateFail_', disabled=False),
            sg.Button("Yes (Missing Key)", key='_updateMissing_', disabled=False), sg.Button("No", key='_NO_'),
            sg.Push()
        ]
    ]
    return sg.Window(title, layout, finalize=True, size=(400, 80), resizable=True)


# Make new window for the summary page
def makeWin2(title):
    file_path = get_file_path()
    file_name = Path(file_path).stem
    selected_file_output = sg.Text(file_name, font=("Helvetica", 14, "bold"))
    layoutImport = [
        [sg.Push(), sg.pin(sg.Text("Golden List Selected:", font=("Helvetica", 14, "bold"))), selected_file_output,
         sg.Push()],

        [sg.Frame('List of Fail/Missing Registry Keys', font=("Helvetica", 12, "bold"), layout=[
            [sg.pin(sg.Text("Search:", key="-SEARCH_TEXT2-", visible=False)),
         sg.InputText(key="-SEARCH_REG2-", size=(20, 1), do_not_clear=True, enable_events=True),
         sg.pin(sg.Button("Update", key='-updateSelected-', visible=False)),
         sg.pin(sg.Button("Update", key='-update-', visible=True, disabled=False)),
         sg.pin(sg.Button("Import", key='_importBackup_')),
         sg.Push(),
         sg.Text("Fail Registry Keys:", font=("Helvetica", 11)),
         sg.Text("0", key="-FAIL_REG-", font=("Helvetica", 11)),
         sg.Text(",    Missing Registry Keys:", font=("Helvetica", 11)),
         sg.Text("0", key="-NOT_FOUND_REG-", font=("Helvetica", 11))
         ],
            [
                sg.Table(
                    values=[],  # Initialize with empty values; to be updated later
                    headings=[" ", "Registry Key/SubKey Path", "Registry Name", "Current Type", "Expected Type",
                              "Current Data", "Expected Data", "Status"],
                    auto_size_columns=False,
                    vertical_scroll_only=False,
                    justification="left",
                    num_rows=20,
                    key="-TABLE_REG_COMPARED-",
                    col_widths=[8, 60, 25, 25, 25, 25, 25, 20],
                    background_color="#045D5D",
                    text_color="white",
                    bind_return_key=True,
                    row_colors=[("white", "#045D5D"), ("white", "yellow"), ("white", "red")],
                    enable_events=True,
                    enable_click_events=True,
                    select_mode=sg.TABLE_SELECT_MODE_EXTENDED,
                    tooltip=None,
                    expand_x=False,
                    expand_y=True
                ),
            ],
        ], element_justification="left", size=(450, 200), expand_x=True, expand_y=True)],

        [sg.Frame('List of Redundant Keys', font=("Helvetica", 12, "bold"), layout=[
            [sg.Text("Search Redundant:"),
             sg.InputText(key='-SEARCH_REDUNDANT-', size=(20, 1), do_not_clear=True, enable_events=True),
             sg.pin(sg.Button("Edit", key="-EDIT-", disabled=True)),
             sg.Button("Delete", key="-DELETE-", disabled=False),
             sg.pin(sg.Button("Restore", key="-RESTORE-", disabled=True)),
             sg.Push(),
             sg.Text("Redundant Registry Keys:", font=("Helvetica", 11)),
             sg.Text("0", key="-REDUNDANT_REG-", font=("Helvetica", 11))
             ],
            [
                sg.Table(
                    values=[],  # Initialize with empty values; to be updated later
                    headings=[" ", "Registry Key/SubKey Path", "Registry Name", "Type", "Data"],
                    auto_size_columns=False,
                    vertical_scroll_only=False,
                    justification="left",
                    num_rows=10,
                    key="-TABLE_REDUNDANT-",
                    col_widths=[8, 60, 50, 45, 45],
                    background_color="#045D5D",
                    text_color="white",
                    bind_return_key=True,
                    row_colors=[("white", "#045D5D"), ("white", "yellow"), ("white", "red")],
                    enable_events=True,
                    enable_click_events=True,
                    select_mode=sg.TABLE_SELECT_MODE_EXTENDED,
                    tooltip=None,
                    expand_x=False,
                    expand_y=True
                ),
            ],
        ], element_justification="left", size=(1000, 200), expand_x=True, expand_y=True)
         ]
    ]
    window = sg.Window(title, layoutImport, finalize=True, size=(1000, 750), resizable=True)

    # Maximize the window
    window.maximize()

    return window


# Make new window for the redundant keys editor page
'''
Registry type info:
String = winreg.REG_SZ, "REG_SZ"
Binary = winreg.REG_BINARY, "REG_BINARY"
DWORD(32-bit) = winreg.REG_DWORD_LITTLE_ENDIAN, "REG_DWORD_LITTLE_ENDIAN"
QWORD(64-bit) = winreg.REG_QWORD_LITTLE_ENDIAN: "REG_QWORD_LITTLE_ENDIAN"
Multi-String = winreg.REG_MULTI_SZ, "REG_MULTI_SZ"
Expandable String = winreg.REG_EXPAND_SZ, "REG_EXPAND_SZ"
'''


# dropdown_values_edit=["String","Binary","DWORD(32-bit)","QWORD(64-bit)", "Multi-String", "Expandable String"]
def makeWinEdit(title, path='', name='', reg_type='', reg_data=''):
    layout_edit = [
        [sg.Text('Registry Key/Subkey Path:', size=(20, 1)),
         sg.InputText(default_text=path, key='-EDIT_PATH-', readonly=True, size=(60, 1))],
        [sg.Text('Registry Name:', size=(20, 1)),
         sg.InputText(default_text=name, key='-EDIT_NAME-', readonly=True, size=(60, 1))],
        [sg.Text('Registry Type:', size=(20, 1)),
         sg.Combo(['String', 'Binary', 'DWORD(32-bit)', 'QWORD(64-bit)', 'Multi-String', 'Expandable String'],
                  default_value=reg_type, key="-EDIT_TYPE_DROPDOWN-", readonly=True),
         sg.Button('Select', key="-SELECT_FORMAT-")],
        [sg.pin(sg.Text('DWORD/QWORD Format:', size=(20, 1), key="-FORMAT-", visible=False)),
         sg.pin(sg.Radio('Decimal', "RADIO1", key="-FORMAT_DECIMAL-", default=True, visible=False)),
         sg.pin(sg.Radio('Hexadecimal', "RADIO1", key="-FORMAT_HEX-", visible=False))],
        [sg.Text('Registry Data:', size=(20, 1)), sg.InputText(default_text=reg_data, key="-EDIT_DATA-", size=(60, 1))],
        [sg.Push(), sg.Button("Save", key="-SAVE-", disabled=True), sg.Button('Cancel', key="-CANCEL-")]
    ]

    return sg.Window(title, layout_edit, finalize=True, size=(600, 200), resizable=True)


# Make the new window for the restore page
def makeWinRestore(title):
    layout = [
        [sg.Frame('List of Deleted Registry keys', font=("Helvetica", 12, "bold"), layout=[[
            # sg.InputText(key="-SEARCH_RESTORE-", size=(20, 1), do_not_clear=True, enable_events=True, default_text="Search", text_color="grey"),
            # sg.Button("Search", key="-SEARCH_BUTTON_RESTORE-"),
            sg.Button("Restore", key='-restoreSelected-', disabled=True),
        ],
            [
                sg.Table(
                    values=[],  # Initialize with empty values; to be updated later
                    headings=[" ", "Registry Key/SubKey Path", "Registry Name", "Type", "Data"],
                    auto_size_columns=False,
                    vertical_scroll_only=False,
                    justification="left",
                    num_rows=15,
                    key="-TABLE_RESTORE-",
                    col_widths=[3, 40, 20, 20, 20],
                    background_color="#045D5D",
                    text_color="white",
                    bind_return_key=True,
                    row_colors=[("white", "#045D5D"), ("white", "yellow"), ("white", "red")],
                    enable_events=True,
                    enable_click_events=True,
                    select_mode=sg.TABLE_SELECT_MODE_EXTENDED,
                    tooltip=None,
                    expand_x=False,
                    expand_y=True
                ),
            ],
            [#sg.Push(), sg.Text("Number of Deleted Registry Keys:", font=("Helvetica", 10)),
             #sg.Text("0", key="-DEL_REG-", font=("Helvetica", 10)), sg.Push()
             ],
        ], element_justification="left", size=(1000, 650))],
    ]
    return sg.Window(title, layout, finalize=True, size=(1000, 450), resizable=True)

'''
# Make the new window for the restore page for Page 4
def makeWinRestorePage4(title):
    layout = [
        [sg.Frame('List of Deleted Registry keys from Page 4', font=("Helvetica", 12, "bold"), layout=[[

            sg.Button("RestoreConfirmation", key='-restoreSelectedPage4-', disabled=True),
        ],
            [
                sg.Table(
                    values=[],  # Initialize with empty values; to be updated later
                    headings=[" ", "Registry Key/SubKey Path", "Registry Name", "Type", "Data"],
                    auto_size_columns=False,
                    vertical_scroll_only=False,
                    justification="left",
                    num_rows=15,
                    key="-TABLE_RESTORE_PAGE4-",
                    col_widths=[3, 40, 20, 20, 15],
                    background_color="#045D5D",
                    text_color="white",
                    bind_return_key=True,
                    row_colors=[("white", "#045D5D"), ("white", "yellow"), ("white", "red")],
                    enable_events=True,
                    enable_click_events=True,
                    select_mode=sg.TABLE_SELECT_MODE_EXTENDED,
                    tooltip=None,
                    expand_x=True,
                    expand_y=True
                ),
            ],
            [#sg.Push(), sg.Text("Number of Deleted Registry Keys from Page 4:", font=("Helvetica", 10)),
             #sg.Text("0", key="-DEL_REG_PAGE4-", font=("Helvetica", 10)), sg.Push()
             ],
        ], element_justification="left", size=(1000, 450))],
    ]
    return sg.Window(title, layout, finalize=True, size=(1000, 450), resizable=True)
'''

'''Page 4 window'''
# Make new window for the registry keys data editor page
'''
Registry type info:
String = winreg.REG_SZ, "REG_SZ"
Binary = winreg.REG_BINARY, "REG_BINARY"
DWORD(32-bit) = winreg.REG_DWORD_LITTLE_ENDIAN, "REG_DWORD_LITTLE_ENDIAN"
QWORD(64-bit) = winreg.REG_QWORD_LITTLE_ENDIAN: "REG_QWORD_LITTLE_ENDIAN"
Multi-String = winreg.REG_MULTI_SZ, "REG_MULTI_SZ"
Expandable String = winreg.REG_EXPAND_SZ, "REG_EXPAND_SZ"
'''


# dropdown_values_edit_page4=["String","Binary","DWORD(32-bit)","QWORD(64-bit)", "Multi-String", "Expandable String"]
def WindowAdd(title):
    layout_edit = [
        [sg.Text('Registry Key/Subkey Path:', size=(20, 1)),
         sg.InputText(default_text='Software\WOW6432Node\MV Technology', key='-ADD_PATH-', size=(60, 1))],
        [sg.Text('Registry Name:', size=(20, 1)), sg.InputText(default_text='', key='-ADD_NAME-', size=(60, 1))],
        [sg.Text('Registry Type:', size=(20, 1)),
         sg.Combo(['String', 'Binary', 'DWORD(32-bit)', 'QWORD(64-bit)', 'Multi-String', 'Expandable String'],
                  default_value='String', key="-ADD_TYPE_DROPDOWN-", readonly=True),
         sg.Button('Select', key="-SELECT_FORMAT_PAGE4-")],
        [sg.pin(sg.Text('DWORD/QWORD Format:', size=(20, 1), key="-ADD_FORMAT-", visible=False)),
         sg.pin(sg.Radio('Decimal', "RADIO1", key="-ADD_FORMAT_DECIMAL-", default=True, visible=False)),
         sg.pin(sg.Radio('Hexadecimal', "RADIO1", key="-ADD_FORMAT_HEX-", visible=False))],
        [sg.Text('Registry Data:', size=(20, 1)), sg.InputText(default_text='', key="-ADD_DATA-", size=(60, 1))],
        [sg.Push(), sg.Button("Add", key="-ADD_NEW_DATA-", disabled=True), sg.Button('Cancel', key="-CANCEL_ADD-")]
    ]

    return sg.Window(title, layout_edit, finalize=True, size=(600, 200), resizable=True)


'''
Registry type info:
String = winreg.REG_SZ, "REG_SZ"
Binary = winreg.REG_BINARY, "REG_BINARY"
DWORD(32-bit) = winreg.REG_DWORD_LITTLE_ENDIAN, "REG_DWORD_LITTLE_ENDIAN"
QWORD(64-bit) = winreg.REG_QWORD_LITTLE_ENDIAN: "REG_QWORD_LITTLE_ENDIAN"
Multi-String = winreg.REG_MULTI_SZ, "REG_MULTI_SZ"
Expandable String = winreg.REG_EXPAND_SZ, "REG_EXPAND_SZ"
'''


# dropdown_values_edit_page4=["String","Binary","DWORD(32-bit)","QWORD(64-bit)", "Multi-String", "Expandable String"]
def WindowEditData(title, path='', name='', reg_type='', reg_data=''):
    layout_edit = [
        [sg.Text('Registry Key/Subkey Path:', size=(20, 1)),
         sg.InputText(default_text=path, key='-PAGE4_EDIT_PATH-', readonly=True, size=(60, 1))],
        [sg.Text('Registry Name:', size=(20, 1)),
         sg.InputText(default_text=name, key='-PAGE4_EDIT_NAME-', readonly=True, size=(60, 1))],
        [sg.Text('Registry Type:', size=(20, 1)),
         sg.Combo(['String', 'Binary', 'DWORD(32-bit)', 'QWORD(64-bit)', 'Multi-String', 'Expandable String'],
                  default_value=reg_type, key="-PAGE4_EDIT_TYPE_DROPDOWN-", readonly=True),
         sg.Button('Select', key="-PAGE4_SELECT_FORMAT-")],
        [sg.pin(sg.Text('DWORD/QWORD Format:', size=(20, 1), key="-PAGE4_FORMAT-", visible=False)),
         sg.pin(sg.Radio('Decimal', "RADIO1", key="-PAGE4_FORMAT_DECIMAL-", default=True, visible=False)),
         sg.pin(sg.Radio('Hexadecimal', "RADIO1", key="-PAGE4_FORMAT_HEX-", visible=False))],
        [sg.Text('Registry Data:', size=(20, 1)),
         sg.InputText(default_text=reg_data, key="-PAGE4_EDIT_DATA-", size=(60, 1))],
        [sg.Push(), sg.Button("Save", key="-PAGE4_EDIT_SAVE-", disabled=True),
         sg.Button('Cancel', key="-PAGE4_EDIT_CANCEL-")]
    ]

    return sg.Window(title, layout_edit, finalize=True, size=(600, 200), resizable=True)

def delete_temp_files(folder_path):
    temp_files = glob.glob(os.path.join(folder_path, "*.temp"))
    for file in temp_files:
        try:
            os.remove(file)  # Delete the .temp file
        except Exception as e:
            print(f"Error deleting {file}: {e}")

require_password_for_tab5 = True
window_active = True

while True:
    event, values = window.read()

    # Detect attempt to close the application
    if event == sg.WIN_CLOSED or event == sg.WIN_CLOSE_ATTEMPTED_EVENT:
        # Check if there are unsaved changes
        if table_changed or window['-SAVE_GOLDEN_FILE-'].metadata:
            confirm_save = sg.popup_yes_no("You have unsaved changes. Do you want to save before exiting?",
                                           title="Unsaved Changes")
            if confirm_save == "Yes":
                save_status = makeWinSave("Save Registry Changes")  # Open save summary page
                if save_status == "saved" or save_status == "discarded":
                    table_changed = False
                    window['-SAVE_GOLDEN_FILE-'].metadata = False
                    window['-SAVE_GOLDEN_FILE-'].update(disabled=True)
                    window['-BROWSE2-'].update(disabled=False)  # Re-enable Browse button
            elif confirm_save == "No":
                delete_temp_files("data")  # Clean up temporary files
                window.close()  # Close the application
                break
        else:
            confirm_exit = sg.popup_yes_no("Are you sure you want to close the application?", title="Close Confirmation")
            if confirm_exit == "Yes":
                delete_temp_files("data")
                window.close()
                break
        continue

    elif event == "-DISABLE_PASSWORD-":
        # Handle toggle for disabling/enabling Page 4 password
        if values["-DISABLE_PASSWORD-"]:
            entered_password = sg.popup_get_text("Enter password to disable Page 4 password:", password_char="*")
            if entered_password == "$ViTrox$":
                require_password_for_tab5 = False
                window["-PASSWORD_STATUS-"].update("Status: Page 4 password is disabled.")
            else:
                sg.popup_error("Incorrect password! Page 4 password remains required.")
                window["-DISABLE_PASSWORD-"].update(value=False)
        else:
            # Re-enable password requirement for Page 4
            require_password_for_tab5 = True
            window["-PASSWORD_STATUS-"].update("Status: Page 4 password is required.")

    elif event == "-TABS-":  # Detects when a tab is selected
        selected_tab = values["-TABS-"]

        # If there are unsaved changes on Page 4, prompt the user
        if (table_changed or window['-SAVE_GOLDEN_FILE-'].metadata) and selected_tab != "-TAB5-":
            confirm_save = sg.popup_yes_no(
                "You have unsaved changes. Do you want to save before switching tabs?",
                title="Unsaved Changes"
            )

            if confirm_save == "Yes":
                save_status = makeWinSave("Save Registry Changes")  # Open save summary page
                if save_status in ["saved", "discarded"]:
                    table_changed = False
                    window['-SAVE_GOLDEN_FILE-'].metadata = False
                    window['-SAVE_GOLDEN_FILE-'].update(disabled=True)
                    window['-BROWSE2-'].update(disabled=False)  # Re-enable Browse button
                # Proceed to the selected tab
                window["-TABS-"].update(selected_tab)

            elif confirm_save == "No":
                delete_temp_files("data")  # Clean up temporary files
                table_changed = False
                window['-SAVE_GOLDEN_FILE-'].metadata = False
                window['-SAVE_GOLDEN_FILE-'].update(disabled=True)
                # Proceed to the selected tab
                window["-TABS-"].update(selected_tab)

            else:
                # User closed the pop-up without action; stay on Page 4
                window["-TABS-"].Widget.select(4)  # Assuming 4 is Page 4's index
                continue  # Skip to the next event iteration to avoid tab change

        # Check if the selected tab is -TAB5-
        elif selected_tab == "-TAB5-":
            delete_temp_files("data")
            reset_page4_state(window)

            # Only prompt for a password if the toggle is off (password required)
            if require_password_for_tab5:
                password = sg.popup_get_text("Enter password to access this tab:", password_char="*")
                if password == "$ViTrox$":
                    print("Access granted to Page 4: Registry Editor.")
                    delete_temp_files("data")
                    reset_page4_state(window)
                else:
                    sg.popup_error("Incorrect password! Access denied.")
                    # Redirect back to another tab if the password is incorrect
                    window["-TABS-"].Widget.select(0)  # Redirects to the first tab as an example
                    window["-STATUS-"].update("Access denied to Page 4: Registry Editor.")
            else:
                print("Password requirement is disabled for Page 4. Access granted.")

        else:
            # Update status with the currently selected tab for confirmation
            window["-STATUS-"].update(f"Current Tab: {selected_tab}")


        # page 1 for sw check
    if event == "-CHECK-":
        run_main_21()
        table_data = window["-TABLE-"].get()
        original_table_data = table_data
        update_gui()
        update_page1_data = True
        # refresh_window()

    # Page 2 for size check
    if event == "-CHECK_SIZE-":
        # Manual folder selection by user
        folder_path = sg.popup_get_folder("Select 'GNU' Folder for Size Checking", no_window=True)
        if folder_path:
            window["-FOLDER-"].update(folder_path)
            run_size_check(folder_path)  # Use the existing function to process the folder
            table_data = window["-SIZE_TABLE-"].get()
            original_size_table_data = table_data
            update_page2_data = True
        else:
            sg.popup_error("Error: 'GNU' folder selection cancelled or missing.", title="Error")

    elif event == "-CHECK_GNU_PATH-":
        # Automatic path selection without user input
        paths_to_check = ["C:\\Tools\\GNU", "C:\\Tools"]
        selected_path = None

        # Find the first existing path
        for path in paths_to_check:
            if os.path.exists(path):
                selected_path = path
                break

        # Update input field and run size check if a path is found
        if selected_path:
            window["-FOLDER-"].update(selected_path)
            sg.popup("Path found:", selected_path, title="Information")
            run_size_check(selected_path)
            table_data = window["-SIZE_TABLE-"].get()
            original_size_table_data = table_data
            update_page2_data = True
        else:
            sg.popup("GNU folder is not located at default path, please browse the folder manually.", title="Information")

    # page 3 for registry checker
    if event == 'Browse':

        if values["Dropdown"] == "MicroAOI":
            # file_path="C:/Users/xiau-yen.kelly-teo/Desktop/SW & Size checker - Copy/sample-MicroAOI.json"
            join_path = os.path.join("Golden File", "sample-MicroAOI.json")
            file_path = get_current_file_path(join_path)

            if file_path:
                # file_path=get_current_file_path('sample-MicroAOI.json')
                print(f"this file path have been chosen{file_path}")
                file_name = Path(file_path).stem
                sg.popup_ok(file_name + " has been selected.", title="Information")
                write_into_event_log(f"[Page 3] This file path have been chosen: {file_path}")
                window.find_element('_compare_').Update(visible=True)

            else:
                write_into_error_log(f"Error finding the selected machine type file.")
                sg.popup_error("Error finding the selected machine type file.", title="Error")

        elif values["Dropdown"] == "Semicon":
            # file_path="C:/Users/xiau-yen.kelly-teo/Desktop/SW & Size checker - Copy/sample-semicon.json"
            join_path = os.path.join("Golden File", "sample-semicon.json")
            file_path = get_current_file_path(join_path)

            if file_path:
                # file_path=get_current_file_path('sample-semicon.json')
                print(f"this file path have been chosen{file_path}")
                file_name = Path(file_path).stem
                sg.popup_ok(file_name + " has been selected.", title="Information")
                write_into_event_log(f"[Page 3] This file path have been chosen: {file_path}")
                window.find_element('_compare_').Update(visible=True)

            else:
                write_into_error_log(f"Error finding the selected machine type file.")
                sg.popup_error("Error finding the selected machine type file.", title="Error")

        elif values["Dropdown"] == "SideCam":
            # file_path="C:/Users/xiau-yen.kelly-teo/Desktop/SW & Size checker - Copy/sample-SideCam.json"
            join_path = os.path.join("Golden File", "sample-SideCam.json")
            file_path = get_current_file_path(join_path)

            if file_path:
                # file_path=get_current_file_path('sample-SideCam.json')
                print(f"this file path have been chosen{file_path}")
                file_name = Path(file_path).stem
                sg.popup_ok(file_name + " has been selected.", title="Information")
                write_into_event_log(f"[Page 3] This file path have been chosen: {file_path}")
                window.find_element('_compare_').Update(visible=True)

            else:
                write_into_error_log(f"Error finding the selected machine type file.")
                sg.popup_error("Error finding the selected machine type file.", title="Error")

        elif values["Dropdown"] == "SMT":
            # file_path="C:/Users/xiau-yen.kelly-teo/Desktop/SW & Size checker - Copy/sample-SMT.json"
            join_path = os.path.join("Golden File", "sample-SMT.json")
            file_path = get_current_file_path(join_path)

            if file_path:
                # file_path=get_current_file_path('sample-SMT.json')
                print(f"this file path have been chosen{file_path}")
                file_name = Path(file_path).stem
                sg.popup_ok(file_name + " has been selected.", title="Information")
                write_into_event_log(f"[Page 3] This file path have been chosen: {file_path}")
                window.find_element('_compare_').Update(visible=True)

            else:
                write_into_error_log(f"Error finding the selected machine type file.")
                sg.popup_error("Error finding the selected machine type file.", title="Error")

        else:
            file_path = sg.popup_get_file("Select 'REGISTRY' file for Registry Checking", title="File selector",
                                          no_window=True, file_types=(('JSON', '*.json'),))

            if file_path == '':
                sg.popup_ok("No file is select")
                write_into_event_log(f"No file is selected.")

            else:
                if file_path:
                    # file_path=get_current_file_path('sample-SMT.json')
                    print(f"this file path have been chosen{file_path}")
                    file_name = Path(file_path).stem
                    sg.popup_ok(file_name + " has been selected.", title="Information")
                    write_into_event_log(f"[Page 3] This file path have been chosen: {file_path}")
                    window.find_element('_compare_').Update(visible=True)

                else:
                    write_into_error_log(f"Error finding the selected machine type file.")
                    sg.popup_error("Error finding the selected machine type file.", title="Error")

        # refresh_window()

        # Page 4 for Registry Editor
        if event == 'BrowsePage4':

            file_path = sg.popup_get_file("Select 'REGISTRY' file for Editing", title="File selector",
                                          no_window=True, file_types=(('JSON', '*.json'),))

            if file_path == '':
                sg.popup_ok("No file is selected")
                write_into_event_log(f"No file is selected on Page 4: Registry Editor.")

            else:
                if file_path:
                    print(f"this file path has been chosen: {file_path}")
                    file_name = Path(file_path).stem
                    sg.popup_ok(file_name + " has been selected for editing.", title="Information")
                    write_into_event_log(f"[Page 4: Registry Editor] This file path has been chosen: {file_path}")
                    window.find_element('_edit_').Update(visible=True)  # Assume '_edit_' button becomes visible

                else:
                    write_into_error_log(f"Error finding the selected file on Page 4: Registry Editor.")
                    sg.popup_error("Error finding the selected file for editing.", title="Error")

    set_file_path(file_path)
    # selected_file_output.update(value=os.path.basename(file_path))

    if event == '_compare_':
        selected_machine_type = values["Dropdown"]

        # Check if the file exists before proceeding
        if not os.path.isfile(file_path):
            sg.popup_error(f"The selected golden file '{file_path}' does not exist.")
            continue  # Skip further processing if file is missing

        # Show the loading window as a modal
        loading_layout = [[sg.Text("Loading...", font=("Helvetica", 16), justification="center")]]
        loading_window = sg.Window("Please Wait", loading_layout, no_titlebar=True, keep_on_top=True, modal=True,
                                   finalize=True)
        loading_window.read(timeout=0)

        # Run functions to refresh JSON files needed for counters
        run_registry_main(file_path)  # Updates result_reg_data.json
        run_registry_main2(file_path)  # Updates compared_result_reg_data.json
        run_registry_main3(file_path)  # Updates redundant_data.json

        # Close the loading window after tasks are complete
        loading_window.close()

        # Refresh counters and table display in the UI
        update_reg_gui()
        window.find_element('-SEARCH_REG-').Update(visible=True)
        window.find_element("-SEARCH_TEXT-").update(visible=True)
        window.find_element('ViewSummary').Update(visible=True)

        # Close the loading window after tasks are complete
        loading_window.close()

        # Refresh counters and table display in the UI
        update_reg_gui()
        window.find_element('-SEARCH_REG-').Update(visible=True)
        window.find_element("-SEARCH_TEXT-").update(visible=True)
        window.find_element('ViewSummary').Update(visible=True)

        # refresh_window()


    elif event == 'Export':

        ch = sg.popup_yes_no('Do you want to export the current registry?', title='Export Registry Keys')

        if ch == 'Yes':

            # Get current date and time to append to the file name
            current_time = datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')
            default_file_name = f"current_pc_registry_data_exported_{current_time}"

            # File save dialog with default file name
            export_file_path = sg.popup_get_file(
                'Save As',
                title='Save Registry As',
                save_as=True,
                no_window=True,
                default_path=f"{default_file_name}",
                file_types=(('Registry Files', '*.reg'), ('JSON Files', '*.json'))
            )

            if export_file_path != '':
                print(f"File selected: {export_file_path}")

                # Check the file extension to determine the export type
                if export_file_path.endswith('.reg'):
                    # Export as .reg (registry export logic)
                    sg.popup_ok(f"Exporting registry as .reg file to {export_file_path}", title="Export Successful")
                    command = ['reg', 'export', r'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\MV Technology',
                               export_file_path, '/y']
                    subprocess.run(command, capture_output=True, text=True, check=True)
                    print(f"Registry exported successfully to {export_file_path}")


                elif export_file_path.endswith('.json'):
                    # Export as .json (current JSON export logic)
                    sg.popup_ok(f"Exporting registry as .json file to {export_file_path}", title="Export Successful")
                    merged_current_registry_data("current_pc_registry_data_exported.json")
                    exported_file = "current_pc_registry_data_exported.json"
                    shutil.move(exported_file, export_file_path)  # Save the exported JSON to the selected path

            else:
                sg.popup_ok("No folder or file is selected", title="Cancel")

        else:
            sg.popup_ok("No file is exported", title="Cancel")

    # refresh_window()

    elif event == 'ViewSummary':

        loading_layout = [[sg.Text("Loading...", font=("Helvetica", 16), justification="center")]]
        loading_window = sg.Window("Please Wait", loading_layout, no_titlebar=True, keep_on_top=True, modal=True,
                                   finalize=True)

        loading_window.read(timeout=100)

        # refresh_window()
        window_view_more_active = True
        window_view_more = makeWin2('Summary Page')

        refresh_table()
        refresh_table2(window_view_more, event, values)
        update_page3_data = True
        update_page4_upper_data = True
        update_page4_bottom_data = True

        '''Summary Page'''
        window_view_more.find_element('-SEARCH_REG2-').update(visible=True)
        window_view_more.find_element("-SEARCH_TEXT2-").update(visible=True)
        window_view_more.find_element("-DELETE-").update(disabled=True)

        '''Update Button'''
        json_file_path = get_current_file_path('data/compared_result_reg_data.json')
        # Read and load the JSON file
        with open(json_file_path, 'r') as json_file:
            try:
                results = json.load(json_file)
            except json.JSONDecodeError:
                write_into_error_log(content="The file '" + json_file_path + "' contains invalid JSON data.")
                raise ValueError(f"The file {json_file_path} contains invalid JSON data.")

        updated_reg_compare_data = [
            [
                BLANK_BOX,
                compared_result_reg_data["Registry Key/Subkey Path"],
                compared_result_reg_data["Registry Name"],
                compared_result_reg_data["Current Type"],
                compared_result_reg_data["Expected Type"],
                compared_result_reg_data["Current Data"],
                compared_result_reg_data["Expected Data"],
                compared_result_reg_data["Status"]
            ]
            for i, compared_result_reg_data in enumerate(results.values())
        ]

        filtered_data = []
        for row in updated_reg_compare_data:
            _, _, _, _, _, _, _, status = row
            if status == "Fail" or status == "Missing":
                filtered_data.append(row)

            if filtered_data:
                window_view_more.find_element("-update-").update(disabled=False)

            else:
                window_view_more.find_element("-update-").update(disabled=True)

        '''Restore Page'''
        # check the redundant data in deleted folder have data or not
        # if got, then visible the restore button
        get_deleted_redundant_path = "data/list_of_deleted_keys.json"
        os.makedirs(os.path.dirname(get_deleted_redundant_path), exist_ok=True)
        if not os.path.isfile(get_deleted_redundant_path):
            write_to_json([], get_deleted_redundant_path)
        get_full_current_path = get_current_file_path(get_deleted_redundant_path)
        get_deleted_redundant_data = load_registry_from_json2(get_full_current_path)

        if get_deleted_redundant_data:
            window_view_more.find_element("-RESTORE-").update(disabled=False)

        else:
            window_view_more.find_element("-RESTORE-").update(disabled=True)

        loading_window.close()

        while True:
            event3, values3 = window_view_more.Read()

            if event3 == sg.WIN_CLOSED:
                window_view_more.Close()
                window_view_more_active = False

                # Reapply sorting and colors after closing summary view
                refresh_table()
                # clear the data in selected_registry_data.json file
                file_path_selected_reg = get_selected_reg_file_path()
                output_file = get_current_file_path(file_path_selected_reg)
                write_to_json([], output_file)

                break

            if event3 == "-updateSelected-":

                # Confirmation dialog
                ch = sg.popup_yes_no('Do you want to update the selected registry keys?', title="Confirmation")
                if ch == 'Yes':

                    # Show loading pop-up
                    loading_layout = [[sg.Text("Loading...", font=("Helvetica", 16), justification="center")]]
                    loading_window = sg.Window("Please Wait", loading_layout, no_titlebar=True, keep_on_top=True,
                                               modal=True, finalize=True)
                    loading_window.read(timeout=0)

                    try:
                        # Backup the JSON file before import
                        exported_file = "current_pc_registry_data.json"
                        backup_folder = "Backup\\Page 3\\Fail_Missing\\Backup(Selected)"
                        backup_registry(exported_file, backup_folder)

                        # Get the selected registry keys path
                        selected_reg_file = get_selected_reg_file_path()
                        selected_registry_data = load_registry_from_json2(selected_reg_file)

                        if selected_registry_data:
                            import_selected_registry_result(selected_registry_data)
                            print(f"{selected_reg_file}")
                            refresh_table2(window_view_more, event3, values3)
                            window_view_more.find_element("-updateSelected-").update(visible=False)
                            window_view_more.find_element("-update-").update(visible=True)
                            window_view_more.find_element("-update-").update(disabled=False)

                            # Close loading window before showing success message
                            loading_window.close()
                            sg.popup_auto_close('Successfully updated the selected keys.', title="Information")

                        else:
                            # Close loading window before showing failure message
                            loading_window.close()
                            refresh_table2(window_view_more, event3, values3)
                            sg.popup_auto_close('Failed to update the selected keys.', title="Error")

                        # Set update flags
                        update_page3_data = True
                        update_page4_upper_data = True
                        update_page4_bottom_data = True

                    finally:
                        # Ensure loading window is closed in case of unexpected error
                        loading_window.close()

                else:
                    sg.popup_ok("No key is updated.")

            if event3 == "_importBackup_":

                ch = sg.popup_ok_cancel("Do you want to import backup file?", title="Import Backup File")

                if ch == "OK":

                    # Show loading pop-up
                    loading_layout = [[sg.Text("Loading...", font=("Helvetica", 16), justification="center")]]
                    loading_window = sg.Window("Please Wait", loading_layout, no_titlebar=True, keep_on_top=True,
                                               modal=True, finalize=True)
                    loading_window.read(timeout=0)

                    try:
                        # File selector for the backup file
                        backup_file_path = sg.popup_get_file("Select a backup file for importing",
                                                             title="File selector",
                                                             no_window=True, file_types=(("JSON", "*json"),))

                        if backup_file_path:
                            # Load the registry data from the selected backup file
                            print(f"{backup_file_path}")
                            backup_registry_key = load_registry_from_json2(backup_file_path)

                            if backup_registry_key:
                                import_registry(backup_registry_key)
                                refresh_table2(window_view_more, event3, values3)
                                window_view_more.find_element("-update-").update(disabled=False)

                                # Close loading pop-up before showing the success message
                                loading_window.close()

                                sg.popup_ok(
                                    'Successfully imported file.\n* Registry key(s) that contain PASSWORD will not be replaced.')

                            else:
                                # Close loading pop-up if no file found
                                loading_window.close()
                                sg.popup_auto_close('No file found.')

                        else:
                            # Close loading pop-up if no file selected
                            loading_window.close()
                            sg.popup_ok("No file is selected")

                    finally:
                        # Ensure loading window is closed in case of any unexpected error
                        loading_window.close()

                else:
                    sg.popup_ok("No key is imported", title="Cancel")

            if event3 == "-EDIT-":

                selected_redundant_file = get_selected_redundant_file_path()
                selected_redundant_data = load_registry_from_json2(selected_redundant_file)

                data = [
                    [

                        result_reg_data["Registry Key/Subkey Path"],
                        result_reg_data["Registry Name"],
                        result_reg_data["Type"],
                        result_reg_data["Data"]
                    ]
                    for i, result_reg_data in enumerate(selected_redundant_data)
                ]

                if len(selected_redundant_data) == 1:
                    for path, name, reg_type, reg_data in data:
                        window_edit_active = True
                        if reg_type == "REG_SZ":
                            reg_type = "String"
                        elif reg_type == "REG_BINARY":
                            reg_type = "Binary"
                        elif reg_type == "REG_DWORD_LITTLE_ENDIAN":
                            reg_type = "DWORD(32-bit)"
                        elif reg_type == "REG_QWORD_LITTLE_ENDIAN":
                            reg_type = "QWORD(64-bit)"
                        elif reg_type == "REG_MULTI_SZ":
                            reg_type = "Multi-String"
                        elif reg_type == "REG_EXPAND_SZ":
                            reg_type = "Expandable String"
                        elif name == "Default":
                            if reg_data == "":
                                reg_data = "(value no set)"

                        windowedit = makeWinEdit('Edit registry key', path, name, reg_type, reg_data)

                while True:

                    event_edit, values_edit = windowedit.Read()

                    if event_edit == sg.WIN_CLOSED:
                        windowedit.Close()
                        window_edit_active = False
                        break

                    if event_edit == "-SAVE-":

                        ch = sg.popup_yes_no("Do you want to save changes?", title="Confirmation")
                        if ch == 'Yes':

                            loading_layout = [
                                [sg.Text("Saving changes...", font=("Helvetica", 16), justification="center")]]
                            loading_window = sg.Window("Please Wait", loading_layout, no_titlebar=True,
                                                       keep_on_top=True,
                                                       modal=True, finalize=True)
                            loading_window.read(timeout=0)

                            try:

                                # backup the json file before import
                                exported_file = "current_pc_registry_data.json"
                                backup_folder = "Backup\Page 3\Redundant\Backup(Edited)"
                                backup_registry(exported_file, backup_folder)

                                selected_redundant_file = get_selected_redundant_file_path()
                                selected_redundant_data = load_registry_from_json2(selected_redundant_file)

                                if selected_redundant_data:
                                    edited_key = edit_selected_row(selected_redundant_data, windowedit)

                                if edited_key:
                                    import_registry(edited_key)
                                    refresh_table2(window_view_more, event3, values3)

                            finally:
                                loading_window.close()
                                time.sleep(0.2)

                            sg.popup_ok("Changes saved and imported.", title="Information")

                        else:
                            sg.popup_ok("No changes made.", title="Information")
                            window_view_more.find_element("-EDIT-").update(disabled=True)
                            window_view_more.find_element("-DELETE-").update(disabled=True)

                        windowedit.close()
                        window_edit_active = False
                        break

                    if event_edit == "-CANCEL-":
                        windowedit.Close()
                        window_edit_active = False
                        break

                    if event_edit == "-SELECT_FORMAT-":
                        if values_edit["-EDIT_TYPE_DROPDOWN-"] == "DWORD(32-bit)":
                            windowedit.find_element("-FORMAT-").update(visible=True)
                            windowedit.find_element("-FORMAT_DECIMAL-").update(visible=True)
                            windowedit.find_element("-FORMAT_HEX-").update(visible=True)
                            windowedit.find_element("-SAVE-").update(disabled=False)

                        elif values_edit["-EDIT_TYPE_DROPDOWN-"] == "QWORD(64-bit)":
                            windowedit.find_element("-FORMAT-").update(visible=True)
                            windowedit.find_element("-FORMAT_DECIMAL-").update(visible=True)
                            windowedit.find_element("-FORMAT_HEX-").update(visible=True)
                            windowedit.find_element("-SAVE-").update(disabled=False)

                        elif values_edit["-EDIT_TYPE_DROPDOWN-"] == "String":
                            windowedit.find_element("-FORMAT-").update(visible=False)
                            windowedit.find_element("-FORMAT_DECIMAL-").update(visible=False)
                            windowedit.find_element("-FORMAT_HEX-").update(visible=False)
                            windowedit.find_element("-SAVE-").update(disabled=False)

                        elif values_edit["-EDIT_TYPE_DROPDOWN-"] == "Binary":
                            windowedit.find_element("-FORMAT-").update(visible=False)
                            windowedit.find_element("-FORMAT_DECIMAL-").update(visible=False)
                            windowedit.find_element("-FORMAT_HEX-").update(visible=False)
                            windowedit.find_element("-SAVE-").update(disabled=False)

                        elif values_edit["-EDIT_TYPE_DROPDOWN-"] == "Multi-String":
                            windowedit.find_element("-FORMAT-").update(visible=False)
                            windowedit.find_element("-FORMAT_DECIMAL-").update(visible=False)
                            windowedit.find_element("-FORMAT_HEX-").update(visible=False)
                            windowedit.find_element("-SAVE-").update(disabled=False)

                        elif values_edit["-EDIT_TYPE_DROPDOWN-"] == "Expandable String":
                            windowedit.find_element("-FORMAT-").update(visible=False)
                            windowedit.find_element("-FORMAT_DECIMAL-").update(visible=False)
                            windowedit.find_element("-FORMAT_HEX-").update(visible=False)
                            windowedit.find_element("-SAVE-").update(disabled=False)

                window_view_more.find_element("-EDIT-").update(disabled=False)
                window_view_more.find_element("-DELETE-").update(disabled=True)

            if event3 == "-DELETE-":
                ch = sg.popup_yes_no("Do you want to DELETE the selected registry keys?")
                if ch == "Yes":

                    loading_layout = [
                        [sg.Text("Deleting selected keys...", font=("Helvetica", 16), justification="center")]]
                    loading_window = sg.Window("Please Wait", loading_layout, no_titlebar=True, keep_on_top=True,
                                               modal=True, finalize=True)
                    loading_window.read(timeout=0)

                    try:
                        # Load the selected redundant file data
                        selected_redundant_file = get_selected_redundant_file_path()
                        selected_redundant_data = load_registry_from_json2(selected_redundant_file)

                        # Backup the selected keys to a new JSON file before deletion
                        exported_file = 'data/selected_redundant_data.json'
                        backup_folder = "data"
                        backup_deleted_registry(exported_file, backup_folder)

                        # Backup the current registry data file before deletion
                        exported_file2 = "current_pc_registry_data.json"
                        backup_folder2 = "Backup\Page 3\Redundant\Backup(Deleted)"
                        backup_registry(exported_file2, backup_folder2)

                        if selected_redundant_data:
                            # Delete redundant keys and refresh the table
                            delete_redundant_keys(selected_redundant_data)
                            print(f"{selected_redundant_data}")
                            refresh_table2(window_view_more, event3, values3)

                            # Disable edit and delete buttons and enable restore button
                            window_view_more.find_element('-EDIT-').update(disabled=True)
                            window_view_more.find_element("-DELETE-").update(disabled=True)
                            window_view_more.find_element('-RESTORE-').update(disabled=False)

                            # Set the success message
                            message = 'Successfully deleted selected keys.'

                        else:
                            # Set the failure message
                            message = 'Failed to import the selected keys.'

                            # Trigger flags for updating data on other pages
                            update_page3_data = True
                            update_page4_upper_data = True
                            update_page4_bottom_data = True

                        # Refresh the table in both cases
                        refresh_table2(window_view_more, event3, values3)

                    finally:
                        # Close the loading popup
                        loading_window.close()

                        # Add a slight delay to prevent popup overlap
                        time.sleep(0.2)

                        # Display the success or failure message after the delay
                    sg.popup_ok(message, title="Information")

            # window_view_more.find_element('-EDIT-').update(disabled=True)
            if event3 == "-RESTORE-":

                window_restore_active = True
                windowRestore = makeWinRestore('Restore Page')
                update_restore_gui(windowRestore)
                update_restore_page_data = True

                while True:

                    event_restore, values_restore = windowRestore.Read()

                    if event_restore == sg.WIN_CLOSED:
                        windowRestore.Close()
                        window_restore_active = False
                        break

                    if event_restore == "-restoreSelected-":
                        ch = sg.popup_yes_no('Do you want to restore the selected registry keys?', title="Confirmation")
                        if ch == 'Yes':
                            # Show loading popup before starting the restore process
                            loading_layout = [
                                [sg.Text("Restoring selected keys...", font=("Helvetica", 16), justification="center")]]
                            loading_window = sg.Window("Please Wait", loading_layout, no_titlebar=True,
                                                       keep_on_top=True,
                                                       modal=True, finalize=True)
                            loading_window.read(timeout=0)

                            try:
                                # Backup the JSON file before import
                                exported_file = "current_pc_registry_data.json"
                                backup_folder = "Backup\Page 3\Redundant\Backup(Restore)"
                                backup_registry(exported_file, backup_folder)

                                # Get the selected registry keys to restore
                                selected_reg_file = get_selected_restore_file_path()
                                selected_registry_data = load_registry_from_json2(selected_reg_file)
                                print(f"{selected_registry_data}")

                                if selected_registry_data:
                                    # Restore selected registry keys
                                    restore_selected_registry_result(selected_registry_data)
                                    print(f"{selected_registry_data}")

                                    # Delete the restore keys in the 'list_of_deleted_key' after restore
                                    get_deleted_redundant_path = os.path.join("data", "list_of_deleted_keys.json")
                                    json_file_path = get_current_file_path(get_deleted_redundant_path)
                                    filter_out_matching_entries(get_deleted_redundant_path, selected_reg_file,
                                                                get_deleted_redundant_path)

                                    # Refresh GUI
                                    refresh_table2(window_view_more, event3, values3)
                                    update_restore_gui(windowRestore)
                                    window_view_more.find_element("-EDIT-").update(disabled=True)
                                    window_view_more.find_element("-DELETE-").update(disabled=True)

                                    # Set the success message
                                    message = 'Successfully restored selected keys.'
                                else:
                                    # Set the failure message
                                    message = 'Failed to restore the selected keys.'

                                    # Update page data flag
                                    update_restore_page_data = True

                            finally:
                                # Close the loading popup
                                loading_window.close()

                                # Add a slight delay to prevent popup overlap
                                time.sleep(0.2)

                            # Display the success or failure message after the delay
                            sg.popup_ok(message, title="Information")

                        else:
                            sg.popup_ok("No key is imported.")

                        # Additional restore-related checks
                        get_deleted_redundant_path = "data/list_of_deleted_keys.json"
                        get_full_current_path = get_current_file_path(get_deleted_redundant_path)
                        get_deleted_redundant_data = load_registry_from_json2(get_full_current_path)

                        if get_deleted_redundant_data:
                            window_view_more.find_element("-RESTORE-").update(disabled=False)
                        else:
                            window_view_more.find_element("-RESTORE-").update(disabled=True)

                    # Other restore-related event handling
                    if event_restore in ("-SEARCH_BUTTON_RESTORE-", "\r", "-SEARCH_RESTORE-"):
                        search_text_restore_page = values_restore["-SEARCH_RESTORE-"].strip().lower()
                        perform_reg_search3(search_text_restore_page, windowRestore, event_restore, values_restore)

                        if not search_text_page4:
                            update_restore_gui(windowRestore)

                    if update_restore_page_data:
                        restore_result = handle_table_selection3(windowRestore, event_restore, values_restore)

                        if restore_result:
                            results_output_file = "data\selected_restore_data.json"
                            write_to_json(restore_result, results_output_file)

                        sort_restore_reg_table(windowRestore, event_restore, col_count=5,
                                               current_sort_order=current_sort_order_table4)

            ##import transfer to view more window
            if event3 == '-update-':

                # refresh_window()
                windowImport_active = True
                # windowImport = sg.Window('Import Registries', layoutImport, finalize=True, size = (500, 70), resizable=True)
                windowImport = makeWin('Update/Add the Registry Keys')

                while True:
                    event2, values2 = windowImport.Read()

                    if event2 == sg.WIN_CLOSED:
                        windowImport.Close()
                        windowImport_active = False
                        break

                    if event2 == '_updateAll_':

                        windowImport.Close()
                        windowImport_active = False
                        ch = sg.popup_ok_cancel("Do you want to update ALL(fail and missing) keys?",
                                                title="Update ALL registry keys")

                        if ch == "OK":
                            loading_layout = [[sg.Text("Updating all keys...", font=("Helvetica", 16), justification="center")]]
                            loading_window = sg.Window("Please Wait", loading_layout, no_titlebar=True, keep_on_top=True, modal=True, finalize=True)
                            loading_window.read(timeout=0)

                            # backup the json file before import
                            exported_file = "current_pc_registry_data.json"
                            backup_folder = "Backup\Page 3\Fail_Missing\Backup(All)"
                            backup_registry(exported_file, backup_folder)

                            try:
                                installed_registry = read_installed_registry()  # Replace with your function to read installed registry keys
                                # json_registry = load_registry_from_json2('sample-registry.json')  # Load JSON file
                                # file_path=get_sample_file()
                                current_path = get_file_path()
                                json_registry = load_registry_from_json2(current_path)

                                if not installed_registry:
                                    raise ValueError("Failed to retrieve installed registry.")

                                failed_missing_keys = []
                                failed_missing_keys = compare_registries(installed_registry, json_registry)

                                if failed_missing_keys:
                                    import_registry(failed_missing_keys)
                                    refresh_table2(window_view_more, event3, values3)
                                    window_view_more.find_element("-update-").update(disabled=True)
                                else:
                                    print("No keys found.")
                                    refresh_table2(window_view_more, event3, values3)
                                    sg.popup_error('No keys found.', title="Error")


                            except Exception as e:
                                print(f"Error updating the registry keys: {e}")
                                write_into_error_log(f"Error updating the registry keys: {e}")

                            finally:
                                loading_window.close()
                                time.sleep(0.2)

                                sg.popup_auto_close('Successfully updated the Failed and Missing keys.', title="Information")

                        else:
                            sg.popup_ok("No key is updated", title="Cancel")


                    elif event2 == '_updateFail_':

                        windowImport.Close()
                        windowImport_active = False
                        ch = sg.popup_ok_cancel("Do you want to update FAIL keys?", title="Update Fail Registry Keys")

                        if ch == "OK":
                            loading_layout = [[sg.Text("Updating Fail keys...", font=("Helvetica", 16), justification="center")]]
                            loading_window = sg.Window("Please Wait", loading_layout, no_titlebar=True, keep_on_top=True, modal=True, finalize=True)
                            loading_window.read(timeout=0)

                            # Backup the JSON file before import
                            exported_file = "current_pc_registry_data.json"
                            backup_folder = "Backup\Page3\Fail_Missing\Backup(Fail)"
                            backup_registry(exported_file, backup_folder)

                            try:
                                # Retrieve installed registry and JSON data
                                installed_registry = read_installed_registry()  # Replace with your function to read installed registry keys
                                current_path = get_file_path()
                                json_registry = load_registry_from_json2(current_path)

                                if not installed_registry:
                                    raise ValueError("Failed to retrieve installed registry.")

                                # Compare to find fail keys
                                fail_keys = compare_fail_registries(installed_registry, json_registry)

                                if fail_keys:
                                    # Import fail keys and set success message
                                    import_registry(fail_keys)
                                    print("Fail keys updated successfully.")
                                    refresh_table2(window_view_more, event3, values3)
                                    message = 'Successfully updated the Failed keys.'
                                    message_type = "info"

                                else:
                                    # Set error message if no fail keys found
                                    message = 'No Fail keys found.'
                                    message_type = "error"
                                    refresh_table2(window_view_more, event3, values3)

                            except Exception as e:
                                # Log and set error message on exception
                                print(f"Error updating fail registry: {e}")
                                write_into_error_log(f"Error updating fail registry: {e}")
                                message = f"Error updating fail registry: {e}"
                                message_type = "error"

                            finally:
                                # Close the loading popup
                                loading_window.close()
                                time.sleep(0.2)

                                # Display the appropriate message after the delay
                                if message_type == "info":
                                    sg.popup_auto_close(message, title="Information")
                                else:
                                    sg.popup_error(message, title="Error")

                        else:
                            sg.popup_ok("No key is updated", title="Cancel")


                    elif event2 == '_updateMissing_':

                        windowImport.Close()
                        windowImport_active = False
                        ch = sg.popup_ok_cancel("Do you want to add MISSING keys?", title="Add Missing Registry Keys")

                        if ch == "OK":
                            loading_layout = [[sg.Text("Updating Missing keys...", font=("Helvetica", 16), justification="center")]]
                            loading_window = sg.Window("Please Wait", loading_layout, no_titlebar=True, keep_on_top=True, modal=True, finalize=True)
                            loading_window.read(timeout=0)

                            # Backup the JSON file before import
                            exported_file = "current_pc_registry_data.json"
                            backup_folder = "Backup\Page3\Fail_Missing\Backup(Missing)"
                            backup_registry(exported_file, backup_folder)

                            try:
                                # Retrieve installed registry and JSON data
                                installed_registry = read_installed_registry()  # Replace with your function to read installed registry keys
                                current_path = get_file_path()
                                json_registry = load_registry_from_json2(current_path)

                                if not installed_registry:
                                    raise ValueError("Failed to retrieve installed registry.")

                                # Compare to find missing keys
                                missing_keys = compare_missing_registries(installed_registry, json_registry)
                                print(f"{missing_keys}")

                                if missing_keys:
                                    # Import missing keys and set success message
                                    import_registry(missing_keys)
                                    refresh_table2(window_view_more, event3, values3)
                                    message = 'Successfully added the Missing keys.'
                                    message_type = "info"

                                else:
                                    # Set error message if no missing keys found
                                    message = 'No Missing keys found.'
                                    message_type = "error"
                                    refresh_table2(window_view_more, event3, values3)

                            except Exception as e:
                                # Log and set error message on exception
                                print(f"Error adding the missing registry: {e}")
                                write_into_error_log(f"Error adding the missing registry: {e}")
                                message = f"Error adding the missing registry: {e}"
                                message_type = "error"

                            finally:
                                # Close the loading popup
                                loading_window.close()
                                time.sleep(0.2)

                                # Display the appropriate message after the delay
                                if message_type == "info":
                                    sg.popup_auto_close(message, title="Information")
                                else:
                                    sg.popup_error(message, title="Error")

                        else:
                            sg.popup_ok("No key is added", title="Cancel")

            if event3 in ("-SEARCH_BUTTON_REG2-", "\r", "-SEARCH_REG2-"):
                # For Page 3
                search_text_page4 = values3["-SEARCH_REG2-"].strip().lower()
                # table_data = window["-TABLE_REG_COMPARED-"].get()
                perform_reg_search2(search_text_page4, window_view_more, event3, values3)

                if not search_text_page4:
                    update_reg_compared_gui(window_view_more)

                    # refresh_window()

            if event3 in ("-SEARCH_BUTTON_REDUNDANT-", "\r", "-SEARCH_REDUNDANT-"):
                # For List of Redundant Keys
                search_text_redundant = values3["-SEARCH_REDUNDANT-"].strip().lower()
                perform_redundant_search(search_text_redundant, window_view_more, event3, values3)

                if not search_text_redundant:
                    update_redundant_gui(window_view_more)

                    # refresh_window()e

            if update_page4_upper_data:

                # for the table above
                selected_result = handle_table_selection(window_view_more, event3, values3)

                # Check if result has data before saving
                if selected_result:
                    # Define the path to save the selected key into the json file
                    results_output_file = "data\selected_registry_results.json"
                    write_to_json(selected_result, results_output_file)

                sort_compared_reg_table(window_view_more, event3, col_count=8,
                                        current_sort_order=current_sort_order_table2)

            if update_page4_bottom_data:

                # for the table bottom
                redundant_result = handle_table_selection2(window_view_more, event3, values3)

                if redundant_result:
                    # Define the output file path for saving the results
                    results_output_file = "data/selected_redundant_data.json"
                    write_to_json(redundant_result, results_output_file)

                sort_redundant_reg_table(window_view_more, event3, col_count=5,
                                         current_sort_order=current_sort_order_table3)

    # Page 4 event for registry editor
    if event == "-BROWSE2-":
        table_changed = False
        update_editor_data = True

        if values["Dropdown2"] == "MicroAOI":
            join_path = os.path.join("Golden File", "sample-MicroAOI.json")
            file_path = get_current_file_path(join_path)

            if file_path:
                # Use the golden file name to create the temp file
                golden_file_name = os.path.basename(file_path).replace(".json", "")
                temp_file_name = f"{golden_file_name}_golden_file.temp"
                temp_file_path = os.path.join("data", temp_file_name)

                # Copy the contents of the original golden file to the temp file
                with open(file_path, "r") as original_file:
                    data = original_file.read()

                with open(temp_file_path, "w") as temp_file:
                    temp_file.write(data)

                print(f"{temp_file_name} created successfully with all data.")
                file_name = os.path.basename(file_path)
                sg.popup_ok(file_name + " has been selected.", title="Information")
                write_into_event_log(f"[Page 4: Registry Editor] This file path has been chosen: {file_path}")
                update_editor_gui(file_path)

                window['-RESTORE_REV-'].update(disabled=False)  # Enable the restore button

        elif values["Dropdown2"] == "Semicon":
            join_path = os.path.join("Golden File", "sample-semicon.json")
            file_path = get_current_file_path(join_path)

            if file_path:
                golden_file_name = os.path.basename(file_path).replace(".json", "")
                temp_file_name = f"{golden_file_name}_golden_file.temp"
                temp_file_path = os.path.join("data", temp_file_name)

                with open(file_path, "r") as original_file:
                    data = original_file.read()

                with open(temp_file_path, "w") as temp_file:
                    temp_file.write(data)

                print(f"{temp_file_name} created successfully with all data.")
                file_name = os.path.basename(file_path)
                sg.popup_ok(file_name + " has been selected.", title="Information")
                write_into_event_log(f"[Page 4: Registry Editor] This file path has been chosen: {file_path}")
                update_editor_gui(file_path)

                window['-RESTORE_REV-'].update(disabled=False)

            else:
                write_into_error_log(f"Error finding the selected machine type file.")
                sg.popup_error("Error finding the selected machine type file.", title="Error")

        elif values["Dropdown2"] == "SideCam":
            join_path = os.path.join("Golden File", "sample-SideCam.json")
            file_path = get_current_file_path(join_path)

            if file_path:
                golden_file_name = os.path.basename(file_path).replace(".json", "")
                temp_file_name = f"{golden_file_name}_golden_file.temp"
                temp_file_path = os.path.join("data", temp_file_name)

                with open(file_path, "r") as original_file:
                    data = original_file.read()

                with open(temp_file_path, "w") as temp_file:
                    temp_file.write(data)

                print(f"{temp_file_name} created successfully with all data.")
                file_name = os.path.basename(file_path)
                sg.popup_ok(file_name + " has been selected.", title="Information")
                write_into_event_log(f"[Page 4: Registry Editor] This file path has been chosen: {file_path}")
                update_editor_gui(file_path)

                window['-RESTORE_REV-'].update(disabled=False)

            else:
                write_into_error_log(f"Error finding the selected machine type file.")
                sg.popup_error("Error finding the selected machine type file.", title="Error")

        elif values["Dropdown2"] == "SMT":
            join_path = os.path.join("Golden File", "sample-SMT.json")
            file_path = get_current_file_path(join_path)

            if file_path:
                golden_file_name = os.path.basename(file_path).replace(".json", "")
                temp_file_name = f"{golden_file_name}_golden_file.temp"
                temp_file_path = os.path.join("data", temp_file_name)

                with open(file_path, "r") as original_file:
                    data = original_file.read()

                with open(temp_file_path, "w") as temp_file:
                    temp_file.write(data)

                print(f"{temp_file_name} created successfully with all data.")
                file_name = os.path.basename(file_path)
                sg.popup_ok(file_name + " has been selected.", title="Information")
                write_into_event_log(f"[Page 4: Registry Editor] This file path has been chosen: {file_path}")
                update_editor_gui(file_path)

                window['-RESTORE_REV-'].update(disabled=False)

            else:
                write_into_error_log(f"Error finding the selected machine type file.")
                sg.popup_error("Error finding the selected machine type file.", title="Error")

        else:
            file_path = sg.popup_get_file("Select 'REGISTRY' file for Registry Checking", title="File selector",
                                          no_window=True, file_types=(('JSON', '*.json'),))
            if file_path == '':
                sg.popup_ok("No file is selected")
                write_into_event_log(f"No file is selected.")
            else:
                if file_path:
                    golden_file_name = os.path.basename(file_path).replace(".json", "")
                    temp_file_name = f"{golden_file_name}_golden_file.temp"
                    temp_file_path = os.path.join("data", temp_file_name)

                    with open(file_path, "r") as original_file:
                        data = original_file.read()

                    with open(temp_file_path, "w") as temp_file:
                        temp_file.write(data)

                    print(f"Custom temp file created with all data: {temp_file_name}")
                    file_name = os.path.basename(file_path)
                    sg.popup_ok(file_name + " has been selected.", title="Information")
                    write_into_event_log(f"[Page 4: Registry Editor] Custom file path has been chosen: {file_path}")
                    update_editor_gui(file_path)

                    window['-RESTORE_REV-'].update(disabled=False)

                else:
                    write_into_error_log(f"Error finding the selected machine type file.")
                    sg.popup_error("Error finding the selected machine type file.", title="Error")

        set_file_path(file_path)
        window.find_element("-ADD_NEW-").update(visible=True)
        window.find_element("-EDIT_GOLDEN_FILE-").update(visible=True)
        window.find_element("-DELETE_FROM_GOLDEN_FILE-").update(visible=True)
        window.find_element("-SAVE_GOLDEN_FILE-").update(visible=True)


        # You can still check if there are deleted items and update the UI accordingly (without popping up the window)
        deleted_data_file = "data/deleted_registry_data_page4.json"
        deleted_data = load_registry_from_json2(deleted_data_file)

        # Update buttons based on whether there are deleted items, but don't show any window
        '''
        if deleted_data:
            window['-RESTORE_PAGE4-'].update(disabled=False)
        else:
            window['-RESTORE_PAGE4-'].update(disabled=True)
        '''

    if event == "-ADD_NEW-":
        write_into_event_log("User opened the 'Add new data' window on Page 4: Registry Editor.")
        window_add_data_active = True
        windowAddData = WindowAdd("Add new data")

        while True:
            eventAdd, valuesAdd = windowAddData.Read()

            if eventAdd == sg.WIN_CLOSED:
                windowAddData.Close()
                window_add_data_active = False
                write_into_event_log("User closed the 'Add new data' window.")
                break

            if eventAdd == "-ADD_NEW_DATA-":
                ch = sg.popup_yes_no('Do you want to add new data?', title='Confirmation')

                if ch == 'Yes':
                    # Get the selected golden file path to add new data
                    selected_json_file = get_file_path()

                    # Get the new data from the user, which will handle conversion for Binary
                    new_data = get_new_data(windowAddData)

                    # Use the original golden file name format for the edit.temp file
                    golden_file_name = os.path.basename(selected_json_file).replace(".json", "")
                    edit_temp_file_name = f"{golden_file_name}_edit.temp"
                    edit_temp_file_path = os.path.join("data", edit_temp_file_name)

                    # Check if the edit.temp file exists, if not, create it
                    if not os.path.exists(edit_temp_file_path):
                        with open(edit_temp_file_path, "w") as edit_temp_file:
                            edit_temp_file.write("[]")  # Initialize as an empty JSON array

                    # Write the new data to the edit.temp file (instead of directly modifying golden_file.json)
                    write_edit_temp_file(golden_file_name, new_data, "Add")  # Use golden_file_name as machine type

                    file_name = os.path.basename(selected_json_file)
                    sg.popup_ok("Successfully added the new data to the temp file: " + file_name, title="Information")

                    # Update the table by reloading from the temp files (edit.temp)
                    update_editor_gui(selected_json_file)

                    # Set table_changed to True since data has been added
                    table_changed = True  # Track that changes have been made

                    # Enable the Save button and set metadata to True when changes are made
                    window['-SAVE_GOLDEN_FILE-'].update(disabled=False)
                    window['-SAVE_GOLDEN_FILE-'].metadata = True

                    # Disable the Browse button as there are unsaved changes
                    window['-BROWSE2-'].update(disabled=True)
                    window['-RESTORE_REV-'].update(disabled=True)

                    # Automatically close the add window after successful addition
                    windowAddData.Close()
                    window_add_data_active = False

                else:
                    sg.popup_ok("No data is added.", title="Cancel")
                    write_into_event_log("User cancelled adding new data.")

            if eventAdd == "-CANCEL_ADD-":
                windowAddData.Close()
                window_add_data_active = False
                write_into_event_log("User cancelled adding new data and closed the window.")
                break

            elif eventAdd == "-SELECT_FORMAT_PAGE4-":
                selected_type = valuesAdd["-ADD_TYPE_DROPDOWN-"]

                if selected_type == "DWORD(32-bit)":
                    windowAddData.find_element("-ADD_FORMAT-").update(visible=True)
                    windowAddData.find_element("-ADD_FORMAT_DECIMAL-").update(visible=True)
                    windowAddData.find_element("-ADD_FORMAT_HEX-").update(visible=True)
                    windowAddData.find_element("-ADD_NEW_DATA-").update(disabled=False)
                    write_into_event_log("User selected DWORD(32-bit) and activated Decimal/Hexadecimal options.")

                elif selected_type == "QWORD(64-bit)":
                    windowAddData.find_element("-ADD_FORMAT-").update(visible=True)
                    windowAddData.find_element("-ADD_FORMAT_DECIMAL-").update(visible=True)
                    windowAddData.find_element("-ADD_FORMAT_HEX-").update(visible=True)
                    windowAddData.find_element("-ADD_NEW_DATA-").update(disabled=False)
                    write_into_event_log("User selected QWORD(64-bit) and activated Decimal/Hexadecimal options.")

                elif selected_type == "String":
                    windowAddData.find_element("-ADD_FORMAT-").update(visible=False)
                    windowAddData.find_element("-ADD_FORMAT_DECIMAL-").update(visible=False)
                    windowAddData.find_element("-ADD_FORMAT_HEX-").update(visible=False)
                    windowAddData.find_element("-ADD_NEW_DATA-").update(disabled=False)
                    write_into_event_log("User selected String type.")

                elif selected_type == "Binary":
                    windowAddData.find_element("-ADD_FORMAT-").update(visible=False)
                    windowAddData.find_element("-ADD_FORMAT_DECIMAL-").update(visible=False)
                    windowAddData.find_element("-ADD_FORMAT_HEX-").update(visible=False)
                    windowAddData.find_element("-ADD_NEW_DATA-").update(disabled=False)
                    write_into_event_log("User selected Binary type.")

                elif selected_type == "Multi-String":
                    windowAddData.find_element("-ADD_FORMAT-").update(visible=False)
                    windowAddData.find_element("-ADD_FORMAT_DECIMAL-").update(visible=False)
                    windowAddData.find_element("-ADD_FORMAT_HEX-").update(visible=False)
                    windowAddData.find_element("-ADD_NEW_DATA-").update(disabled=False)
                    write_into_event_log("User selected Multi-String type.")

                elif selected_type == "Expandable String":
                    windowAddData.find_element("-ADD_FORMAT-").update(visible=False)
                    windowAddData.find_element("-ADD_FORMAT_DECIMAL-").update(visible=False)
                    windowAddData.find_element("-ADD_FORMAT_HEX-").update(visible=False)
                    windowAddData.find_element("-ADD_NEW_DATA-").update(disabled=False)
                    write_into_event_log("User selected Expandable String type.")

            # When user confirms the selection with the add button inside the popup window
            if eventAdd == "-ADD_NEW_DATA-":
                registry_path = valuesAdd["-ADD_PATH-"]
                registry_name = valuesAdd["-ADD_NAME-"]
                registry_type = valuesAdd["-ADD_TYPE_DROPDOWN-"]
                registry_data = valuesAdd["-ADD_DATA-"]

                # Call add_registry_entry inside the block where registry data is defined
                success = add_registry_entry(registry_path, registry_name, registry_type, registry_data)

                if success:
                    write_into_event_log(
                        f"User added registry key: Path='{registry_path}', Name='{registry_name}', Type='{registry_type}', Data='{registry_data}' to the summary page, but not confirmed yet.")
                    window['-SAVE_GOLDEN_FILE-'].update(disabled=False)
                    window['-BROWSE2-'].update(disabled=True)

                else:
                    sg.popup_error(f"Failed to add registry key: {registry_name}.", title="Error")
                    write_into_error_log(f"Failed to add registry key: {registry_name} at {registry_path}.")

    if event == "-EDIT_GOLDEN_FILE-":
        update_editor_data = True
        selected_edit_file_path = 'data/editor/selected_data.json'
        selected_edit_data = load_registry_from_json2(selected_edit_file_path)

        data = [
            [row[0], row[1], row[2], row[3]]
            for row in selected_edit_data
        ]

        if len(selected_edit_data) == 1:
            window.find_element("-EDIT_GOLDEN_FILE-").update(disabled=False)
            for path, name, reg_data, reg_type in data:
                window_edit_data_active = True
                if reg_type == "REG_SZ":
                    reg_type = "String"
                elif reg_type == "REG_BINARY":
                    reg_type = "Binary"
                elif reg_type == "REG_DWORD_LITTLE_ENDIAN":
                    reg_type = "DWORD(32-bit)"
                elif reg_type == "REG_QWORD_LITTLE_ENDIAN":
                    reg_type = "QWORD(64-bit)"
                elif reg_type == "REG_MULTI_SZ":
                    reg_type = "Multi-String"
                elif reg_type == "REG_EXPAND_SZ":
                    reg_type = "Expandable String"

                write_into_event_log(
                    f"User is editing registry key: {name} at path: {path} with type: {reg_type} and data: {reg_data}."
                )
                window_edit_data = WindowEditData('Edit registry key data', path, name, reg_type, reg_data)

        else:
            window.find_element("-EDIT_GOLDEN_FILE-").update(disabled=True)

        while True:
            event_edit_data, values_edit_data = window_edit_data.read()

            if event_edit_data == sg.WIN_CLOSED:
                window_edit_data.Close()
                window_edit_data_active = False
                write_into_event_log("User closed the edit window without saving changes.")
                break

            if event_edit_data == "-PAGE4_EDIT_SAVE-":
                ch = sg.popup_yes_no("Do you want to save changes?")
                if ch == 'Yes':
                    select_file = "data/editor/selected_data.json"
                    selected_edit_data = load_registry_from_json2(select_file)

                    if selected_edit_data:
                        edited_data = edit_selected_data(selected_edit_data, window_edit_data)

                    edited_entries = []
                    file_path = get_file_path()

                    # Extract the golden file name without extension
                    golden_file_name = os.path.basename(file_path).replace(".json", "")

                    # Create the edit.temp file with the golden file name
                    edit_temp_file_name = f"{golden_file_name}_edit.temp"
                    edit_temp_file_path = os.path.join("data", edit_temp_file_name)

                    # Load existing edit.temp file data
                    current_edit_temp_data = load_registry_from_json2(edit_temp_file_path)

                    for path, name, reg_type, reg_data in edited_data:
                        # Load the original data from the golden file
                        original_data = load_registry_from_json2(file_path)

                        # Load the edit.temp file if it exists, otherwise initialize as an empty list
                        current_edit_temp_data = load_registry_from_json2(edit_temp_file_path)
                        if current_edit_temp_data is None:
                            current_edit_temp_data = []  # Initialize as empty if the file doesn't exist or is empty

                        for path, name, reg_type, reg_data in edited_data:
                            previous_type = None  # Initialize as None for newly added keys
                            previous_data = None  # Initialize as None for newly added keys

                            # Ensure we are comparing with the actual current values
                            for ori_entry in original_data:
                                if ori_entry[0] == path and ori_entry[1] == name:
                                    previous_type = ori_entry[3]
                                    previous_data = ori_entry[2]
                                    break

                            current_type = reg_type
                            current_data = reg_data

                            # If there are changes, check if it was added before
                            for temp_entry in current_edit_temp_data:
                                if temp_entry['Registry Key/Subkey Path'] == path and temp_entry[
                                    'Registry Name'] == name and temp_entry['Action'] == 'Add':
                                    temp_entry['Data'] = current_data
                                    temp_entry['Type'] = current_type
                                    break
                            else:
                                # If not previously added, treat it as a new "Edit"
                                edited_entries.append({
                                    'Registry Key/Subkey Path': path,
                                    'Registry Name': name,
                                    'Previous Data': previous_data if previous_data is not None else "N/A",
                                    'Previous Type': previous_type if previous_type is not None else "N/A",
                                    'Current Data': current_data,
                                    'Current Type': current_type,
                                    'Action': 'Edit'
                                })

                            write_into_event_log(
                                f"Registry key '{name}' at path '{path}' edited. "
                                f"Previous Type: '{previous_type if previous_type is not None else 'N/A'}', "
                                f"New Type: '{current_type}'. "
                                f"Previous Data: '{previous_data if previous_data is not None else 'N/A'}', "
                                f"New Data: '{current_data}'"
                            )

                    # Write the updated data to the edit.temp file
                    combined_entries = current_edit_temp_data + edited_entries if edited_entries else current_edit_temp_data
                    write_to_json(combined_entries, edit_temp_file_path)

                    sg.popup_ok("Changes recorded, please proceed to the Save button to save the changes.", title="Successfully")

                    # Refresh the table based on the temp files, not the golden file
                    update_editor_gui(file_path)

                    window.find_element("-EDIT_GOLDEN_FILE-").update(disabled=True)
                    window.find_element("-DELETE_FROM_GOLDEN_FILE-").update(disabled=True)

                    table_changed = True  # Track changes

                    # Enable Save button and disable Browse button
                    window['-SAVE_GOLDEN_FILE-'].update(disabled=False)
                    window['-BROWSE2-'].update(disabled=True)
                    window['-RESTORE_REV-'].update(disabled=True)

                    window_edit_data.Close()
                    window_edit_data_active = False
                    break

                else:
                    sg.popup_ok("No changes made.", title="Cancel")
                    write_into_event_log(f"User cancelled saving changes to registry key: {name} at {path}.")
                    window_edit_data.Close()
                    window_edit_data_active = False
                    break

            if event_edit_data == "-PAGE4_EDIT_CANCEL-":
                window_edit_data.Close()
                window_edit_data_active = False
                write_into_event_log(f"User cancelled editing registry key: {name} at {path}.")
                break

            if event_edit_data == "-PAGE4_SELECT_FORMAT-":
                if values_edit_data["-PAGE4_EDIT_TYPE_DROPDOWN-"] == "DWORD(32-bit)":
                    window_edit_data.find_element("-PAGE4_FORMAT-").update(visible=True)
                    window_edit_data.find_element("-PAGE4_FORMAT_DECIMAL-").update(visible=True)
                    window_edit_data.find_element("-PAGE4_FORMAT_HEX-").update(visible=True)
                    window_edit_data.find_element("-PAGE4_EDIT_SAVE-").update(disabled=False)
                    write_into_event_log(
                        f"User selected DWORD(32-bit) format for registry key: {name}.")

                elif values_edit_data["-PAGE4_EDIT_TYPE_DROPDOWN-"] == "QWORD(64-bit)":
                    window_edit_data.find_element("-PAGE4_FORMAT-").update(visible=True)
                    window_edit_data.find_element("-PAGE4_FORMAT_DECIMAL-").update(visible=True)
                    window_edit_data.find_element("-PAGE4_FORMAT_HEX-").update(visible=True)
                    window_edit_data.find_element("-PAGE4_EDIT_SAVE-").update(disabled=False)
                    write_into_event_log(
                        f"User selected QWORD(64-bit) format for registry key: {name}.")

                elif values_edit_data["-PAGE4_EDIT_TYPE_DROPDOWN-"] == "String":
                    window_edit_data.find_element("-PAGE4_FORMAT-").update(visible=False)
                    window_edit_data.find_element("-PAGE4_FORMAT_DECIMAL-").update(visible=False)
                    window_edit_data.find_element("-PAGE4_FORMAT_HEX-").update(visible=False)
                    window_edit_data.find_element("-PAGE4_EDIT_SAVE-").update(disabled=False)
                    write_into_event_log(f"User selected String format for registry key: {name}.")

                elif values_edit_data["-PAGE4_EDIT_TYPE_DROPDOWN-"] == "Binary":
                    window_edit_data.find_element("-PAGE4_FORMAT-").update(visible=False)
                    window_edit_data.find_element("-PAGE4_FORMAT_DECIMAL-").update(visible=False)
                    window_edit_data.find_element("-PAGE4_FORMAT_HEX-").update(visible=False)
                    window_edit_data.find_element("-PAGE4_EDIT_SAVE-").update(disabled=False)
                    write_into_event_log(f"User selected Binary format for registry key: {name}.")

                elif values_edit_data["-PAGE4_EDIT_TYPE_DROPDOWN-"] == "Multi-String":
                    window_edit_data.find_element("-PAGE4_FORMAT-").update(visible=False)
                    window_edit_data.find_element("-PAGE4_FORMAT_DECIMAL-").update(visible=False)
                    window_edit_data.find_element("-PAGE4_FORMAT_HEX-").update(visible=False)
                    window_edit_data.find_element("-PAGE4_EDIT_SAVE-").update(disabled=False)
                    write_into_event_log(
                        f"User selected Multi-String format for registry key: {name}.")

                elif values_edit_data["-PAGE4_EDIT_TYPE_DROPDOWN-"] == "Expandable String":
                    window_edit_data.find_element("-PAGE4_FORMAT-").update(visible=False)
                    window_edit_data.find_element("-PAGE4_FORMAT_DECIMAL-").update(visible=False)
                    window_edit_data.find_element("-PAGE4_FORMAT_HEX-").update(visible=False)
                    window_edit_data.find_element("-PAGE4_EDIT_SAVE-").update(disabled=False)
                    write_into_event_log(
                        f"User selected Expandable String format for registry key: {name}.")


    if event == "-DELETE_FROM_GOLDEN_FILE-":
        ch = sg.popup_yes_no('Do you want to DELETE the selected data?', title='Delete Registry Key Confirmation')
        if ch == 'Yes':
            # File path for selected data
            select_file = "data/editor/selected_data.json"

            # Backup the selected data (deleted items) before deletion
            exported_file = 'data/editor/selected_data.json'
            backup_folder = "Backup/Page 4/Backup(Deleted_Page4)"
            backup_deleted_registry_page4(exported_file, backup_folder)

            # Load the selected data to delete
            selected_golden_data = load_registry_from_json2(select_file)

            if selected_golden_data:
                # Log deleted registry details
                for entry in selected_golden_data:
                    registry_key = entry[0]  # Registry Path
                    registry_name = entry[1]  # Registry Name
                    registry_type = entry[3]  # Registry Type
                    registry_data = entry[2]  # Registry Data
                    log_deleted_registry_key(registry_key, registry_name, registry_type, registry_data)

                # Write the deleted entries to the edit temp file
                machine_type = os.path.splitext(os.path.basename(file_path))[0]
                write_edit_temp_file(machine_type, selected_golden_data, "Delete")

                # Update the GUI table with the latest changes
                update_editor_gui(file_path)

                sg.popup_auto_close('Successfully deleted selected keys.', title="Successful")

                # Track table changes
                table_changed = True

                # Enable the Save button and set metadata to True
                window['-SAVE_GOLDEN_FILE-'].update(disabled=False)
                window['-SAVE_GOLDEN_FILE-'].metadata = True

                # Disable the Browse button to indicate unsaved changes
                window['-BROWSE2-'].update(disabled=True)

                # Disable the delete button & edit button
                window['-DELETE_FROM_GOLDEN_FILE-'].update(disabled=True)
                window['-EDIT_GOLDEN_FILE-'].update(disabled=True)

                # Optional: Disable other action buttons if required
                window['-RESTORE_REV-'].update(disabled=True)

            else:
                sg.popup_error("Failed to delete registry keys.", title="Error")

    # Re-enable the delete button when a checkbox is ticked
    if event == "-TABLE_EDITOR-":
        # Check if any checkbox is ticked
        selected_indices = values["-TABLE_EDITOR-"]
        if selected_indices:
            # Enable the delete button when an item is selected
            window['-DELETE_FROM_GOLDEN_FILE-'].update(disabled=False)

        # Update flags if other pages/tables need to be refreshed
        update_editor_data = True
        update_page3_data = True
        update_page4_upper_data = True
        update_page4_bottom_data = True

    if event == "-SAVE_GOLDEN_FILE-":
        # Call makeWinSave when the Save button is clicked
        save_status = makeWinSave("Save Registry Changes")  # Opens save summary page and handles save/discard actions

        # Check the result of makeWinSave to determine if changes were saved or discarded
        if save_status == "saved":

            # Reset indicators for unsaved changes after saving
            table_changed = False
            window['-SAVE_GOLDEN_FILE-'].metadata = False
            window['-SAVE_GOLDEN_FILE-'].update(disabled=True)  # Disable Save button
            window['-BROWSE2-'].update(disabled=False)  # Re-enable Browse button

        elif save_status == "discarded":
            sg.popup("Changes discarded successfully.", title="Discard Confirmation")

            # Reset indicators for unsaved changes after discarding
            table_changed = False
            window['-SAVE_GOLDEN_FILE-'].metadata = False
            window['-SAVE_GOLDEN_FILE-'].update(disabled=True)  # Disable Save button
            window['-BROWSE2-'].update(disabled=False)  # Re-enable Browse button

        elif save_status == "closed_without_action":
            # No changes were saved or discarded; do nothing with the flags
            sg.popup("Save summary page closed without saving or discarding.", title="No Action Taken")

    # Event handler for RESTORE_REV
    if event == "-RESTORE_REV-":
        # Check if a machine type has been selected first
        if not file_path or file_path == "" or file_path == "Browse":
            sg.popup_ok("Please select a machine type and browse before attempting to restore a backup.", title="Select Machine Type")
        else:
            # Define the default backup folder path where the golden_file_rev files are stored
            backup_folder = os.path.join("Backup", "Backup(Golden File)")

            # Step 1: Open the file explorer directly to select the backup file
            restore_file_path = sg.popup_get_file(
                "Select a backup file to restore",
                default_path=backup_folder,
                file_types=(("JSON Files", "*.json"),),
                no_window=True  # No additional window prompt
            )

            # Check if the user selected a file
            if restore_file_path and os.path.exists(restore_file_path):
                try:
                    # Get the current golden file path
                    file_name = os.path.basename(file_path).replace(".json",
                                                                    "")  # Remove any trailing .json if it exists
                    golden_file_path = os.path.join("Golden File", f"{file_name}.json")
                    action_temp_path = os.path.join("data", f"{file_name}_action.temp")  # Create a temp action file

                    if os.path.exists(golden_file_path):
                        # Step 2: Compare the current golden file with the selected rev file and create action.temp
                        added_deleted_entries, edited_entries = compare_golden_and_rev_json(golden_file_path, restore_file_path, action_temp_path)

                        # Step 3: Create a window (similar to makeWinSave) to display the comparison
                        makeWin4Restore(added_deleted_entries, edited_entries)
                    else:
                        sg.popup_error(f"Golden file '{file_name}.json' not found at: {golden_file_path}",
                                       title="File Not Found")
                except Exception as e:
                    sg.popup_error(f"An error occurred while comparing files: {str(e)}", title="Error")
            else:
                sg.popup_ok("No file selected.", title="Information")

    # A dictionary to map machine types to their respective .temp files
    machine_temp_files = {
        "MicroAOI": "MicroAOI.temp",
        "Semicon": "Semicon.temp",
        "SideCam": "SideCam.temp",
        "SMT": "SMT.temp",
        "custom_selection": "custom_selection.temp"
    }


    '''
    if event == '-RESTORE_PAGE4-':
        # Load the deleted data from the file
        deleted_data_file = "data/deleted_registry_data_page4.json"
        deleted_data = load_registry_from_json2(deleted_data_file)

        if deleted_data:
            # Ensure that there are items to restore
            if not deleted_data:
                sg.popup_error("No items available for restoration.")
            else:
                # Open the restore confirmation window
                window_restore_page4 = makeWinRestorePage4('List of Deleted Registry keys from Page 4')
                window_restore_page4.finalize()

                # Populate the table in the window with the deleted data
                restored_table_data = [
                    [BLANK_BOX, row[0], row[1], row[2], row[3]]  # Display order: Checkbox, Key, Name, Type, Data
                    for row in deleted_data
                ]
                window_restore_page4['-TABLE_RESTORE_PAGE4-'].update(values=restored_table_data)

                # Wait for user actions within the restore window
                while True:
                    event_restore, values_restore = window_restore_page4.read()

                    # Break if window is closed
                    if event_restore == sg.WIN_CLOSED:
                        break

                    if isinstance(event_restore, tuple) and event_restore[0] == '-TABLE_RESTORE_PAGE4-':
                        row_index = event_restore[2][0]
                        column_index = event_restore[2][1]

                        # Toggle the checkbox for selection
                        if row_index is not None and column_index == 0:
                            table_values = window_restore_page4['-TABLE_RESTORE_PAGE4-'].get()

                            # Toggle checkbox
                            if table_values[row_index][0] == BLANK_BOX:
                                table_values[row_index][0] = CHECKED_BOX
                            else:
                                table_values[row_index][0] = BLANK_BOX

                            window_restore_page4['-TABLE_RESTORE_PAGE4-'].update(values=table_values)

                            # Enable/Disable restore button based on selection
                            selected_rows = [i for i, row in enumerate(table_values) if row[0] == CHECKED_BOX]
                            window_restore_page4['-restoreSelectedPage4-'].update(disabled=not selected_rows)

                    # Inside the '-restoreSelectedPage4-' event
                    if event_restore == '-restoreSelectedPage4-':
                        ch_confirm = sg.popup_yes_no('Do you want to restore the selected items?',
                                                     title='Restore Confirmation')

                        if ch_confirm == 'Yes':
                            try:
                                # Get the full table values from the window
                                table_values = window_restore_page4['-TABLE_RESTORE_PAGE4-'].get()

                                # Iterate through the selected rows for restoration
                                restored_data = []
                                for row_index in selected_rows:
                                    row = table_values[row_index]

                                    # Fetching registry data in correct order: Key, Name, Type, Data
                                    registry_key = row[1]
                                    registry_name = row[2]
                                    registry_type = row[3]
                                    registry_data = row[4]

                                    # Debug log for key, type, and data
                                    print(
                                        f"Restoring key: {registry_key}, Type: {registry_type}, Data: {registry_data}")

                                    # Validate the registry type before restoring
                                    valid_registry_types = ['REG_SZ', 'REG_BINARY', 'REG_DWORD', 'REG_QWORD',
                                                            'REG_MULTI_SZ', 'REG_EXPAND_SZ', 'REG_DWORD_LITTLE_ENDIAN',
                                                            'REG_QWORD_LITTLE_ENDIAN']
                                    if registry_type not in valid_registry_types or registry_type == '':
                                        sg.popup_error(
                                            f"Error restoring key '{registry_key}': Unsupported or missing registry type: {registry_type}")
                                        continue

                                    if registry_type in ['REG_DWORD', 'REG_DWORD_LITTLE_ENDIAN']:
                                        try:
                                            # Convert hex to int for DWORD
                                            registry_data_int = int(registry_data,
                                                                    16)  # Ensure we define registry_data_int here
                                            # Ensure it's displayed with 8 digits (32 bits for DWORD)
                                            registry_data_hex = f"0x{registry_data_int:08X}"
                                            print(f"Restoring {registry_type} as hex: {registry_data_hex}")
                                            registry_data = registry_data_hex  # Use the hex value for restoration and display
                                        except ValueError as e:
                                            print(f"Error converting {registry_type} data: {e}")
                                            sg.popup_error(f"Error converting {registry_type} data: {e}")
                                            continue

                                    elif registry_type in ['REG_QWORD', 'REG_QWORD_LITTLE_ENDIAN']:
                                        try:
                                            # Convert hex to int for QWORD
                                            registry_data_int = int(registry_data,
                                                                    16)  # Ensure we define registry_data_int here
                                            # Ensure it's displayed with 16 digits (64 bits for QWORD)
                                            registry_data_hex = f"0x{registry_data_int:016X}"
                                            print(f"Restoring {registry_type} as hex: {registry_data_hex}")
                                            registry_data = registry_data_hex  # Use the hex value for restoration and display
                                        except ValueError as e:
                                            print(f"Error converting {registry_type} data: {e}")
                                            sg.popup_error(f"Error converting {registry_type} data: {e}")
                                            continue

                                    try:
                                        # Perform restoration of registry key
                                        restore_registry_key(registry_key, registry_name, registry_type, registry_data)
                                        # Collect restored data for writing back to JSON and .txt file
                                        restored_data.append(
                                            [registry_key, registry_name, registry_data, registry_type])
                                    except Exception as e:
                                        sg.popup_error(f"Error restoring key '{registry_key}': {str(e)}")
                                        continue

                                # Write restored data back to the Golden File (sample-MicroAOI.json)
                                machine_type = values[
                                    'Dropdown2'].lower()  # Get the machine type (converted to lowercase)
                                golden_file_map = {
                                    "microaoi": "sample-MicroAOI.json",
                                    "semicon": "sample-semicon.json",
                                    "sidecam": "sample-SideCam.json",
                                    "smt": "sample-SMT.json"
                                }

                                golden_file = f"C:/Users/zi-kai.soon/Documents/zk/Source code/Golden File/{golden_file_map[machine_type]}"
                                original_data = get_json_data(golden_file)

                                # Append restored data to original file
                                updated_data = original_data + restored_data

                                write_restored_data_to_json(updated_data, golden_file)

                                # Backup the restored data as a .txt file based on machine type
                                timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
                                backup_filename = f"golden_list_{machine_type}_{timestamp}.txt"
                                backup_folder = "Backup/Backup(Restored_Page4)"
                                os.makedirs(backup_folder, exist_ok=True)  # Ensure backup folder exists
                                backup_file_path = os.path.join(backup_folder, backup_filename)

                                # Write restored data to the .txt file in JSON-like format
                                with open(backup_file_path, 'w') as backup_file:
                                    backup_file.write("[\n")
                                    for index, item in enumerate(restored_data):
                                        formatted_entry = f'    [\n        "{item[0]}",\n        "{item[1]}",\n        "{item[2]}",\n        "{item[3]}"\n    ]'
                                        if index < len(restored_data) - 1:
                                            formatted_entry += ","  # Add a comma after every entry except the last one
                                        backup_file.write(formatted_entry + "\n")
                                    backup_file.write("]\n")

                                # Overwrite the selected_restored_data_page4.json file with new restored data
                                restored_data_file = "data/selected_restored_data_page4.json"
                                write_restored_data_to_json(restored_data, restored_data_file)

                                # Refresh the editor table with the updated data
                                update_editor_gui(file_path=golden_file)

                                # Remove the restored items from the deleted data list
                                for row_index in sorted(selected_rows, reverse=True):
                                    del restored_table_data[row_index]
                                window_restore_page4['-TABLE_RESTORE_PAGE4-'].update(values=restored_table_data)

                                # Update the JSON file for deleted items after restoration
                                remaining_deleted_data = [row for i, row in enumerate(deleted_data) if
                                                          i not in selected_rows]
                                write_restored_data_to_json(remaining_deleted_data, deleted_data_file)

                                # If no items left in restore table, disable restore button and close window
                                if not restored_table_data:
                                    window_restore_page4['-restoreSelectedPage4-'].update(disabled=True)
                                    window['-RESTORE_PAGE4-'].update(disabled=True)

                                sg.popup_auto_close(
                                    f'Successfully restored selected keys and backed up to {backup_filename}.')
                                window_restore_page4.close()

                            except Exception as e:
                                sg.popup_error(f"An error occurred during restore: {e}")

                        else:
                            sg.popup_ok("Restoration cancelled.")
'''
    if event in ("-SEARCH_BUTTON-", "\r", "-SEARCH-"):
        # For Page 1
        search_text_page1 = values["-SEARCH-"].strip().lower()
        table_data = window["-TABLE-"].get()
        perform_search(search_text_page1, table_data)

        if not search_text_page1:
            window["-TABLE-"].update(values=original_table_data)

    # page 2
    if event in ("-SEARCH_BUTTON_SIZE-", "\r", "-SEARCH_SIZE-"):
        # For Page 2
        search_text_page2 = values["-SEARCH_SIZE-"].strip().lower()
        table_data = window["-SIZE_TABLE-"].get()
        perform_size_search(search_text_page2, table_data)

        if not search_text_page2:
            window["-SIZE_TABLE-"].update(values=original_size_table_data)

    # page3
    if event in ("-SEARCH_BUTTON_REG-", "\r", "-SEARCH_REG-"):
        # For Page 3
        search_text_page3 = values["-SEARCH_REG-"].strip().lower()
        if search_text_page3 != '':
            table_data = window["-TABLE_REG-"].get()
            perform_reg_search(search_text_page3, table_data)
            # sort_reg_table(window, event, col_count=6, current_sort_order=current_sort_order_table)

        if not search_text_page3:
            update_reg_gui()

    # Page 4 Search event handler
    if event in ("\r", "-SEARCH_EDITOR-"):
        search_text_page4 = values["-SEARCH_EDITOR-"].strip().lower()  # Ensure the correct search variable is used
        table_data = window["-TABLE_EDITOR-"].get()
        perform_editor_search(search_text_page4, table_data)

        # If search is empty, reload the table
        if not search_text_page4:  # Ensure the correct variable is used here
            file_path = get_file_path()
            update_editor_gui(file_path)

    # Update data on Page 1 when the check button is clicked
    if update_page1_data:
        if not values["-SEARCH-"]:
            table_data = window["-TABLE-"].get()
            window["-TABLE-"].update(values=table_data)
            ###
            # sort_table_without_click(window, event, table_data,0, current_sort_order=current_sort_order_table)
            # Example usage for a 4-row, 5-column table
            sort_table(window, event, row_count=5, col_count=5, current_sort_order=current_sort_order_table)
            # Example usage for a 4-row, 5-column table
            sort_size_table(window, event, row_count=4, col_count=4, current_sort_order=current_sort_order_table)

    # refresh_window()

    # Update data on Page 2 when the browse button is clicked
    if update_page2_data:
        if not values["-SEARCH_SIZE-"]:
            table_data = window["-SIZE_TABLE-"].get()
            window["-SIZE_TABLE-"].update(values=table_data)
            ###
            ##sort_size_table_without_click(window, event, table_data, 0, current_sort_order=current_sort_order_table)
            # Example usage for a 4-row, 5-column table
            sort_table(window, event, row_count=5, col_count=5, current_sort_order=current_sort_order_table)
            # Example usage for a 4-row, 5-column table
            sort_size_table(window, event, row_count=5, col_count=5, current_sort_order=current_sort_order_table)

    # refresh_window()

    if update_page3_data:
        if not values["-SEARCH_REG-"]:
            table_data = window["-TABLE_REG-"].get()
            row_colors = set_status_color(table_data)
            window["-TABLE_REG-"].update(values=table_data, row_colors=row_colors)

            # get the selected row data
            # handle_table_selection(window, event, values)
            # sort_reg_table(window, event, col_count=6, current_sort_order=current_sort_order_table)
        sort_reg_table(window, event, col_count=6, current_sort_order=current_sort_order_table)
        # refresh_window()

    if update_editor_data:

        sort_editor_table(window, event, col_count=5, current_sort_order=current_sort_order_table)
        # handle the table selection and save the selected keys into the json file
        get_selected_delete_keys_result = handle_table_selection4(window, event, values)

        if get_selected_delete_keys_result:
            output_file = "data/editor/selected_data.json"
            os.makedirs(os.path.dirname(output_file), exist_ok=True)
            write_new_data_to_json(get_selected_delete_keys_result, output_file)

window.close()


