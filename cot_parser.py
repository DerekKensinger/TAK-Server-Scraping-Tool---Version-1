import os
import re
import traceback
import xml.etree.ElementTree as ET
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
import customtkinter as ctk
from tkinter import filedialog, messagebox

#############################
# Global State Class
#############################
class GlobalState:
    selectedFile = None
    selectedFiles = []
    selectedFolder = None
    operationHistory = []
    exportCounter = 0
    currentOperation = ""
    startTime = None
    endTime = None
    processedFilePath = None

#############################
# Utility Functions
#############################
def validateFilePath(filePath):
    return filePath.strip('"')

def createParsedLogsFolder(filePath):
    directory = os.path.dirname(filePath)
    parsedLogsPath = os.path.join(directory, "ParsedLogs")

    if not os.path.exists(parsedLogsPath):
        os.makedirs(parsedLogsPath)
    return parsedLogsPath

#############################
# File Processing Functions
#############################
def enforce_indentation(event):
    # This function is no longer necessary if we are no longer cleaning the file,
    # but we leave it here in case it's still called elsewhere.
    lines = event.split(b'\n')
    indented_lines = []
    indent_level = 0
    for line in lines:
        stripped_line = line.strip()
        if stripped_line.startswith(b'</'):
            indent_level -= 1
        indented_lines.append(b'    ' * indent_level + stripped_line)
        if stripped_line.startswith(b'<') and not stripped_line.startswith(b'</') and not stripped_line.endswith(b'/>'):
            indent_level += 1
    return b'\n'.join(indented_lines)

def cleanFileContent(content):
    # This cleaning step is no longer required before export.
    # If needed, you can remove or leave this function as is.
    cleaned_content = []
    events = content.split(b'</event>')
    for i, event in enumerate(events[:-1]):
        cleaned_event = event.strip() + b'</event>'
        cleaned_event = enforce_indentation(cleaned_event)
        cleaned_content.append(cleaned_event)
    return b'\n'.join(cleaned_content)

def loadFiles(filePath):
    GlobalState.selectedFiles = []
    if os.path.isfile(filePath):
        GlobalState.selectedFiles.append(filePath)
        GlobalState.operationHistory.append(f"Loaded file: {os.path.basename(filePath)}")
    elif os.path.isdir(filePath):
        for root, dirs, files in os.walk(filePath):
            for file in files:
                fullPath = os.path.join(root, file)
                GlobalState.selectedFiles.append(fullPath)
        GlobalState.operationHistory.append(f"Loaded folder: {filePath}")
    else:
        raise ValueError("The provided path is neither a file nor a folder. Please check the path and try again.")

def removeDuplicates(filePath, log_callback):
    if not os.path.isfile(filePath):
        log_callback(f"{filePath} is not a valid file.")
        return
    answer = messagebox.askyesno("Remove Duplicates", f"Remove duplicates from {filePath}? This action cannot be undone.")
    if not answer:
        log_callback("Operation canceled by the user.")
        return
    log_callback(f"Removing duplicates from: {filePath}")
    with open(filePath, 'rb') as f:
        content = f.read()
    content = content.replace(b'<?xml version="1.0" encoding="UTF-8"?>', b'')
    events = content.split(b'</event>')
    seen = set()
    unique_events = []
    total = len(events) - 1
    for idx, event in enumerate(events[:-1]):
        event_str = event + b'</event>'
        try:
            root = ET.fromstring(event_str)
            uid = root.get('uid')
            time = root.get('time')
            if time:
                time.split('.')[0]
            key = (uid, time)
            if key not in seen:
                seen.add(key)
                unique_events.append(event_str)
        except ET.ParseError:
            pass
        if idx % 1000 == 0:
            log_callback(f"Processed {idx}/{total} events...")
    newFileName = "NoDuplicates_" + os.path.basename(filePath)
    newFilePath = os.path.join(os.path.dirname(filePath), newFileName)
    with open(newFilePath, 'wb') as cleanedFile:
        cleanedFile.write(b'<?xml version="1.0" encoding="UTF-8"?>\n')
        for event in unique_events:
            cleanedFile.write(event)
    log_callback(f"Duplicates removed. Cleaned file saved as: {newFilePath}")
    GlobalState.selectedFiles = [newFilePath]

def cleanTimeString(time_str):
    cleaned_time_str = re.sub(r'[^0-9T:\-Z]', '', time_str)
    date_part = cleaned_time_str[:10]
    time_part = cleaned_time_str[11:].replace("Z", "")
    time_parts = time_part.split(":")
    if len(time_parts) == 3:
        hours, minutes, seconds = time_parts
    else:
        hours = time_parts[0] if len(time_parts) > 0 else "00"
        minutes = time_parts[1] if len(time_parts) > 1 else "00"
        seconds = time_parts[2] if len(time_parts) > 2 else "00"
    hours = hours.zfill(2)[:2]
    minutes = minutes.zfill(2)[:2]
    seconds = seconds.zfill(2)[:2]
    return f"{date_part}T{hours}:{minutes}:{seconds}Z"

def adjustEventTimes(filePath, new_time_str, log_callback):
    if not os.path.isfile(filePath):
        log_callback(f"{filePath} is not a valid file.")
        return
    try:
        if new_time_str.endswith('Z'):
            user_time = datetime.strptime(new_time_str, "%Y-%m-%dT%H:%M:%SZ")
        else:
            user_time = datetime.strptime(new_time_str, "%Y-%m-%dT%H:%M:%S")
    except ValueError:
        log_callback(f"Invalid time format: {new_time_str}")
        return
    log_callback(f"Adjusting event times in: {filePath}")
    with open(filePath, 'rb') as f:
        content = f.read()
    events = content.split(b'</event>')
    if not events or len(events) == 1:
        log_callback(f"No events found in the file: {filePath}")
        return
    first_event_str = events[0] + b'</event>'
    first_event_root = ET.fromstring(first_event_str)
    first_event_time_str = first_event_root.get('time').split('.')[0]
    first_event_time_str = cleanTimeString(first_event_time_str)
    try:
        first_event_time = datetime.strptime(first_event_time_str, "%Y-%m-%dT%H:%M:%SZ")
    except ValueError:
        first_event_time = datetime.strptime(first_event_time_str, "%Y-%m-%dT%H:%M:%S")
    time_offset = user_time - first_event_time
    unique_events = []
    total = len(events) - 1
    for idx, event in enumerate(events[:-1]):
        event_str = event + b'</event>'
        try:
            root = ET.fromstring(event_str)
            for time_attr in ['time', 'start', 'stale']:
                original = root.get(time_attr)
                if not original:
                    continue
                event_time_str = original.split('.')[0]
                event_time_str = cleanTimeString(event_time_str)
                try:
                    event_time = datetime.strptime(event_time_str, "%Y-%m-%dT%H:%M:%SZ")
                except ValueError:
                    event_time = datetime.strptime(event_time_str, "%Y-%m-%dT%H:%M:%S")
                new_event_time = event_time + time_offset
                root.set(time_attr, new_event_time.strftime("%Y-%m-%dT%H:%M:%SZ"))
            event_string = ET.tostring(root, encoding='utf-8')
            event_string = event_string.replace(b' />', b'/>')
            unique_events.append(event_string)
        except ET.ParseError:
            pass
        if idx % 1000 == 0:
            log_callback(f"Adjusted {idx}/{total} events...")
    newFileName = "TimeAdjusted_" + os.path.basename(filePath)
    newFilePath = os.path.join(os.path.dirname(filePath), newFileName)
    with open(newFilePath, 'wb') as adjustedFile:
        adjustedFile.write(b'<?xml version="1.0" encoding="UTF-8"?>\n')
        for event in unique_events:
            adjustedFile.write(event)
    log_callback(f"Event times adjusted. File saved as: {newFilePath}")
    GlobalState.selectedFiles = [newFilePath]

def formatEvent(event_str):
    try:
        root = ET.fromstring(event_str)
        event_string = ET.tostring(root, encoding='utf-8', method='xml').decode('utf-8')
        event_string = event_string.replace(' />', '/>')
        return event_string + "\n"
    except ET.ParseError:
        return event_str.decode('utf-8') + "\n"

def writeLogFile(log_content, original_file_path, log_number):
    directory = os.path.dirname(original_file_path)
    base_name = os.path.basename(original_file_path)
    log_file_name = f"{os.path.splitext(base_name)[0]}_Log{log_number}.txt"
    log_file_path = os.path.join(directory, log_file_name)
    with open(log_file_path, 'w', encoding='utf-8') as log_file:
        log_file.write(''.join(log_content))
    return log_file_name

def splitAndExportFile(filePath, max_file_size_mb, log_callback):
    if not os.path.isfile(filePath):
        log_callback(f"{filePath} is not a valid file.")
        return
    try:
        max_file_size_bytes = int(max_file_size_mb) * 1024 * 1024
        if max_file_size_bytes > 100 * 1024 * 1024:
            log_callback("Maximum file size exceeded.")
            return
    except ValueError:
        log_callback("Invalid input. Please enter a valid number.")
        return
    log_callback(f"Splitting and exporting the file: {filePath}")
    with open(filePath, 'rb') as f:
        content = f.read()
    events = content.split(b'</event>')
    if not events or len(events) == 1:
        log_callback(f"No events found in the file: {filePath}")
        return
    current_file_size = 0
    log_number = 1
    current_log_content = ['<?xml version="1.0" encoding="UTF-8"?>\n']
    total_events = len(events) - 1
    for i, event in enumerate(events[:-1]):
        event_str = event.strip() + b'</event>\n'
        formatted_event = formatEvent(event_str)
        event_size = len(formatted_event)
        if current_file_size + event_size > max_file_size_bytes:
            created = writeLogFile(current_log_content, filePath, log_number)
            log_callback(f"Created: {created}")
            log_number += 1
            current_file_size = 0
            current_log_content = ['<?xml version="1.0" encoding="UTF-8"?>\n']
        current_log_content.append(formatted_event)
        current_file_size += event_size
        if i % 1000 == 0:
            log_callback(f"Processed {i}/{total_events} events...")
    if current_log_content:
        created = writeLogFile(current_log_content, filePath, log_number)
        log_callback(f"Created: {created}")
    log_callback(f"File splitting and export complete. {log_number} files created.")

def extractUIDsAndCallsigns(filePath, log_callback):
    if not os.path.isfile(filePath):
        log_callback(f"{filePath} is not a valid file.")
        return
    log_callback(f"Extracting UIDs and Callsigns from: {filePath}")
    unique_values = set()
    error_log = []
    try:
        with open(filePath, 'rb') as f:
            content = f.read()
        events = content.split(b'</event>')
        total = len(events) - 1
        for idx, event in enumerate(events[:-1]):
            event_str = event + b'</event>'
            try:
                root = ET.fromstring(event_str)
                uid = root.get('uid')
                if uid is None:
                    error_log.append("Missing UID in event")
                    continue
                if "ANDROID-" in uid:
                    callsign_element = root.find(".//contact")
                    if callsign_element is not None and 'callsign' in callsign_element.attrib:
                        callsign = callsign_element.attrib['callsign']
                        unique_values.add(callsign)
                    else:
                        error_log.append(f"Missing or malformed callsign for UID: {uid}")
                else:
                    unique_values.add(uid)
            except ET.ParseError:
                error_log.append("Malformed XML in event")
            if idx % 1000 == 0:
                log_callback(f"Processed {idx}/{total} events...")
    except Exception as e:
        log_callback(f"Error reading file: {e}")
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "UIDs and Callsigns"
    for i, value in enumerate(sorted(unique_values), start=1):
        ws[f'A{i}'] = value
    excel_filename = os.path.join(os.path.dirname(filePath), "UIDs_Callsigns.xlsx")
    wb.save(excel_filename)
    log_callback(f"Extraction complete. UIDs and Callsigns saved to: {excel_filename}")
    if error_log:
        error_log_filename = os.path.join(os.path.dirname(filePath), "Error_Log.txt")
        with open(error_log_filename, 'w') as error_file:
            error_file.write("\n".join(error_log))
        log_callback(f"Errors encountered. See {error_log_filename} for details.")
    GlobalState.selectedFiles = [filePath]

#############################
# Updated Export CoT Details Code for Uncleaned Logs
#############################
def export_cot_details(filePath, log_callback):
    if not os.path.isfile(filePath):
        log_callback(f"{filePath} is not a valid file.")
        return

    log_callback("Parsing file for CoT details export with robust fallback extraction...")

    # Read and decode content, replacing invalid chars
    with open(filePath, 'rb') as f:
        raw_content = f.read()

    content_str = raw_content.decode('utf-8', 'replace')
    # Remove invalid XML chars
    content_str = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', content_str)

    events = content_str.split('</event>')
    total = len(events) - 1
    uid_data = {}

    def parse_remarks_line_by_line(remarks_text, data_dict):
        for line in remarks_text.splitlines():
            line = line.strip()
            if ":" in line:
                key, value = line.split(":", 1)
                clean_key = key.strip().replace(" ", "_")
                data_dict[f"remarks_{clean_key}"] = value.strip()

    def fallback_extract_data(event_text):
        fallback_data = {}
        # UID
        uid_match = re.search(r'uid="([^"]+)"', event_text)
        uid_val = uid_match.group(1) if uid_match else "UNKNOWN_UID"
        fallback_data["uid"] = uid_val

        # Extract detail block
        detail_match = re.search(r'<detail>([\s\S]*?)</detail>', event_text)
        if detail_match:
            detail_content = detail_match.group(1)

            # Extract remarks text if present
            remarks_match = re.search(r'<remarks[^>]*>([\s\S]*?)</remarks>', detail_content)
            if remarks_match:
                remarks = remarks_match.group(1).strip()
                # Parse line-by-line for key-value pairs
                parse_remarks_line_by_line(remarks, fallback_data)
                # Also store the raw remarks if desired
                fallback_data['detail_detail_remarks_text'] = remarks

            # Extract attributes from all tags within detail
            # This handles status, __group, takv, track, contact, etc.
            # We'll find all tags inside <detail> except the closing tag.
            tags_in_detail = re.findall(r'<(\w[\w\-_]*)([^>]*)>', detail_content)

            for tag_name, attrs_str in tags_in_detail:
                # Skip the 'detail' closing scenario
                if tag_name.lower() == 'detail':
                    continue

                # If this is remarks, we've already handled it line-by-line above
                if tag_name.lower() == 'remarks':
                    continue

                # Extract attributes for this tag
                attrs = re.findall(r'(\w[\w\-_]*)="([^"]*)"', attrs_str)
                for attr_name, attr_val in attrs:
                    # Store as detail_<tag>_<attribute> = value
                    fallback_data[f"detail_{tag_name}_{attr_name}"] = attr_val

        else:
            # If no <detail> tag, just do what we did previously for known tags:
            contact_match = re.search(r'<contact[^>]*callsign="([^"]+)"', event_text)
            if contact_match:
                fallback_data['contact_callsign'] = contact_match.group(1)

            # Remarks outside of detail (rare, but just in case)
            remarks_match = re.search(r'<remarks[^>]*>([\s\S]*?)</remarks>', event_text)
            if remarks_match:
                remarks = remarks_match.group(1).strip()
                parse_remarks_line_by_line(remarks, fallback_data)
                fallback_data['detail_detail_remarks_text'] = remarks

            # Point, track (if outside detail, which is uncommon)
            point_match = re.search(r'<point[^>]+>', event_text)
            if point_match:
                point_tag = point_match.group(0)
                for attr in ["lat", "lon", "hae", "ce", "le"]:
                    attr_match = re.search(fr'{attr}="([^"]+)"', point_tag)
                    if attr_match:
                        fallback_data[f"point_{attr}"] = attr_match.group(1)

            track_match = re.search(r'<track[^>]+>', event_text)
            if track_match:
                track_tag = track_match.group(0)
                for attr in ["speed", "course"]:
                    attr_match = re.search(fr'{attr}="([^"]+)"', track_tag)
                    if attr_match:
                        fallback_data[f"detail_track_{attr}"] = attr_match.group(1)

        return fallback_data

    def flatten_element(element, prefix="detail"):
        data = {}
        for attr_name, attr_value in element.attrib.items():
            data[f"{prefix}_{element.tag}_{attr_name}"] = attr_value

        if element.text and element.text.strip():
            data[f"{prefix}_{element.tag}_text"] = element.text.strip()

        for child in element:
            child_data = flatten_element(child, prefix=f"{prefix}_{element.tag}")
            data.update(child_data)
        return data

    for idx, event_str in enumerate(events[:-1]):
        event_str = event_str.strip() + '</event>'
        if not event_str.strip():
            continue

        try:
            root = ET.fromstring(event_str)
        except ET.ParseError:
            log_callback(f"XML parse failed at index {idx}, using fallback extraction...")
            fallback_data = fallback_extract_data(event_str)
            uid = fallback_data.get("uid", "UNKNOWN_UID")
            if uid not in uid_data:
                uid_data[uid] = []
            uid_data[uid].append(fallback_data)
            continue

        uid = root.get("uid", "UNKNOWN_UID")
        if uid not in uid_data:
            uid_data[uid] = []

        event_data = {}
        detail = root.find("detail")
        if detail is not None:
            detail_data = flatten_element(detail, prefix="detail")
            event_data.update(detail_data)

        point = root.find("point")
        if point is not None:
            for attr in ["lat", "lon", "hae", "ce", "le"]:
                val = point.get(attr)
                if val:
                    event_data[f"point_{attr}"] = val

        # Parse remarks line-by-line if we have a remarks text field
        remarks_keys = [k for k in event_data.keys() if 'remarks_text' in k]
        for rk in remarks_keys:
            remarks_content = event_data[rk]
            parse_remarks_line_by_line(remarks_content, event_data)
            # Optionally remove raw remarks text column if desired
            # del event_data[rk]

        # Determine if we have a callsign
        callsign_key = [k for k in event_data.keys() if 'contact_callsign' in k]
        if callsign_key:
            event_data['contact_callsign'] = event_data[callsign_key[0]]

        uid_data[uid].append(event_data)

        if idx % 1000 == 0 and idx > 0:
            log_callback(f"Processed {idx}/{total} events...")

    def sanitize_sheet_name(sheet_name):
        sanitized = re.sub(r"[\\/*?:\[\]]", "", sheet_name)
        return sanitized[:31]

    script_dir = os.path.dirname(filePath)
    output_file = os.path.join(script_dir, "CoT_Data.xlsx")
    log_callback("Writing extracted CoT details to Excel...")

    sheets_created = False
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for uid, records in uid_data.items():
            if not records:
                continue

            # Determine sheet name
            sheet_name = uid
            for record in records:
                if 'contact_callsign' in record and record['contact_callsign']:
                    sheet_name = record['contact_callsign']
                    break

            sanitized_sheet_name = sanitize_sheet_name(sheet_name)
            df = pd.DataFrame(records)
            df.to_excel(writer, sheet_name=sanitized_sheet_name, index=False)
            sheets_created = True

        if not sheets_created:
            log_callback("No valid event data found. Creating a dummy sheet.")
            df = pd.DataFrame({"Info": ["No data available."]})
            df.to_excel(writer, sheet_name="No_Data", index=False)

    log_callback(f"Excel file created: {output_file}")

#############################
# GUI Application
#############################
class CoTParserGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("CoT Data Processor")
        self.geometry("800x800")

        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        main_frame = ctk.CTkFrame(self)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)

        header_label = ctk.CTkLabel(main_frame, text="CoT Data Processor", font=("Arial", 20, "bold"))
        header_label.grid(row=0, column=0, pady=10)

        buttons_frame = ctk.CTkFrame(main_frame)
        buttons_frame.grid(row=1, column=0, sticky="nsew", pady=10)
        buttons_frame.grid_columnconfigure(0, weight=1)
        for i in range(8):
            buttons_frame.grid_rowconfigure(i, weight=0)

        self.load_button = ctk.CTkButton(buttons_frame, text="Load File/Folder", command=self.load_file, width=200)
        self.load_button.grid(row=0, column=0, pady=10)

        self.remove_dup_button = ctk.CTkButton(buttons_frame, text="Remove Duplicates", command=self.remove_duplicates_action, width=200)
        self.remove_dup_button.grid(row=1, column=0, pady=10)

        self.adjust_time_button = ctk.CTkButton(buttons_frame, text="Adjust Event Times", command=self.adjust_times_action, width=200)
        self.adjust_time_button.grid(row=2, column=0, pady=10)

        self.reduce_size_button = ctk.CTkButton(buttons_frame, text="Reduce File Size", command=self.reduce_size_action, width=200)
        self.reduce_size_button.grid(row=3, column=0, pady=10)

        self.callsigns_button = ctk.CTkButton(buttons_frame, text="Get Callsigns", command=self.callsigns_action, width=200)
        self.callsigns_button.grid(row=4, column=0, pady=10)

        self.export_cot_button = ctk.CTkButton(buttons_frame, text="Export CoT Details", command=self.export_cot_details_action, width=200)
        self.export_cot_button.grid(row=5, column=0, pady=10)

        self.return_home_button = ctk.CTkButton(buttons_frame, text="Return to Home", command=self.return_home_action, width=200)
        self.return_home_button.grid(row=6, column=0, pady=10)

        log_frame = ctk.CTkFrame(main_frame)
        log_frame.grid(row=2, column=0, sticky="nsew", pady=10)
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)

        self.log_text = ctk.CTkTextbox(log_frame, width=750, height=350)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.configure(state="disabled")

    def log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.configure(state="disabled")
        self.log_text.see("end")
        self.update_idletasks()

    def load_file(self):
        path = filedialog.askopenfilename(title="Select a File", filetypes=[("All files", "*.*")])
        if not path:
            path = filedialog.askdirectory(title="Select a Folder")
            if not path:
                return
        try:
            path = validateFilePath(path)
            if not os.path.exists(path):
                messagebox.showerror("Error", "The provided file or folder does not exist.")
                return
            loadFiles(path)
            self.log(f"Loaded files from: {path}")
        except ValueError as e:
            messagebox.showerror("Error", str(e))

    def remove_duplicates_action(self):
        if not GlobalState.selectedFiles:
            messagebox.showwarning("No File", "No file selected. Please load a file first.")
            return
        file = GlobalState.selectedFiles[0]
        removeDuplicates(file, self.log)

    def adjust_times_action(self):
        if not GlobalState.selectedFiles:
            messagebox.showwarning("No File", "No file selected. Please load a file first.")
            return
        file = GlobalState.selectedFiles[0]
        dialog = ctk.CTkInputDialog(text="Enter new base time in Zulu format (e.g. 2024-09-16T17:13:48Z):", title="Adjust Event Times")
        new_time = dialog.get_input()
        if not new_time:
            return
        adjustEventTimes(file, new_time, self.log)

    def reduce_size_action(self):
        if not GlobalState.selectedFiles:
            messagebox.showwarning("No File", "No file selected. Please load a file first.")
            return
        file = GlobalState.selectedFiles[0]
        dialog = ctk.CTkInputDialog(text="Enter max file size (MB, up to 100):", title="Reduce File Size")
        result = dialog.get_input()
        if not result:
            return
        splitAndExportFile(file, result, self.log)

    def callsigns_action(self):
        if not GlobalState.selectedFiles:
            messagebox.showwarning("No File", "No file selected. Please load a file first.")
            return
        file = GlobalState.selectedFiles[0]
        extractUIDsAndCallsigns(file, self.log)

    def export_cot_details_action(self):
        if not GlobalState.selectedFiles:
            messagebox.showwarning("No File", "No file selected. Please load a file first.")
            return
        file = GlobalState.selectedFiles[0]
        export_cot_details(file, self.log)

    def return_home_action(self):
        self.destroy()
        # Import home_page here to avoid circular imports at the top
        import Home_Page
        Home_Page.open_home_page()

if __name__ == "__main__":
    try:
        app = CoTParserGUI()
        app.mainloop()
    except Exception as e:
        print("Uncaught exception:", e)
        traceback.print_exc()
