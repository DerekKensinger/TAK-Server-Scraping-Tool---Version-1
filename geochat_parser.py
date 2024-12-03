import customtkinter as ctk
import sqlite3
import pandas as pd
from datetime import datetime, timezone, timedelta
import os
from tkinter import filedialog, messagebox

import Home_Page


class GeoChatParserGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window settings
        self.title("GeoChat Parser")
        self.geometry("800x600")

        # Set up main frame
        frame = ctk.CTkFrame(self)
        frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Title label
        title_label = ctk.CTkLabel(frame, text="GeoChat Parser", font=("Arial", 24))
        title_label.pack(pady=10)

        # Subtitle label
        subtitle_label = ctk.CTkLabel(frame, text="Connect to the server you are wanting the GeoChat Logs from on WinTAK before using the parser.",
                                      font=("Arial", 14))  # Smaller text for subtitle
        subtitle_label.pack(pady=(0, 20))  # Add padding to create space below the subtitle

        # Input for database path - Start with an empty value
        self.db_path_var = ctk.StringVar(value="")  # Initialize with an empty string

        db_label = ctk.CTkLabel(frame, text="Database Path:")
        db_label.pack(pady=(10, 0))

        db_entry = ctk.CTkEntry(frame, textvariable=self.db_path_var, width=400)
        db_entry.pack(pady=(0, 10))

        browse_button = ctk.CTkButton(frame, text="Browse", command=self.browse_db)
        browse_button.pack(pady=(0, 10))

        # Dropdown menu for timezone selection
        self.timezone_var = ctk.StringVar(value="EST")  # Default to EST
        timezone_label = ctk.CTkLabel(frame, text="Select Timezone:")
        timezone_label.pack(pady=(10, 0))
        
        timezone_menu = ctk.CTkOptionMenu(frame, variable=self.timezone_var, values=["EST", "CST", "MST", "PST"])
        timezone_menu.pack(pady=(0, 10))

        # Start Parsing Button
        start_button = ctk.CTkButton(frame, text="Start Parsing", command=self.start_parsing)
        start_button.pack(pady=20)

        # Return to Home Page Button
        return_button = ctk.CTkButton(frame, text="Return to Home Page", command=self.return_to_home)
        return_button.pack(pady=20)

        self.mainloop()

    def browse_db(self):
        # Open a file dialog to select the database file
        file_path = filedialog.askopenfilename(filetypes=[("SQLite Database", "*.sqlite")])
        if file_path:
            self.db_path_var.set(file_path)

    def start_parsing(self):
        db_path = self.db_path_var.get()

        # Check if the database file exists
        if not os.path.exists(db_path):
            messagebox.showerror("Error", f"The database file was not found at {db_path}")
            return

        # Get the selected timezone from the dropdown
        selected_timezone = self.timezone_var.get()

        # Define timezone offsets based on the selected timezone
        timezone_offsets = {
            "EST": -5,
            "CST": -6,
            "MST": -7,
            "PST": -8
        }
        tz_offset = timezone_offsets.get(selected_timezone, -5)  # Default to EST if not found

        try:
            # Connect to the SQLite database
            conn = sqlite3.connect(db_path)

            # Query to extract the necessary columns from the Chat and Groups tables
            chat_query = """
            SELECT conversationId, receiveTime, sentTime, message, senderCallsign, status
            FROM Chat
            """
            groups_query = """
            SELECT conversationId, conversationName, createdLocally, destinations, parent
            FROM Groups
            """

            # Load the data into pandas DataFrames
            chat_df = pd.read_sql_query(chat_query, conn)
            groups_df = pd.read_sql_query(groups_query, conn)

            # Close the connection
            conn.close()

            # Function to convert the UNIX timestamp to the selected timezone
            def convert_to_timezone(unix_timestamp):
                if pd.isna(unix_timestamp):
                    return None
                utc_time = datetime.fromtimestamp(unix_timestamp / 1000, tz=timezone.utc)  # Convert from milliseconds to seconds
                local_time = utc_time.astimezone(timezone(timedelta(hours=tz_offset)))  # Convert to the selected timezone
                return local_time.strftime('%Y-%m-%d %H:%M:%S')

            # Apply the conversion function to receiveTime and sentTime columns in the Chat DataFrame
            chat_df['receiveTime'] = chat_df['receiveTime'].apply(convert_to_timezone)
            chat_df['sentTime'] = chat_df['sentTime'].apply(convert_to_timezone)

            # Create a dictionary mapping conversationId to conversationName
            conversation_mapping = dict(zip(groups_df['conversationId'], groups_df['conversationName']))

            # Replace conversationId in chat_df with corresponding conversationName, if available
            chat_df['conversationId'] = chat_df['conversationId'].map(conversation_mapping).fillna(chat_df['conversationId'])

            # Define the path for the folder and Excel file
            home_dir = os.path.expanduser("~")
            folder_path = os.path.join(home_dir, 'Desktop', 'GeoChat Logs')  # Folder on Desktop

            # Check if the folder exists; if not, create it
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)

            # Add the current date and time to the file name
            current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")  # Format: YYYY-MM-DD_HH-MM-SS
            excel_file_path = os.path.join(folder_path, f'GeoChat Logs_{current_time}.xlsx')

            # Use ExcelWriter to save multiple DataFrames to separate sheets in the same Excel file
            with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
                chat_df.to_excel(writer, sheet_name='Chat Data', index=False)
                groups_df.to_excel(writer, sheet_name='Groups Data', index=False)

            messagebox.showinfo("Success", f"Data has been successfully exported to {excel_file_path} with separate tabs for Chat and Groups data.")
        
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")


    def return_to_home(self):
        # Close the current GeoChat Parser window
        self.destroy()

        # Open the Home Page
        Home_Page.open_home_page()  # Assuming you have a function in your home_page script that opens the main page

if __name__ == "__main__":
    app = GeoChatParserGUI()
