import customtkinter as ctk
import tak_report_parser
import geochat_parser
import video_editor

def open_home_page():
    root = ctk.CTk()
    root.title("TAK Server Scraping Tool")
    root.geometry("600x400")

    main_frame = ctk.CTkFrame(master=root)
    main_frame.pack(pady=40, padx=40, fill="both", expand=True)

    title_label = ctk.CTkLabel(master=main_frame, text="TAK Server Scraping Tool", font=("Arial", 24))
    title_label.pack(pady=20)

    button_frame = ctk.CTkFrame(master=main_frame)
    button_frame.pack(pady=20)

    def open_script(script_name):
        root.destroy()
        if script_name == "tak_report_parser":
            tak_report_parser.TAKReportGUI()
        elif script_name == "geochat_parser":
            geochat_parser.GeoChatParserGUI()
        elif script_name == "video_editor":
            video_editor.VideoEditorGUI()
        elif script_name == "cot_parser":
            import cot_parser
            app = cot_parser.CoTParserGUI()
            app.mainloop()

    buttons_info = [
        ("TAK Report Parser", "tak_report_parser"),
        ("GeoChat Parser", "geochat_parser"),
        ("Video Editor", "video_editor"),
        ("CoT Data Processor", "cot_parser"),
    ]

    for i, (button_text, script_name) in enumerate(buttons_info):
        button = ctk.CTkButton(master=button_frame, text=button_text, 
                               command=lambda script=script_name: open_script(script))
        row, col = divmod(i, 2)
        button.grid(row=row, column=col, padx=15, pady=15, sticky="nsew")

    button_frame.grid_rowconfigure(0, weight=1)
    button_frame.grid_rowconfigure(1, weight=1)
    button_frame.grid_columnconfigure(0, weight=1)
    button_frame.grid_columnconfigure(1, weight=1)

    root.mainloop()

if __name__ == "__main__":
    open_home_page()
