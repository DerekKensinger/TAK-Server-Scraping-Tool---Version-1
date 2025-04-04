import customtkinter as ctk
from tkinter import messagebox
import subprocess
import Home_Page
import os
import re
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog

class CustomInputDialog(ctk.CTkToplevel):
    def __init__(self, parent, title="Input", prompt="Enter value:"):
        super().__init__(parent)
        self.title(title)
        self.geometry("500x150")
        self.resizable(False, False)
        self.grab_set()

        self.result = None

        self.label = ctk.CTkLabel(self, text=prompt)
        self.label.pack(pady=(20, 10))

        self.entry = ctk.CTkEntry(self, width=400)
        self.entry.pack(pady=5)
        self.entry.focus()

        button_frame = ctk.CTkFrame(self)
        button_frame.pack(pady=10)

        ok_button = ctk.CTkButton(button_frame, text="OK", command=self.on_ok)
        ok_button.pack(side="left", padx=10)

        cancel_button = ctk.CTkButton(button_frame, text="Cancel", command=self.on_cancel)
        cancel_button.pack(side="right", padx=10)

        self.bind("<Return>", lambda event: self.on_ok())
        self.bind("<Escape>", lambda event: self.on_cancel())

    def on_ok(self):
        self.result = self.entry.get()
        self.destroy()

    def on_cancel(self):
        self.result = None
        self.destroy()

class VideoEditorGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Video Editor")
        self.geometry("800x600")

        main_frame = ctk.CTkFrame(self)
        main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        title_label = ctk.CTkLabel(main_frame, text="Video Editor", font=("Arial", 24))
        title_label.pack(pady=10)

        button_texts = [
            "Convert Video to mp4 1:1",
            "Compressing of video (resolution, framerate, audio)",
            "Combining of clips",
            "Split clip",
            "Clip a clip",
            "Generate gif"
        ]

        commands = [
            self.convert_to_mp4,
            self.compress_video,
            self.combine_clips,
            self.split_clip,
            self.clip_clip,
            self.generate_gif
        ]

        for text, command in zip(button_texts, commands):
            btn = ctk.CTkButton(main_frame, text=text, command=command)
            btn.pack(pady=5)

        return_button = ctk.CTkButton(main_frame, text="Return to Home Page", command=self.return_to_home)
        return_button.pack(pady=20)

        self.mainloop()

    def convert_to_mp4(self):
        messagebox.showinfo("Convert", "This will convert video to mp4 - implementation pending.")

    def compress_video(self):
        messagebox.showinfo("Compress", "This will compress the video - implementation pending.")

    def combine_clips(self):
        def extract_timestamp(filename):
            match = re.search(r"\d{4}-\d{2}-\d{2}T\d{2}_\d{2}_\d{2}", filename)
            if match:
                try:
                    return datetime.strptime(match.group(), "%Y-%m-%dT%H_%M_%S")
                except ValueError:
                    return None
            return None

        def collect_filtered_and_sorted_videos(root_dir, required_substring="realtime_720p"):
            video_files = list(Path(root_dir).rglob("*.mp4"))
            video_with_timestamps = []
            skipped_files = []

            for file in video_files:
                if required_substring not in str(file):
                    skipped_files.append((file, f"Skipped: Does not contain '{required_substring}'"))
                    continue

                if file.name.startswith("._") or file.stat().st_size == 0:
                    skipped_files.append((file, "Skipped: MacOS metadata or empty file"))
                    continue

                ts = extract_timestamp(file.name)
                if ts:
                    video_with_timestamps.append((ts, file))
                else:
                    skipped_files.append((file, "Skipped: No valid timestamp"))

            sorted_videos = sorted(video_with_timestamps, key=lambda x: x[0])
            sorted_video_paths = [str(file) for _, file in sorted_videos]
            return sorted_video_paths, skipped_files

        def create_ffmpeg_filelist(video_paths):
            temp_file = tempfile.NamedTemporaryFile(mode='w+', delete=False, suffix='.txt', encoding='utf-8')
            for path in video_paths:
                temp_file.write(f"file '{path}'\n")
            temp_file.flush()
            return temp_file.name

        def merge_videos_ffmpeg(filelist_path, output_path):
            cmd = [
                "ffmpeg",
                "-f", "concat",
                "-safe", "0",
                "-i", filelist_path,
                "-c", "copy",
                output_path
            ]
            subprocess.run(cmd, check=True)

        def generate_output_path(root_dir, base_name, first_video_path):
            timestamp = extract_timestamp(Path(first_video_path).name)
            if not timestamp:
                suffix = "_unknown_time"
            else:
                adjusted_ts = timestamp - timedelta(hours=7)
                suffix = "_" + adjusted_ts.strftime("%Y-%m-%d_%H-%M-%S")
            filename = f"{base_name}{suffix}.mp4"
            return os.path.join(root_dir, filename)

        root_dir = filedialog.askdirectory(title="Select Folder Containing Clips")
        if not root_dir:
            return

        # Use custom CTk input dialog
        dialog = CustomInputDialog(self, title="Output Name", prompt="Enter base name for output file (no extension):")
        self.wait_window(dialog)
        base_name = dialog.result

        if not base_name:
            messagebox.showwarning("Input Missing", "Output name is required.")
            return

        try:
            video_files, skipped = collect_filtered_and_sorted_videos(root_dir)
            if not video_files:
                messagebox.showerror("No Valid Videos", "No matching video files found in selected folder.")
                return

            output_path = generate_output_path(root_dir, base_name, video_files[0])
            filelist_path = create_ffmpeg_filelist(video_files)
            merge_videos_ffmpeg(filelist_path, output_path)
            os.remove(filelist_path)

            msg = f"Successfully merged {len(video_files)} clips to:\n{output_path}"
            if skipped:
                msg += f"\n\nSkipped {len(skipped)} files due to filtering."

            messagebox.showinfo("Merge Complete", msg)
        except subprocess.CalledProcessError as e:
            messagebox.showerror("FFmpeg Error", f"FFmpeg failed:\n{e}")
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")

    def split_clip(self):
        messagebox.showinfo("Split", "This will split a clip - implementation pending.")

    def clip_clip(self):
        messagebox.showinfo("Clip", "This will cut a section of the video - implementation pending.")

    def generate_gif(self):
        messagebox.showinfo("GIF", "This will generate a gif from the video - implementation pending.")

    def return_to_home(self):
        self.destroy()
        Home_Page.open_home_page()

if __name__ == "__main__":
    VideoEditorGUI()
