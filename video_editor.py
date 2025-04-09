# Imports
import customtkinter as ctk
from tkinter import messagebox, filedialog
import subprocess
import os
import re
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
import threading
import platform
import Home_Page

# Input dialog class
class CustomInputDialog(ctk.CTkToplevel):
    def __init__(self, parent, title="Input", prompt="Enter value:"):
        super().__init__(parent)
        self.title(title)
        self.geometry("500x150")
        self.resizable(False, False)
        self.grab_set()

        self.result = None

        ctk.CTkLabel(self, text=prompt).pack(pady=(20, 10))
        self.entry = ctk.CTkEntry(self, width=400)
        self.entry.pack(pady=5)
        self.entry.focus()

        button_frame = ctk.CTkFrame(self)
        button_frame.pack(pady=10)

        ctk.CTkButton(button_frame, text="OK", command=self.on_ok).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Cancel", command=self.on_cancel).pack(side="right", padx=10)

        self.bind("<Return>", lambda event: self.on_ok())
        self.bind("<Escape>", lambda event: self.on_cancel())

    def on_ok(self):
        self.result = self.entry.get()
        self.destroy()

    def on_cancel(self):
        self.result = None
        self.destroy()

# Main GUI
class VideoEditorGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Video Editor")
        self.geometry("800x700")

        self.last_output_folder = None

        main_frame = ctk.CTkFrame(self)
        main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        ctk.CTkLabel(main_frame, text="Video Editor", font=("Arial", 24)).pack(pady=10)

        button_texts = [
            "Convert Video From MKV to MP4",
            "Compress A Video (Resolution, Framerate, Audio)",
            "Combine Clips",
            "Split A Clip",
            "Clip a Clip",
            "Generate A GIF"
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
            ctk.CTkButton(main_frame, text=text, command=command).pack(pady=5)

        ctk.CTkButton(main_frame, text="Return to Home Page", command=self.return_to_home).pack(pady=20)

        self.log_textbox = ctk.CTkTextbox(main_frame, height=200)
        self.log_textbox.pack(pady=10, fill="both", expand=True)

        # Per-file progress
        self.progress_var = ctk.DoubleVar(value=0)
        self.progress_bar = ctk.CTkProgressBar(main_frame, variable=self.progress_var, mode="determinate")
        self.progress_bar.pack(pady=(0, 5), fill="x")
        self.progress_bar.set(0)
        self.progress_label = ctk.CTkLabel(main_frame, text="Progress: 0%")
        self.progress_label.pack()

        # Folder-level progress
        self.total_progress_var = ctk.DoubleVar(value=0)
        self.total_progress_bar = ctk.CTkProgressBar(main_frame, variable=self.total_progress_var, mode="determinate")
        self.total_progress_bar.pack(pady=(5, 0), fill="x")
        self.total_progress_bar.set(0)
        self.total_progress_label = ctk.CTkLabel(main_frame, text="Folder Progress: 0%")
        self.total_progress_label.pack()

        # Output folder button
        self.output_button = ctk.CTkButton(main_frame, text="📂 Open Output Folder", command=self.open_output_folder)
        self.output_button.pack(pady=5)
        self.output_button.pack_forget()

        self.mainloop()

    def log(self, message):
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")

    def safe_log(self, message):
        self.after(0, lambda: self.log(message))

    def threaded(fn):
        def wrapper(self, *args, **kwargs):
            threading.Thread(target=fn, args=(self, *args), kwargs=kwargs).start()
        return wrapper

    def open_output_folder(self):
        if self.last_output_folder and os.path.isdir(self.last_output_folder):
            try:
                if platform.system() == "Windows":
                    os.startfile(self.last_output_folder)
                elif platform.system() == "Darwin":
                    subprocess.Popen(["open", self.last_output_folder])
                else:
                    subprocess.Popen(["xdg-open", self.last_output_folder])
            except Exception as e:
                self.safe_log(f"⚠️ Failed to open folder: {e}")
        else:
            self.safe_log("❌ Output folder not found.")

    def update_progress_bar(self, current_sec, total_sec, frame_line, frame_index=None):
        progress = min(current_sec / total_sec, 1.0)
        try:
            if frame_index:
                self.log_textbox.delete(frame_index, f"{frame_index} + 1 lines")
            else:
                frame_index = self.log_textbox.index("end-1c")
            self.log_textbox.insert(frame_index, frame_line + "\n")
            self.log_textbox.see("end")
            self.progress_var.set(progress)
            self.progress_label.configure(text=f"Progress: {int(progress * 100)}%")
        except Exception:
            pass
        return frame_index

    def update_folder_progress(self, current, total):
        try:
            progress = current / total
            self.total_progress_var.set(progress)
            self.total_progress_label.configure(text=f"Folder Progress: {int(progress * 100)}%")
        except Exception:
            pass

    def return_to_home(self):
        self.destroy()
        Home_Page.open_home_page()

    @threaded
    def convert_to_mp4(self):
        import re
        from pathlib import Path

        def parse_ffmpeg_time(time_str):
            h, m, s = time_str.split(':')
            sec, micro = s.split('.')
            return int(h) * 3600 + int(m) * 60 + int(sec) + int(micro) / 100

        def get_video_duration(input_path):
            try:
                result = subprocess.run([
                    "ffprobe", "-v", "error",
                    "-show_entries", "format=duration",
                    "-of", "default=noprint_wrappers=1:nokey=1",
                    str(input_path)
                ], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, check=True)
                return float(result.stdout.strip())
            except Exception as e:
                self.safe_log(f"⚠️ Could not get duration: {e}")
                return 1

        file_or_dir = filedialog.askopenfilename(
            title="Select MKV File or Cancel to Pick Folder",
            filetypes=[("MKV files", "*.mkv")]
        )

        paths_to_convert = []
        if file_or_dir:
            input_path = Path(file_or_dir)
            if input_path.suffix.lower() == ".mkv":
                paths_to_convert = [input_path]
        else:
            folder = filedialog.askdirectory(title="Select Folder with MKV Files")
            if folder:
                folder_path = Path(folder)
                paths_to_convert = list(folder_path.glob("*.mkv"))

        if not paths_to_convert:
            messagebox.showinfo("No Files", "No MKV files selected or found.")
            return

        total_files = len(paths_to_convert)
        self.after(0, lambda: self.total_progress_var.set(0))
        self.after(0, lambda: self.total_progress_label.configure(text="Folder Progress: 0%"))

        for index, input_path in enumerate(paths_to_convert):
            self.after(0, self.output_button.pack_forget)
            output_path = input_path.with_suffix(".mp4")
            self.safe_log(f"🎬 Converting: {input_path.name}")
            self.after(0, lambda: self.progress_bar.set(0))
            self.after(0, lambda: self.progress_label.configure(text="Progress: 0%"))

            total_duration = get_video_duration(input_path)

            try:
                cmd = [
                    "ffmpeg", "-y", "-i", str(input_path),
                    str(output_path)
                ]

                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1
                )

                time_pattern = re.compile(r"time=(\d{2}:\d{2}:\d{2}\.\d{2})")
                frame_line_index = None

                for line in process.stdout:
                    line = line.strip()
                    if not line:
                        continue
                    match = time_pattern.search(line)
                    if match:
                        current_time_str = match.group(1)
                        current_time_sec = parse_ffmpeg_time(current_time_str)
                        self.after(0, lambda s=current_time_sec, t=total_duration, l=line: 
                                self.update_progress_bar(s, t, l))
                    else:
                        self.safe_log(line)

                process.wait()
                if process.returncode != 0:
                    raise subprocess.CalledProcessError(process.returncode, cmd)

                self.safe_log(f"✅ Done: {output_path.name}")
                self.last_output_folder = str(output_path.parent)
                self.after(0, self.output_button.pack)
                self.after(0, lambda: self.progress_var.set(1))
                self.after(0, lambda: self.progress_label.configure(text="Progress: 100%"))
                self.after(0, lambda: self.update_folder_progress(index + 1, total_files))

            except Exception as e:
                self.safe_log(f"❌ Failed to convert {input_path.name}: {e}")
                self.after(0, lambda: self.progress_var.set(0))
                self.after(0, lambda: self.progress_label.configure(text="Progress: 0%"))

        messagebox.showinfo("Conversion Complete", "All files processed.")

    @threaded
    def compress_video(self):
        from pathlib import Path

        folder = filedialog.askdirectory(title="Select Folder Containing MP4 Files")
        if not folder:
            return

        resolution_options = {
            "144p": "256:-2", "240p": "426:-2", "360p": "640:-2",
            "480p": "854:-2", "720p": "1280:-2", "1080p": "1920:-2",
            "1440p": "2560:-2", "4K (2160p)": "3840:-2"
        }

        resolution_var = ctk.StringVar(value="720p")
        res_window = ctk.CTkToplevel(self)
        res_window.title("Select Resolution")
        res_window.geometry("300x150")
        res_window.grab_set()

        ctk.CTkLabel(res_window, text="Select target resolution:").pack(pady=10)
        ctk.CTkOptionMenu(res_window, values=list(resolution_options.keys()), variable=resolution_var).pack(pady=10)
        ctk.CTkButton(res_window, text="Confirm", command=res_window.destroy).pack(pady=10)
        self.wait_window(res_window)
        scale_value = resolution_options.get(resolution_var.get(), "1920:-2")

        dialog_crf = CustomInputDialog(self, title="Compression Quality", prompt="Enter CRF value (18–28, default: 23):")
        self.wait_window(dialog_crf)
        try: crf = int(dialog_crf.result or 23)
        except: crf = 23

        dialog_fps = CustomInputDialog(self, title="Frame Rate", prompt="Enter target FPS (default: 30):")
        self.wait_window(dialog_fps)
        try: fps = int(dialog_fps.result or 30)
        except: fps = 30

        dialog_audio = CustomInputDialog(self, title="Audio Option", prompt="Remove audio? (yes/no, default: yes):")
        self.wait_window(dialog_audio)
        remove_audio = (dialog_audio.result or "yes").strip().lower() != "no"

        video_files = list(Path(folder).rglob("*.mp4"))
        if not video_files:
            self.safe_log("❌ No MP4 files found in the selected folder.")
            return

        total_files = len(video_files)
        self.after(0, lambda: self.total_progress_var.set(0))
        self.after(0, lambda: self.total_progress_label.configure(text="Folder Progress: 0%"))

        for index, video_file in enumerate(video_files):
            self.after(0, lambda: self.progress_bar.set(0))
            self.after(0, lambda: self.progress_label.configure(text="Progress: 0%"))
            self.safe_log(f"▶️ [{index + 1}/{len(video_files)}] Compressing: {video_file.name}")

            try:
                final_mp4_path = video_file.parent / f"Compressed_{video_file.stem}.mp4"
                cmd = [
                    "ffmpeg", "-i", str(video_file),
                    "-vf", f"scale={scale_value}",
                    "-r", str(fps),
                    "-c:v", "libx264",
                    "-crf", str(crf),
                    "-preset", "slow"
                ]
                if remove_audio:
                    cmd.append("-an")
                cmd.append(str(final_mp4_path))

                process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
                frame_line_index = None
                for line in process.stdout:
                    line = line.strip()
                    if "time=" in line:
                        match = re.search(r"time=(\d{2}:\d{2}:\d{2}\.\d{2})", line)
                        if match:
                            t = match.group(1)
                            h, m, s = t.split(":")
                            s, ms = s.split(".")
                            current_time = int(h) * 3600 + int(m) * 60 + int(s) + int(ms) / 100
                            total_time = 60 * 10  # Optional fallback
                            self.after(0, lambda ct=current_time, tt=total_time, l=line:
                                    self.update_progress_bar(ct, tt, l))
                    else:
                        self.safe_log(line)

                process.wait()
                if process.returncode != 0:
                    raise subprocess.CalledProcessError(process.returncode, cmd)

                self.safe_log(f"✅ Saved: {final_mp4_path.name}")
                self.after(0, lambda: self.progress_var.set(1))
                self.after(0, lambda: self.progress_label.configure(text="Progress: 100%"))
                self.after(0, lambda: self.update_folder_progress(index + 1, total_files))

            except Exception as e:
                self.safe_log(f"   ❌ Failed: {video_file.name} — {e}")
                self.after(0, lambda: self.progress_var.set(0))
                self.after(0, lambda: self.progress_label.configure(text="Progress: 0%"))

        self.safe_log("🎉 Compression complete.")

    @threaded
    def combine_clips(self):
        from pathlib import Path

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

            self.after(0, lambda: self.progress_bar.set(0))
            self.after(0, lambda: self.progress_label.configure(text="Progress: 0%"))

            self.after(0, lambda: self.total_progress_var.set(0))
            self.after(0, lambda: self.total_progress_label.configure(text="Folder Progress: 0%"))

            output_path = generate_output_path(root_dir, base_name, video_files[0])
            filelist_path = create_ffmpeg_filelist(video_files)

            self.safe_log(f"🎞️ Combining {len(video_files)} clips...")

            cmd = [
                "ffmpeg", "-f", "concat", "-safe", "0",
                "-i", filelist_path,
                "-c", "copy",
                output_path
            ]

            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            frame_line_index = None

            for line in process.stdout:
                line = line.strip()
                if not line:
                    continue
                if "frame=" in line or "time=" in line:
                    match = re.search(r"time=(\d{2}:\d{2}:\d{2}\.\d{2})", line)
                    if match:
                        t = match.group(1)
                        h, m, s = t.split(":")
                        s, ms = s.split(".")
                        current_sec = int(h) * 3600 + int(m) * 60 + int(s) + int(ms) / 100
                        total_estimate = 10 * 60  # Simulate duration for now
                        self.after(0, lambda: self.update_progress_bar(current_sec, total_estimate, line))
                else:
                    self.safe_log(line)

            process.wait()
            if process.returncode != 0:
                raise subprocess.CalledProcessError(process.returncode, cmd)

            os.remove(filelist_path)

            self.last_output_folder = str(Path(output_path).parent)
            self.after(0, self.output_button.pack)
            self.after(0, lambda: self.progress_var.set(1))
            self.after(0, lambda: self.progress_label.configure(text="Progress: 100%"))
            self.after(0, lambda: self.total_progress_var.set(1))
            self.after(0, lambda: self.total_progress_label.configure(text="Folder Progress: 100%"))

            msg = f"✅ Successfully merged {len(video_files)} clips to:\n{output_path}"
            if skipped:
                msg += f"\n\n⚠️ Skipped {len(skipped)} files due to filtering."
            messagebox.showinfo("Merge Complete", msg)

        except subprocess.CalledProcessError as e:
            messagebox.showerror("FFmpeg Error", f"FFmpeg failed:\n{e}")
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")

    @threaded
    def split_clip(self):
        from pathlib import Path

        video_path = filedialog.askopenfilename(title="Select MP4 File to Split", filetypes=[("MP4 files", "*.mp4")])
        if not video_path:
            return

        dialog_time = CustomInputDialog(self, title="Split Offset", prompt="Enter time offset (HH:MM:SS):")
        self.wait_window(dialog_time)
        timestamp = dialog_time.result
        if not timestamp:
            messagebox.showerror("Input Missing", "Time offset is required.")
            return

        self.safe_log(f"⏳ Splitting at {timestamp}...")

        try:
            input_file = Path(video_path)
            output_file = input_file.with_name(f"Split_{input_file.stem}.mp4")

            self.after(0, lambda: self.progress_bar.set(0))
            self.after(0, lambda: self.progress_label.configure(text="Progress: 0%"))

            cmd = [
                "ffmpeg", "-ss", timestamp, "-i", str(input_file),
                "-c", "copy", str(output_file)
            ]

            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            for line in process.stdout:
                self.safe_log(line.strip())

            process.wait()
            if process.returncode != 0:
                raise subprocess.CalledProcessError(process.returncode, cmd)

            self.after(0, lambda: self.progress_bar.set(1))
            self.after(0, lambda: self.progress_label.configure(text="Progress: 100%"))
            self.safe_log(f"✅ Split saved as: {output_file}")

        except Exception as e:
            self.safe_log(f"❌ Failed to split video: {e}")
            messagebox.showerror("Error", f"Failed to split video:\n{e}")

    @threaded
    def clip_clip(self):
        from pathlib import Path

        video_path = filedialog.askopenfilename(title="Select MP4 File to Clip", filetypes=[("MP4 files", "*.mp4")])
        if not video_path:
            return

        dialog_start = CustomInputDialog(self, title="Start Time", prompt="Enter start time (e.g., 00:01:00):")
        self.wait_window(dialog_start)
        start_time = dialog_start.result
        if not start_time:
            messagebox.showerror("Input Missing", "Start time is required.")
            return

        dialog_end = CustomInputDialog(self, title="End Time", prompt="Enter end time (e.g., 00:02:00):")
        self.wait_window(dialog_end)
        end_time = dialog_end.result
        if not end_time:
            messagebox.showerror("Input Missing", "End time is required.")
            return

        self.safe_log(f"⏳ Clipping from {start_time} to {end_time}...")

        try:
            input_file = Path(video_path)
            output_file = input_file.with_name(f"Clipped_{input_file.stem}.mp4")

            self.after(0, lambda: self.progress_bar.set(0))
            self.after(0, lambda: self.progress_label.configure(text="Progress: 0%"))

            cmd = [
                "ffmpeg", "-ss", start_time, "-to", end_time,
                "-i", str(input_file), "-c", "copy", str(output_file)
            ]

            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            for line in process.stdout:
                self.safe_log(line.strip())

            process.wait()
            if process.returncode != 0:
                raise subprocess.CalledProcessError(process.returncode, cmd)

            self.after(0, lambda: self.progress_bar.set(1))
            self.after(0, lambda: self.progress_label.configure(text="Progress: 100%"))
            self.safe_log(f"✅ Clipped segment saved as: {output_file}")

        except Exception as e:
            self.safe_log(f"❌ Failed to clip video: {e}")
            messagebox.showerror("Error", f"Failed to clip video:\n{e}")

    @threaded
    def generate_gif(self):
        from pathlib import Path

        def get_video_duration(path):
            try:
                result = subprocess.run(
                    ["ffprobe", "-v", "error", "-show_entries", "format=duration",
                    "-of", "default=noprint_wrappers=1:nokey=1", str(path)],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=True
                )
                return float(result.stdout.strip())
            except Exception as e:
                self.safe_log(f"⚠️ Could not determine video duration: {e}")
                return None

        def parse_time_to_seconds(timestr):
            h, m, s = map(float, timestr.split(":"))
            return h * 3600 + m * 60 + s

        video_path = filedialog.askopenfilename(
            title="Select MP4 File to Convert to GIF", filetypes=[("MP4 files", "*.mp4")]
        )
        if not video_path:
            return

        dialog_start = CustomInputDialog(self, title="GIF Start Time", prompt="Enter start time (e.g., 00:00:02):")
        self.wait_window(dialog_start)
        start_time = dialog_start.result
        if not start_time:
            messagebox.showerror("Input Missing", "Start time is required.")
            return

        dialog_duration = CustomInputDialog(self, title="GIF Duration", prompt="Enter duration in seconds (e.g., 12):")
        self.wait_window(dialog_duration)
        duration = dialog_duration.result
        if not duration:
            messagebox.showerror("Input Missing", "Duration is required.")
            return

        try:
            duration_sec = float(duration)
            start_sec = parse_time_to_seconds(start_time)
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid time format and numeric duration.")
            return

        total_video_duration = get_video_duration(video_path)
        if total_video_duration and (start_sec + duration_sec > total_video_duration):
            messagebox.showerror("Invalid Range", "The GIF range exceeds the video's total duration.")
            return

        self.safe_log(f"🎬 Generating GIF from {start_time} for {duration} seconds...")

        try:
            input_file = Path(video_path)
            output_gif = input_file.with_name(f"GIF_{input_file.stem}.gif")
            palette_path = input_file.with_name("palette.png")

            self.after(0, lambda: self.progress_bar.set(0))
            self.after(0, lambda: self.progress_label.configure(text="Progress: 0%"))

            cmd_palette = [
                "ffmpeg", "-ss", start_time, "-t", str(duration), "-i", str(input_file),
                "-vf", "fps=15,scale=1920:-1:force_original_aspect_ratio=decrease,palettegen",
                str(palette_path)
            ]
            self.safe_log("🖌️ Generating palette...")
            subprocess.run(cmd_palette, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)

            cmd_gif = [
                "ffmpeg", "-ss", start_time, "-t", str(duration),
                "-i", str(input_file), "-i", str(palette_path),
                "-filter_complex", "[0:v] fps=15,scale=1920:-1:force_original_aspect_ratio=decrease [x]; [x][1:v] paletteuse",
                "-loop", "0", str(output_gif)
            ]
            self.safe_log("🎞️ Creating GIF...")
            subprocess.run(cmd_gif, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)

            if palette_path.exists():
                palette_path.unlink()

            self.after(0, lambda: self.progress_bar.set(1))
            self.after(0, lambda: self.progress_label.configure(text="Progress: 100%"))
            self.safe_log(f"✅ GIF saved as: {output_gif}")

        except Exception as e:
            self.safe_log(f"❌ Failed to generate GIF: {e}")
            messagebox.showerror("Error", f"Failed to generate GIF:\n{e}")

    def return_to_home(self):
        self.destroy()
        Home_Page.open_home_page()

if __name__ == "__main__":
    VideoEditorGUI()
