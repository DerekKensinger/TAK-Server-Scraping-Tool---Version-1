
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
            btn = ctk.CTkButton(main_frame, text=text, command=command)
            btn.pack(pady=5)

        return_button = ctk.CTkButton(main_frame, text="Return to Home Page", command=self.return_to_home)
        return_button.pack(pady=20)

        # Log window
        self.log_textbox = ctk.CTkTextbox(main_frame, height=200)
        self.log_textbox.pack(pady=10, fill="both", expand=True)

        self.mainloop()

    def log(self, message):
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")  # Auto-scroll

    def convert_to_mp4(self):
        from tkinter import filedialog
        import os
        import subprocess
        from pathlib import Path

        def convert_file(input_path):
            output_path = input_path.with_suffix('.mp4')
            try:
                subprocess.run([
                    "ffmpeg",
                    "-i", str(input_path),
                    str(output_path)
                ], check=True)
                return True, str(output_path)
            except subprocess.CalledProcessError as e:
                return False, str(e)

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

        successes, failures = [], []

        for path in paths_to_convert:
            success, result = convert_file(path)
            if success:
                successes.append(result)
            else:
                failures.append((str(path), result))

        summary = f"Converted {len(successes)} file(s) successfully."
        if failures:
            summary += f"\nFailed to convert {len(failures)} file(s):"
            for f, reason in failures:
                summary += f"\n - {f}: {reason}"

        messagebox.showinfo("Conversion Complete", summary)

    
    def compress_video(self):
        from tkinter import filedialog
        import os

        # Step 1: Select folder
        folder = filedialog.askdirectory(title="Select Folder Containing MP4 Files")
        if not folder:
            return

        # Step 2: Prompt for CRF, FPS, Resolution, Audio removal

        # Resolution dropdown using CTkOptionMenu
        resolution_options = {
            "144p": "256:-2",
            "240p": "426:-2",
            "360p": "640:-2",
            "480p": "854:-2",
            "720p": "1280:-2",
            "1080p": "1920:-2",
            "1440p": "2560:-2",
            "4K (2160p)": "3840:-2"
        }

        resolution_var = ctk.StringVar(value="720p")
        res_window = ctk.CTkToplevel(self)
        res_window.title("Select Resolution")
        res_window.geometry("300x150")
        res_window.grab_set()

        label = ctk.CTkLabel(res_window, text="Select target resolution:")
        label.pack(pady=10)

        menu = ctk.CTkOptionMenu(res_window, values=list(resolution_options.keys()), variable=resolution_var)
        menu.pack(pady=10)

        confirm_btn = ctk.CTkButton(res_window, text="Confirm", command=res_window.destroy)
        confirm_btn.pack(pady=10)

        self.wait_window(res_window)
        selected_resolution = resolution_var.get()
        scale_value = resolution_options.get(selected_resolution, "1920:-2")
    
        dialog_crf = CustomInputDialog(self, title="Compression Quality", prompt="Enter CRF value (18–28, lower = better quality, default: 23):")
        self.wait_window(dialog_crf)
        try:
            crf = int(dialog_crf.result or 23)
        except (TypeError, ValueError):
            crf = 23

        dialog_fps = CustomInputDialog(self, title="Frame Rate", prompt="Enter target FPS (default: 30):")
        self.wait_window(dialog_fps)
        try:
            fps = int(dialog_fps.result or 30)
        except (TypeError, ValueError):
            fps = 30

        # Simple yes/no audio prompt
        dialog_audio = CustomInputDialog(self, title="Audio Option", prompt="Remove audio? (yes/no, default: yes):")
        self.wait_window(dialog_audio)
        remove_audio = (dialog_audio.result or "yes").strip().lower() != "no"

        # Step 3: Begin conversion
        from pathlib import Path
        import subprocess

        video_files = list(Path(folder).rglob("*.mp4"))
        if not video_files:
            self.log("❌ No MP4 files found in the selected folder.")
            return

        self.log(f"🔧 Starting compression: CRF={crf}, FPS={fps}, Remove Audio={remove_audio}")
        converted = []
        skipped = []

        for index, video_file in enumerate(video_files, start=1):
            self.log(f"▶️ [{index}/{len(video_files)}] Processing: {video_file.name}")
            try:
                base_name = video_file.stem
                parent_dir = video_file.parent

                final_mp4_path = parent_dir / f"Compressed_{base_name}.mp4"

                # Step 2: Apply compression
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

                self.log(f"   ⏳ Compressing...")
                process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
                for line in process.stdout:
                    self.log(line.strip())
                process.wait()
                if process.returncode != 0:
                    raise subprocess.CalledProcessError(process.returncode, process.args)

                self.log(f"   ✅ Saved: {final_mp4_path.name}")
                converted.append(str(final_mp4_path))

            except Exception as e:
                self.log(f"   ❌ Failed: {video_file.name} — {e}")
                skipped.append(str(video_file))

        self.log("🎉 Compression complete.")
        self.log(f"✅ Files converted: {len(converted)}")
        if skipped:
            self.log(f"⚠️ Files skipped: {len(skipped)}")

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
        from tkinter import filedialog
        import subprocess
        import os
        from pathlib import Path

        # Prompt for .mp4 file
        video_path = filedialog.askopenfilename(title="Select MP4 File to Split", filetypes=[("MP4 files", "*.mp4")])
        if not video_path:
            return

        # Prompt for time offset (format: HH:MM:SS)
        dialog_time = CustomInputDialog(self, title="Split Offset", prompt="Enter time offset (HH:MM:SS):")
        self.wait_window(dialog_time)
        timestamp = dialog_time.result

        if not timestamp:
            messagebox.showerror("Input Missing", "Time offset is required.")
            return

        self.log(f"⏳ Splitting at {timestamp}...")

        try:
            input_file = Path(video_path)
            output_file = input_file.with_name(f"Split_{input_file.stem}.mp4")

            cmd = [
                "ffmpeg",
                "-ss", timestamp,
                "-i", str(input_file),
                "-c", "copy",
                str(output_file)
            ]

            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            for line in process.stdout:
                self.log(line.strip())
            process.wait()
            if process.returncode != 0:
                raise subprocess.CalledProcessError(process.returncode, cmd)

            self.log(f"✅ Split saved as: {output_file}")
        except Exception as e:
            self.log(f"❌ Failed to split video: {e}")
            messagebox.showerror("Error", f"Failed to split video:{e}")

    def clip_clip(self):
        from tkinter import filedialog
        import subprocess
        import os
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

        self.log(f"⏳ Clipping from {start_time} to {end_time}...")

        try:
            input_file = Path(video_path)
            output_file = input_file.with_name(f"Clipped_{input_file.stem}.mp4")

            cmd = [
                "ffmpeg",
                "-ss", start_time,
                "-to", end_time,
                "-i", str(input_file),
                "-c", "copy",
                str(output_file)
            ]

            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            for line in process.stdout:
                self.log(line.strip())
            process.wait()
            if process.returncode != 0:
                raise subprocess.CalledProcessError(process.returncode, cmd)

            self.log(f"✅ Clipped segment saved as: {output_file}")
        except Exception as e:
            self.log(f"❌ Failed to clip video: {e}")
            messagebox.showerror("Error", f"Failed to clip video:\n{e}")

    def generate_gif(self):
        from tkinter import filedialog
        import subprocess
        import os
        from pathlib import Path

        def get_video_duration(path):
            try:
                result = subprocess.run(
                    ["ffprobe", "-v", "error", "-show_entries", "format=duration",
                    "-of", "default=noprint_wrappers=1:nokey=1", str(path)],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    check=True
                )
                return float(result.stdout.strip())
            except Exception as e:
                self.log(f"⚠️ Could not determine video duration: {e}")
                return None

        def parse_time_to_seconds(timestr):
            h, m, s = map(float, timestr.split(":"))
            return h * 3600 + m * 60 + s

        # Select .mp4 input file
        video_path = filedialog.askopenfilename(
            title="Select MP4 File to Convert to GIF", 
            filetypes=[("MP4 files", "*.mp4")]
        )
        if not video_path:
            return

        # Prompt for start time
        dialog_start = CustomInputDialog(
            self, title="GIF Start Time", prompt="Enter start time (e.g., 00:00:02):"
        )
        self.wait_window(dialog_start)
        start_time = dialog_start.result
        if not start_time:
            messagebox.showerror("Input Missing", "Start time is required.")
            return

        # Prompt for duration
        dialog_duration = CustomInputDialog(
            self, title="GIF Duration", prompt="Enter duration in seconds (e.g., 12):"
        )
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

        self.log(f"🎬 Generating GIF from {start_time} for {duration} seconds...")

        try:
            input_file = Path(video_path)
            output_gif = input_file.with_name(f"GIF_{input_file.stem}.gif")
            palette_path = input_file.with_name("palette.png")

            # 1st pass: generate palette
            cmd_palette = [
                "ffmpeg", "-ss", start_time, "-t", str(duration), "-i", str(input_file),
                "-vf", "fps=15,scale=1920:-1:force_original_aspect_ratio=decrease,palettegen",
                str(palette_path)
            ]
            self.log("🖌️ Generating palette...")
            process1 = subprocess.Popen(cmd_palette, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            for line in process1.stdout:
                self.log(line.strip())
            process1.wait()
            if process1.returncode != 0:
                raise subprocess.CalledProcessError(process1.returncode, cmd_palette)

            # 2nd pass: use palette to generate gif
            cmd_gif = [
                "ffmpeg", "-ss", start_time, "-t", str(duration), "-i", str(input_file), "-i", str(palette_path),
                "-filter_complex", "[0:v] fps=15,scale=1920:-1:force_original_aspect_ratio=decrease [x]; [x][1:v] paletteuse",
                "-loop", "0", str(output_gif)
            ]
            self.log("🎞️ Creating GIF...")
            process2 = subprocess.Popen(cmd_gif, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            for line in process2.stdout:
                self.log(line.strip())
            process2.wait()
            if process2.returncode != 0:
                raise subprocess.CalledProcessError(process2.returncode, cmd_gif)

            # Clean up palette
            if palette_path.exists():
                palette_path.unlink()

            self.log(f"✅ GIF saved as: {output_gif}")
        except Exception as e:
            self.log(f"❌ Failed to generate GIF: {e}")
            messagebox.showerror("Error", f"Failed to generate GIF:\n{e}")

    def return_to_home(self):
        self.destroy()
        Home_Page.open_home_page()

if __name__ == "__main__":
    VideoEditorGUI()
