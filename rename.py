import time
import os
import re
import shutil
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import ollama

# --- CONFIGURATION ---
WATCH_FOLDER = str(Path.home() / "Downloads")

# SPECIFIC MODEL: Uses the 2B parameter version (Fast/Light)
MODEL_NAME = "qwen3-vl:2b"

TARGET_PREFIXES = ["Screenshot", "Screen Shot"]
ALLOWED_EXTS = {".png", ".jpg", ".jpeg"}

class RenamerHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory: return
        self.process_file(event.src_path)

    def on_moved(self, event):
        if event.is_directory: return
        self.process_file(event.dest_path)

    def process_file(self, filepath):
        path = Path(filepath)
        # Check extension and prefix
        if path.suffix.lower() not in ALLOWED_EXTS: return
        if not any(path.name.startswith(prefix) for prefix in TARGET_PREFIXES): return

        # Wait 1 second for file write to finish
        time.sleep(1.0)
        print(f"Sending to Ollama ({MODEL_NAME}): {path.name}...")

        try:
            # Prompt designed for Qwen3
            prompt = (
                "Analyze this image. "
                "If it is a slide or document, extract the Main Title only. "
                "If it is a generic scene, describe it in 3 words. "
                "Output snake_case only. Do NOT output the file extension."
            )
            
            # Send to Ollama App
            response = ollama.chat(
                model=MODEL_NAME,
                messages=[{
                    'role': 'user',
                    'content': prompt,
                    'images': [filepath]
                }]
            )
            
            # Clean the result
            ai_suggestion = response['message']['content']
            final_name = self.clean_filename(ai_suggestion)
            
            self.rename_file(path, final_name)

        except Exception as e:
            print(f"Error processing {path.name}: {e}")

    def clean_filename(self, name):
        # Remove chatty prefixes if Qwen gets talkative
        name = re.sub(r'^(Here is|The title is|Output|Filename|Title):?\s*', '', name, flags=re.IGNORECASE)
        
        # Take first line only and remove markdown
        name = name.split('\n')[0].replace("`", "").strip()
        
        # Standard clean: Remove symbols, lowercase, snake_case
        name = re.sub(r'[^\w\s-]', '', name)
        name = re.sub(r'[-\s]+', '_', name).strip().lower()
        
        # Safety fallbacks
        if len(name) > 60: name = name[:60]
        if not name: name = "renamed_screenshot"
        
        return name

    def rename_file(self, original_path, new_name):
        new_filename = f"{new_name}{original_path.suffix}"
        new_path = original_path.parent / new_filename

        # Handle duplicates
        counter = 1
        while new_path.exists():
            new_filename = f"{new_name}_{counter}{original_path.suffix}"
            new_path = original_path.parent / new_filename
            counter += 1

        shutil.move(original_path, new_path)
        print(f"Renamed to: {new_filename}")

def main():
    print(f"Connecting to Ollama with model: {MODEL_NAME}...")
    
    # Verify connection
    try:
        ollama.list()
        print("Ollama connected successfully. Watching Downloads...")
    except Exception:
        print("ERROR: Ollama is not running! Please open the Ollama app in your Applications folder.")
        return

    observer = Observer()
    event_handler = RenamerHandler()
    observer.schedule(event_handler, WATCH_FOLDER, recursive=False)
    observer.start()
    
    try:
        while True: time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    main()
