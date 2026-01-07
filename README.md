## 1. AI-Powered Auto-Renamer (`rename.py`)

This script watches your **Downloads** folder for new screenshots. When one appears, it uses a local AI model (via Ollama) to analyze the image, extract titles from slides, or describe the scene, and renames the file into a clean `snake_case` format.

### Setup

1. **Install Dependencies:**
```bash
pip install watchdog ollama

```


2. **Download the Model:**
This script uses the `qwen3-vl:2b` vision model (lightweight and fast).
```bash
ollama pull qwen3-vl:2b

```


3. **Run:**
```bash
python3 rename.py

```
(I ended up using automator to make it into an app because I was able to make it run more smoothly and automatically like that)


---

## 2. Smart PPT Slide Capture (`screenshot-ppt-slide.sh`)

A Raycast script designed for those who study from PowerPoint in "Window" mode rather than Fullscreen. With one shortcut, it:

* Identifies the active PowerPoint window.
* **Auto-Crops:** Uses a custom Swift script to detect the actual slide boundaries (removing the PPT gray UI).
* **Smart Naming:** Tries to pull the title directly from the PPT metadata; if that fails, it performs local OCR on the slide header.
* **Course Logic:** Automatically extracts course codes (like "BIOL" or "ANAT") from the PowerPoint filename.

### Setup (macOS Only)

1. **Raycast:** Move the script to your Raycast script directory.
2. **Permissions:** You must grant **Raycast** (or your Terminal) the following permissions in *System Settings > Privacy & Security*:
* **Accessibility** (to read PPT window info).
* **Screen Recording** (to capture the slide).


3. **No Installs Needed:** This uses native AppleScript, Swift, and Vision frameworks.

---

## Important Notes

* **OS Compatibility:** These scripts were built for **macOS**. The Python script can be adapted for Windows by changing the folder path, but the `.sh` script is Mac-exclusive as it relies on AppleScript and Swift.
* **Ollama:** Ensure the Ollama app is running in the background for the Python script to work.
