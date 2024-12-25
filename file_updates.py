import os
import subprocess
import sys

# List of libraries to update
libraries = [
    "moviepy",
    "PyPDF2",
    "pyttsx3",
    "python-pptx",
    "docx2pdf",
    "Pillow"
    "tkinter"
]

# Function to update libraries
def update_libraries():
    for lib in libraries:
        print(f"Updating {lib}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", lib])
        print(f"{lib} updated successfully.")

# Run the updater
if __name__ == "__main__":
    update_libraries()