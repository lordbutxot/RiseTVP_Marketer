import tkinter as tk
from tkinter import messagebox
from PIL import Image
import os

# ---------------- CONFIG ----------------
FOLDER = "images"
# Base crop area (will be adjusted)
BASE_CROP_AREA = (100, 200, 1200, 800)  # (left, top, right, bottom)

# ---------------- Processing ----------------

def process_folder():
    if not os.path.exists(FOLDER):
        messagebox.showerror("Error", f"Folder '{FOLDER}' not found")
        return

    files = os.listdir(FOLDER)
    processed_count = 0

    for filename in files:
        if not filename.lower().endswith(".png"):
            continue

        if not filename.startswith("FULL_"):
            continue

        input_path = os.path.join(FOLDER, filename)

        # Extract name (remove FULL_ and extension)
        name_part = filename.replace("FULL_", "", 1).rsplit(".", 1)[0]
        output_name = f"END_{name_part}.png"
        output_path = os.path.join(FOLDER, output_name)

        try:
            img = Image.open(input_path)

            # --- Adjust crop dynamically ---
            left, top, right, bottom = BASE_CROP_AREA
            width = right - left
            height = bottom - top

            # Move right by 1/3 of width
            shift = width / 3
            left += shift
            right += shift

            # Show more of the top by expanding upward
            extra_top = height * 0.05
            top -= extra_top

            cropped = img.crop((int(left), int(top), int(right), int(bottom)))
            cropped.save(output_path)
            processed_count += 1
        except Exception as e:
            print(f"Error processing {filename}: {e}")

    messagebox.showinfo("Done", f"Processed {processed_count} images")

# ---------------- GUI ----------------

def create_gui():
    root = tk.Tk()
    root.title("Batch Screenshot Crop Tool")

    tk.Label(root, text="Auto-process FULL_ images in /images folder", font=("Arial", 12)).pack(pady=10)

    tk.Button(root,
              text="Process All Images",
              width=30,
              command=process_folder
              ).pack(pady=20)

    tk.Label(root, text=f"Crop area: dynamic (shift right +33%, height +5%)", fg="gray").pack(pady=5)

    root.mainloop()

# ---------------- Main ----------------

if __name__ == "__main__":
    create_gui()
