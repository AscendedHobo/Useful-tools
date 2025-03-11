import pyautogui
import time
import os
from PIL import Image
import tkinter as tk
import keyboard


class RegionSelectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Region Selector")
        self.root.geometry("400x200")

        # Initialize variables
        self.top_left = None
        self.bottom_right = None

        # Create UI elements
        self.label_instruction = tk.Label(root, text="Hover and press Enter to set the region")
        self.label_instruction.pack(pady=10)

        self.label_result = tk.Label(root, text="Resulting Region: Not Set Yet", wraplength=350)
        self.label_result.pack(pady=10)

        self.button_start = tk.Button(root, text="Start Region Selection", command=self.start_selection)
        self.button_start.pack(pady=10)

        self.quit_button = tk.Button(root, text="Quit", command=root.quit)
        self.quit_button.pack(pady=5)

    def start_selection(self):
        self.label_instruction.config(text="Hover over the TOP-LEFT corner and press Enter...")
        self.root.update()

        keyboard.wait("enter")
        self.top_left = pyautogui.position()
        print(f"Top-left corner: {self.top_left}")

        self.label_instruction.config(text="Hover over the BOTTOM-RIGHT corner and press Enter...")
        self.root.update()

        keyboard.wait("enter")
        self.bottom_right = pyautogui.position()
        print(f"Bottom-right corner: {self.bottom_right}")

        # Calculate region
        x1, y1 = self.top_left
        x2, y2 = self.bottom_right
        region = (x1, y1, x2 - x1, y2 - y1)
        print(f"Region: {region}")

        # Update result on UI
        self.label_result.config(text=f"Resulting Region: {region}")

        # Save a screenshot of the region with the region coordinates in the filename
        save_folder = r"C:\Users\alanw\Desktop\temp region dump"
        if not os.path.exists(save_folder):
            os.makedirs(save_folder)  # Create folder if it doesn't exist

        filename = f"region_{x1}x{y1}_{x2-x1}x{y2-y1}.png"
        screenshot_path = os.path.join(save_folder, filename)
        pyautogui.screenshot(screenshot_path, region=region)
        print(f"Screenshot saved at: {screenshot_path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = RegionSelectorApp(root)
    root.mainloop()
