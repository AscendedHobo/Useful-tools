import pyautogui
import tkinter as tk
from tkinter import ttk
import time
import threading

class MouseTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Mouse Tracker")
        self.running = False  # Toggle state
        self.counter = 1  # Initialize counter
        self.thread = None  # For tracking thread

        # Start/Stop Button
        self.start_stop_button = ttk.Button(root, text="Start", command=self.toggle_tracking)
        self.start_stop_button.pack(pady=20)

        # Status Label
        self.status_label = ttk.Label(root, text="Status: Stopped", font=("Arial", 12))
        self.status_label.pack(pady=10)

    def toggle_tracking(self):
        if not self.running:
            # Start tracking
            self.running = True
            self.start_stop_button.config(text="Stop")
            self.status_label.config(text="Status: Starting in 2 seconds...")
            
            # Wait 2 seconds before starting
            self.root.after(2000, self.start_tracking)
        else:
            # Stop tracking
            self.running = False
            self.start_stop_button.config(text="Start")
            self.status_label.config(text="Status: Stopped")

    def start_tracking(self):
        self.status_label.config(text="Status: Tracking mouse position and RGB...")
        # Run tracking in a separate thread
        self.thread = threading.Thread(target=self.track_mouse_position, daemon=True)
        self.thread.start()

    def track_mouse_position(self):
        while self.running:
            x, y = pyautogui.position()  # Get mouse position
            r, g, b = pyautogui.pixel(x, y)  # Get RGB value of the pixel at (x, y)
            print(f"Ref {self.counter}: Position -> X: {x}, Y: {y}, RGB -> ({r}, {g}, {b})")
            self.counter += 1
            time.sleep(3)  # Wait for 3 seconds

# Main execution
if __name__ == "__main__":
    root = tk.Tk()
    app = MouseTrackerApp(root)
    root.mainloop()
