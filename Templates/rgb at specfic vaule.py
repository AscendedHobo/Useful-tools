import pyautogui
from tkinter import Tk, Button, Label

def check_rgb():
    # Position to check
    x, y = 1182, 1219
    # Get RGB value at the position
    rgb = pyautogui.screenshot().getpixel((x, y))
    # Update the label with the RGB value
    rgb_label.config(text=f"RGB at ({x}, {y}): {rgb}")
    # Print the RGB value to the terminal
    print(f"RGB at ({x}, {y}): {rgb}")

# Create the tkinter GUI
root = Tk()
root.title("RGB Checker")

# Add a button to check RGB
check_button = Button(root, text="Check RGB", command=check_rgb)
check_button.pack(pady=10)

# Label to display the RGB value
rgb_label = Label(root, text="Click 'Check RGB' to get the value", font=("Arial", 12))
rgb_label.pack(pady=10)

# Run the tkinter event loop
root.mainloop()
