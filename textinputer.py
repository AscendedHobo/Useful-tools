import tkinter as tk
import pyautogui
import keyboard
import threading
import pyperclip  # Add this import

# Disable pyautogui's built-in delay
pyautogui.PAUSE = 0.0

is_typing = False
stop_flag = False

def type_text():
    global is_typing, stop_flag
    is_typing = True
    stop_flag = False
    # Get all text from the text widget
    text = text_box.get("1.0", tk.END)
    
    # Method 1: Fast character typing
    if use_typing_var.get():
        for char in text:
            if stop_flag:
                break
            pyautogui.write(char, interval=0.0)
    # Method 2: Clipboard paste (much faster)
    else:
        pyperclip.copy(text)
        pyautogui.hotkey('ctrl', 'v')
        
    is_typing = False

def on_hash():
    global is_typing, stop_flag
    # If typing is in progress, set stop_flag to cancel it.
    if is_typing:
        stop_flag = True
    # Otherwise, start typing in a new thread.
    else:
        threading.Thread(target=type_text, daemon=True).start()

def hotkey_listener():
    # Register a global hotkey '#' that calls on_hash
    keyboard.add_hotkey('#', on_hash)
    keyboard.wait()

# Set up the main Tkinter window
root = tk.Tk()
root.title("Retype Program")

label = tk.Label(root, text="Paste your text into the box below.\nPress '#' to start typing.\nPress '#' during typing to cancel.")
label.pack(padx=10, pady=5)

# Use a multi-line Text widget for larger text input
text_box = tk.Text(root, width=60, height=20)
text_box.pack(padx=10, pady=5)

# In the GUI setup section, add a checkbox for choosing the input method
use_typing_var = tk.BooleanVar(value=False)
typing_checkbox = tk.Checkbutton(root, text="Use character-by-character typing", variable=use_typing_var)
typing_checkbox.pack(padx=10, pady=5)

# Start the global hotkey listener in a separate daemon thread
listener_thread = threading.Thread(target=hotkey_listener, daemon=True)
listener_thread.start()

root.mainloop()
