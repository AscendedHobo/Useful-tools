import tkinter as tk
import pyautogui
import keyboard
import threading

# Disable pyautogui's built-in delay
pyautogui.PAUSE = 0.0

is_typing = False
stop_flag = False

def type_text():
    global is_typing, stop_flag
    is_typing = True
    stop_flag = False
    
    # First backspace to remove the triggering # character
    pyautogui.press('backspace')
    
    # Get all text from the text widget
    text = text_box.get("1.0", tk.END)
    
    # Filter out unwanted characters and type character by character
    for char in text:
        if stop_flag:
            break
        # Skip unwanted characters
        if char not in ['#', '*', '-']:
            pyautogui.write(char, interval=0.0)
        
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

# Start the global hotkey listener in a separate daemon thread
listener_thread = threading.Thread(target=hotkey_listener, daemon=True)
listener_thread.start()

root.mainloop()