import tkinter as tk
from tkinter import ttk, simpledialog, filedialog, colorchooser, scrolledtext
import pyautogui
import time
import threading
import json
import os
import shutil
import uuid # For unique listener IDs
from PIL import Image, ImageTk, ImageGrab

# Attempt to import the keyboard library for global hotkeys
try:
    import keyboard
except ImportError:
    simpledialog.messagebox.showerror("Missing Library",
                                      "The 'keyboard' library is not installed. "
                                      "Please install it by running: pip install keyboard\n\n"
                                      "Global hotkey for capturing icons will not work.")
    keyboard = None


# --- Global Variables & Constants ---
DEFAULT_PROFILE_NAME = "UntitledProfile"
CAPTURE_HOTKEY_STORAGE = "capture_hotkey.json" # To store the preferred hotkey

# --- Helper Functions ---
def get_screen_center_for_window(window_width, window_height, root):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    return x, y

# --- Main Application Class ---
class WoWAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WoW Icon Listener Tool")
        x, y = get_screen_center_for_window(700, 500, root)
        self.root.geometry(f"700x500+{x}+{y}")

        self.listeners = []  # List of listener dicts
        self.current_project_path = None # Path to the folder where the .json and images/ are
        self.current_profile_name = DEFAULT_PROFILE_NAME
        self.profile_modified = False

        self.capture_hotkey = "f12" # Default capture hotkey
        self.load_capture_hotkey_preference()

        self.watcher_thread = None
        self.watcher_active = threading.Event() # False by default
        self.scan_interval = 0.2 # Seconds between full screen scans if no icon found

        self.drag_select_window = None
        self.drag_start_x = None
        self.drag_start_y = None
        self.drag_rect_id = None

        self.container = tk.Frame(root)
        self.container.pack(fill="both", expand=True)

        self.frames = {}
        # Only one main frame for now
        main_frame_instance = ListenerManagerFrame(parent=self.container, controller=self)
        self.frames[ListenerManagerFrame.__name__] = main_frame_instance
        main_frame_instance.grid(row=0, column=0, sticky="nsew")

        self.show_frame(ListenerManagerFrame.__name__)
        self.setup_global_capture_hotkey()
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        pyautogui.FAILSAFE = True


    def mark_profile_modified(self, modified=True):
        self.profile_modified = modified
        title = "WoW Icon Listener Tool - " + self.current_profile_name
        if modified:
            self.root.title(title + "*")
        else:
            self.root.title(title)

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()
        if hasattr(frame, 'refresh_content'):
            frame.refresh_content()
        self.mark_profile_modified(self.profile_modified) # Update title correctly

    def setup_global_capture_hotkey(self):
        if not keyboard:
            print("Keyboard library not available. Global hotkey disabled.")
            return
        try:
            # Remove previous hotkey if any
            try:
                keyboard.unhook_all_hotkeys()
            except AttributeError:
                # Fallback for older versions
                try:
                    keyboard.remove_all_hotkeys()
                except AttributeError:
                    # Last resort: just continue without removing
                    print("Warning: Could not remove existing hotkeys using standard methods")
            
            keyboard.add_hotkey(self.capture_hotkey, self.initiate_capture_from_hotkey_threadsafe, suppress=True)
            print(f"Global capture hotkey '{self.capture_hotkey}' set.")
            if self.frames[ListenerManagerFrame.__name__]:
                self.frames[ListenerManagerFrame.__name__].update_hotkey_display()
        except Exception as e:
            simpledialog.messagebox.showerror("Hotkey Error", f"Could not set hotkey '{self.capture_hotkey}': {e}\nTry running as administrator if it's a permissions issue, or choose a different hotkey.", parent=self.root)
            print(f"Error setting hotkey: {e}")

    def initiate_capture_from_hotkey_threadsafe(self):
        # This is called from the keyboard library's thread
        if self.drag_select_window and self.drag_select_window.winfo_exists():
             print("Capture already in progress.")
             return
        self.root.after(0, self.start_icon_capture_drag_mode) # Schedule GUI update in main thread

    def start_icon_capture_drag_mode(self):
        if self.drag_select_window and self.drag_select_window.winfo_exists():
            return # Already active
        
        self.root.withdraw() # Hide main window briefly
        time.sleep(0.2) # Give screen time to settle if WoW was fullscreen

        self.drag_select_window = tk.Toplevel(self.root)
        self.drag_select_window.attributes('-fullscreen', True)
        self.drag_select_window.attributes('-alpha', 0.3)
        self.drag_select_window.attributes('-topmost', True)
        drag_canvas = tk.Canvas(self.drag_select_window, bg="gray", cursor="crosshair", highlightthickness=0)
        drag_canvas.pack(fill="both", expand=True)
        
        status_label_text = "Click & drag to select icon. Release to confirm. ESC to cancel."
        try: # Try to get screen dimensions for better label placement
            screen_w = self.root.winfo_screenwidth()
            screen_h = self.root.winfo_screenheight()
            tk.Label(drag_canvas, text=status_label_text, bg="lightyellow", fg="black", font=("Arial", 14)).place(x=screen_w/2, y=30, anchor="n")
        except: # Fallback if screen info not available yet
            tk.Label(drag_canvas, text=status_label_text, bg="lightyellow", fg="black", font=("Arial", 14)).place(x=10,y=10)


        def on_b1_press(event):
            self.drag_start_x = event.x
            self.drag_start_y = event.y
            self.drag_rect_id = drag_canvas.create_rectangle(
                self.drag_start_x, self.drag_start_y,
                self.drag_start_x, self.drag_start_y,
                outline='red', width=2
            )

        def on_b1_motion(event):
            if self.drag_rect_id:
                drag_canvas.coords(self.drag_rect_id, self.drag_start_x, self.drag_start_y, event.x, event.y)

        def on_b1_release(event):
            if self.drag_start_x is None: return
            x1, y1 = min(self.drag_start_x, event.x), min(self.drag_start_y, event.y)
            x2, y2 = max(self.drag_start_x, event.x), max(self.drag_start_y, event.y)
            
            self.drag_select_window.destroy()
            self.drag_select_window = None
            self.drag_start_x, self.drag_start_y, self.drag_rect_id = None, None, None

            if abs(x1 - x2) < 10 or abs(y1 - y2) < 10: # Minimum size for an icon
                simpledialog.messagebox.showinfo("Info", "Selection too small.", parent=self.root)
                self.root.deiconify()
                return

            capture_coords = (int(x1), int(y1), int(x2 - x1), int(y2 - y1))
            
            # Capture the image (main window is already hidden)
            try:
                # No need to hide/deiconify root here as it was done before creating drag_select_window
                # However, ensure drag_select_window is truly gone before screenshot
                time.sleep(0.1) # Short delay to ensure transparent window is gone
                img = pyautogui.screenshot(region=capture_coords)
            except Exception as e:
                simpledialog.messagebox.showerror("Capture Error", f"Could not capture screenshot: {e}", parent=self.root)
                self.root.deiconify()
                return
            
            self.root.deiconify() # Show main window again to interact with dialogs
            self.root.focus_force()
            self.process_new_icon_capture(img, capture_coords)

        def on_escape_drag(event=None):
            if self.drag_select_window:
                self.drag_select_window.destroy()
                self.drag_select_window = None
            self.drag_start_x, self.drag_start_y, self.drag_rect_id = None, None, None
            print("Drag selection cancelled.")
            self.root.deiconify() # Ensure main window is shown

        drag_canvas.bind("<ButtonPress-1>", on_b1_press)
        drag_canvas.bind("<B1-Motion>", on_b1_motion)
        drag_canvas.bind("<ButtonRelease-1>", on_b1_release)
        self.drag_select_window.bind("<Escape>", on_escape_drag)
        self.drag_select_window.focus_force()

    def process_new_icon_capture(self, image_obj, capture_coords):
        listener_name = simpledialog.askstring("Listener Name", "Enter a name for this icon/listener:", parent=self.root)
        if not listener_name: return

        keybind_raw = simpledialog.askstring("Keybind",
                                             "Enter keybind to press when icon is detected.\n"
                                             "(e.g., '1', 'f', 'ctrl+s', 'alt+shift+z')",
                                             parent=self.root)
        if not keybind_raw: return

        confidence = simpledialog.askfloat("Confidence", "Matching confidence (0.1 to 1.0):",
                                           parent=self.root, minvalue=0.1, maxvalue=1.0, initialvalue=0.8)
        if confidence is None: return # User cancelled

        post_press_delay = simpledialog.askfloat("Post-Press Delay", "Delay (seconds) between key presses if icon persists:",
                                           parent=self.root, minvalue=0.01, maxvalue=2.0, initialvalue=0.1)
        if post_press_delay is None: post_press_delay = 0.1

        max_seq_presses = simpledialog.askinteger("Max Sequential Presses", "Max times to press key before re-evaluating screen (0 for unlimited until icon gone):",
                                           parent=self.root, minvalue=0, initialvalue=3)
        if max_seq_presses is None: max_seq_presses = 3


        # Generate image filename
        base_img_filename = f"{listener_name.replace(' ', '_').replace('.', '_')}_{uuid.uuid4().hex[:8]}.png"
        
        # Determine image save path
        abs_image_path = ""
        relative_image_path_for_json = ""

        if self.current_project_path:
            images_dir = os.path.join(self.current_project_path, "images")
            os.makedirs(images_dir, exist_ok=True)
            abs_image_path = os.path.join(images_dir, base_img_filename)
            relative_image_path_for_json = os.path.join("images", base_img_filename)
        else:
            # No project loaded/saved, save to a temporary spot or current dir
            # For simplicity, let's enforce saving the profile first or handle it more robustly.
            # For now, save to current working directory and warn
            images_dir = os.path.join(os.getcwd(), "images_temp") # Or just os.getcwd()
            os.makedirs(images_dir, exist_ok=True) # Ensure it exists
            abs_image_path = os.path.join(images_dir, base_img_filename)
            relative_image_path_for_json = abs_image_path # Store absolute path if no project
            simpledialog.messagebox.showwarning("No Profile Loaded",
                                                f"Image saved to: {abs_image_path}\n"
                                                "This path will be relative to the profile file once you 'Save Profile As...'.",
                                                parent=self.root)
        try:
            image_obj.save(abs_image_path)
        except Exception as e:
            simpledialog.messagebox.showerror("Image Save Error", f"Could not save image: {e}", parent=self.root)
            return

        new_listener = {
            "id": uuid.uuid4().hex,
            "name": listener_name,
            "image_filename": base_img_filename, # Store filename, full path constructed at runtime
            "keybind_raw": keybind_raw.lower(),
            "active": True,
            "confidence": confidence,
            "capture_coords": capture_coords,
            "post_press_delay": post_press_delay,
            "max_sequential_presses": max_seq_presses,
            "_image_abs_path_temp": abs_image_path if not self.current_project_path else None # Store temp abs path
        }
        self.listeners.append(new_listener)
        self.frames[ListenerManagerFrame.__name__].refresh_listeners_list()
        self.mark_profile_modified()

    def _check_unsaved_changes(self):
        if self.profile_modified:
            response = simpledialog.messagebox.askyesnocancel("Unsaved Changes",
                                                 f"Profile '{self.current_profile_name}' has unsaved changes. Save now?",
                                                 parent=self.root)
            if response is True:  # Yes
                return self.save_profile()
            elif response is False:  # No
                return True
            else:  # Cancel
                return False
        return True # No unsaved changes or user chose not to save

    def new_profile(self):
        if not self._check_unsaved_changes():
            return
        self.listeners = []
        self.current_project_path = None
        self.current_profile_name = DEFAULT_PROFILE_NAME
        self.frames[ListenerManagerFrame.__name__].refresh_listeners_list()
        self.mark_profile_modified(False)
        print("New profile created.")

    def save_profile(self):
        if not self.current_project_path:
            return self.save_profile_as()
        else:
            profile_filename = os.path.join(self.current_project_path, f"{self.current_profile_name}.json")
            project_images_dir = os.path.join(self.current_project_path, "images")
            os.makedirs(project_images_dir, exist_ok=True)

            listeners_to_save = []
            for listener in self.listeners:
                listener_copy = listener.copy()
                
                # Handle image path: ensure it's relative and file is in project images dir
                current_image_path_source = listener_copy.pop("_image_abs_path_temp", None) # Get and remove temp path
                if not current_image_path_source: # If no temp path, construct from filename (already in project)
                    current_image_path_source = os.path.join(project_images_dir, listener_copy["image_filename"])


                if not os.path.isabs(current_image_path_source): # Should not happen if _image_abs_path_temp was used
                     # This might be a relative path from a previous load, ensure it's valid
                     potential_abs_path = os.path.join(self.current_project_path, current_image_path_source)
                     if os.path.exists(potential_abs_path):
                         current_image_path_source = potential_abs_path
                     else: # Fallback if it was already just a filename from a loaded project
                         current_image_path_source = os.path.join(project_images_dir, listener_copy["image_filename"])


                target_abs_path_in_project = os.path.join(project_images_dir, listener_copy["image_filename"])

                if os.path.exists(current_image_path_source):
                    if os.path.normpath(current_image_path_source) != os.path.normpath(target_abs_path_in_project):
                        try:
                            shutil.copy2(current_image_path_source, target_abs_path_in_project)
                            print(f"Copied image for '{listener_copy['name']}' to project images.")
                            # If original was in images_temp, we might want to clean it up later or on exit.
                        except Exception as e:
                            simpledialog.messagebox.showerror("Save Error", f"Could not copy image '{listener_copy['image_filename']}' to project: {e}", parent=self.root)
                            # Continue saving JSON but image might be missing
                elif not os.path.exists(target_abs_path_in_project):
                     print(f"Warning: Image file for '{listener_copy['name']}' ({listener_copy['image_filename']}) not found at source or target. JSON will reference it but file may be missing.")

                # JSON should always store relative path to images dir
                listener_copy["image_filename"] = listener_copy["image_filename"] # Already just filename

                listeners_to_save.append(listener_copy)
            
            data_to_save = {
                "profile_name": self.current_profile_name,
                "listeners": listeners_to_save,
                "capture_hotkey": self.capture_hotkey # Save hotkey with profile
            }

            try:
                with open(profile_filename, 'w') as f:
                    json.dump(data_to_save, f, indent=4)
                simpledialog.messagebox.showinfo("Save Profile", f"Profile '{self.current_profile_name}' saved.", parent=self.root)
                self.mark_profile_modified(False)
                # Clean up temp images_temp if we created it
                temp_images_dir = os.path.join(os.getcwd(), "images_temp")
                if os.path.exists(temp_images_dir) and not os.listdir(temp_images_dir): # Check if empty
                    try:
                        os.rmdir(temp_images_dir)
                    except OSError: # Not empty or other issue
                        pass 
                elif os.path.exists(temp_images_dir) and any(l.get("_image_abs_path_temp") and temp_images_dir in l.get("_image_abs_path_temp") for l in self.listeners) :
                    # if any listener still points to temp, don't delete, but this should not happen after copy
                    pass


                return True
            except Exception as e:
                simpledialog.messagebox.showerror("Save Error", f"Could not save profile: {e}", parent=self.root)
                return False

    def save_profile_as(self):
        project_dir = filedialog.askdirectory(title="Select Folder to Save Profile", parent=self.root)
        if not project_dir:
            return False

        # The "profile name" will be the name of the JSON file and part of the folder structure
        # For simplicity, let's say the chosen directory IS the project path.
        # The JSON file will be named after the profile.
        
        default_name = self.current_profile_name if self.current_profile_name != DEFAULT_PROFILE_NAME else "MyWoWProfile"
        profile_file_name_base = simpledialog.askstring("Profile Name",
                                                 "Enter a name for this profile (this will be the .json file name):",
                                                 initialvalue=default_name, parent=self.root)
        if not profile_file_name_base:
            return False

        self.current_project_path = project_dir # The selected folder is the root of the project
        self.current_profile_name = profile_file_name_base
        
        return self.save_profile() # Now save_profile will use the new path and name

    def load_profile(self):
        if not self._check_unsaved_changes():
            return

        filepath = filedialog.askopenfilename(
            title="Load Profile File",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            parent=self.root
        )
        if not filepath:
            return

        try:
            with open(filepath, 'r') as f:
                loaded_data = json.load(f)

            self.current_project_path = os.path.dirname(filepath) # Directory containing the .json
            self.current_profile_name = loaded_data.get("profile_name", os.path.splitext(os.path.basename(filepath))[0])
            
            self.listeners = []
            for listener_data in loaded_data.get("listeners", []):
                # Image path construction: image_filename is relative to "images" subfolder in project_path
                img_filename = listener_data.get("image_filename")
                if img_filename:
                    abs_image_path = os.path.join(self.current_project_path, "images", img_filename)
                    if not os.path.exists(abs_image_path):
                        print(f"Warning: Image file '{img_filename}' for listener '{listener_data.get('name')}' not found at '{abs_image_path}'.")
                        # Listener will be added, but image detection might fail.
                self.listeners.append(listener_data)
            
            loaded_hotkey = loaded_data.get("capture_hotkey", self.capture_hotkey)
            if self.capture_hotkey != loaded_hotkey:
                 self.capture_hotkey = loaded_hotkey
                 self.save_capture_hotkey_preference() # Save it globally too
                 self.setup_global_capture_hotkey() # Re-apply the loaded hotkey


            self.frames[ListenerManagerFrame.__name__].refresh_listeners_list()
            self.mark_profile_modified(False)
            simpledialog.messagebox.showinfo("Load Profile", f"Profile '{self.current_profile_name}' loaded.", parent=self.root)

        except Exception as e:
            simpledialog.messagebox.showerror("Load Error", f"Could not load profile: {e}", parent=self.root)
            self.new_profile() # Reset to a clean state

    def save_capture_hotkey_preference(self):
        try:
            with open(CAPTURE_HOTKEY_STORAGE, 'w') as f:
                json.dump({"capture_hotkey": self.capture_hotkey}, f)
        except Exception as e:
            print(f"Could not save hotkey preference: {e}")

    def load_capture_hotkey_preference(self):
        try:
            if os.path.exists(CAPTURE_HOTKEY_STORAGE):
                with open(CAPTURE_HOTKEY_STORAGE, 'r') as f:
                    data = json.load(f)
                    self.capture_hotkey = data.get("capture_hotkey", self.capture_hotkey)
        except Exception as e:
            print(f"Could not load hotkey preference: {e}")

    def set_new_capture_hotkey(self):
        if self.watcher_active.is_set():
            simpledialog.messagebox.showwarning("Warning", "Stop watching before changing the hotkey.", parent=self.root)
            return

        new_hotkey = simpledialog.askstring("Set Capture Hotkey",
                                            "Enter new global hotkey for capturing icons (e.g., 'f12', 'ctrl+shift+c'):",
                                            initialvalue=self.capture_hotkey, parent=self.root)
        if new_hotkey and new_hotkey != self.capture_hotkey:
            try:
                if keyboard: 
                    # Try different methods to remove hotkeys as the API might have changed
                    try:
                        keyboard.unhook_all_hotkeys()
                    except AttributeError:
                        # Fallback for older versions
                        try:
                            keyboard.remove_all_hotkeys()
                        except AttributeError:
                            # Last resort: try to clear all hotkeys individually
                            print("Warning: Could not remove hotkeys using standard methods")
                self.capture_hotkey = new_hotkey.lower()
                self.save_capture_hotkey_preference()
                self.setup_global_capture_hotkey() # This will re-register
            except Exception as e:
                simpledialog.messagebox.showerror("Hotkey Error", f"Error setting hotkey: {e}", parent=self.root)
                print(f"Error in set_new_capture_hotkey: {e}")


    def start_watching(self):
        if not self.listeners:
            simpledialog.messagebox.showinfo("Start Watching", "No listeners configured.", parent=self.root)
            return
        if self.watcher_thread and self.watcher_thread.is_alive():
            simpledialog.messagebox.showinfo("Start Watching", "Watcher is already running.", parent=self.root)
            return

        self.watcher_active.set() # Signal thread to run
        self.watcher_thread = threading.Thread(target=self._watch_loop, daemon=True)
        self.watcher_thread.start()
        self.frames[ListenerManagerFrame.__name__].update_watch_button_state(True)
        print("--- Started Watching ---")

    def stop_watching(self):
        self.watcher_active.clear() # Signal thread to stop
        if self.watcher_thread and self.watcher_thread.is_alive():
            print("--- Stopping Watcher (waiting for thread to join)... ---")
            self.watcher_thread.join(timeout=2) # Wait for thread to finish
            if self.watcher_thread.is_alive():
                print("--- Watcher thread did not join in time. ---")
        self.watcher_thread = None
        self.frames[ListenerManagerFrame.__name__].update_watch_button_state(False)
        print("--- Stopped Watching ---")

    def _watch_loop(self):
        self.root.iconify() # Minimize main window while watching
        time.sleep(0.5) # Give it time to minimize

        while self.watcher_active.is_set():
            processed_one_this_cycle = False
            # Get a snapshot of listeners by priority (IDs)
            # The actual listener data is fetched by ID to get the most current 'active' state
            # Order is determined by self.listeners which is manipulated by UI
            
            # Create a temporary list of active listeners in their current UI order
            active_listeners_in_order = [l for l in self.listeners if l.get('active', False)]

            for listener in active_listeners_in_order:
                if not self.watcher_active.is_set(): break # Check event before each potentially long operation

                try:
                    if not self.current_project_path and listener.get("_image_abs_path_temp"):
                         image_to_check = listener["_image_abs_path_temp"]
                    elif self.current_project_path and listener.get("image_filename"):
                         image_to_check = os.path.join(self.current_project_path, "images", listener["image_filename"])
                    else:
                        print(f"Skipping listener {listener['name']}: Image path information missing.")
                        continue
                    
                    if not os.path.exists(image_to_check):
                        # This check can be frequent, consider logging less verbosely or only once
                        # print(f"Image not found for {listener['name']}: {image_to_check}")
                        continue

                    location = pyautogui.locateOnScreen(
                        image_to_check,
                        confidence=listener.get('confidence', 0.8)
                    )

                    if location:
                        print(f"Detected: {listener['name']}. Pressing: {listener['keybind_raw']}")
                        keys_to_press = listener['keybind_raw'].split('+')
                        
                        press_count = 0
                        max_presses = listener.get('max_sequential_presses', 3)
                        delay = listener.get('post_press_delay', 0.1)

                        # Inner loop: Press keybind until icon disappears or max_presses reached
                        while self.watcher_active.is_set() and pyautogui.locateOnScreen(image_to_check, confidence=listener.get('confidence', 0.8)):
                            if len(keys_to_press) == 1:
                                pyautogui.press(keys_to_press[0])
                            else:
                                pyautogui.hotkey(*keys_to_press)
                            
                            press_count += 1
                            time.sleep(delay)

                            if max_presses > 0 and press_count >= max_presses:
                                print(f"Max ({max_presses}) sequential presses for {listener['name']}. Re-evaluating.")
                                break
                        
                        print(f"Icon {listener['name']} action complete (pressed {press_count} times).")
                        processed_one_this_cycle = True
                        # IMPORTANT: Restart scan from highest priority after an action
                        # This is achieved by breaking this inner loop and the outer for-loop will restart
                        break 
                
                except pyautogui.FailSafeException:
                    print("!!! FAILSAFE TRIGGERED (mouse to top-left) !!!")
                    self.watcher_active.clear() # Stop the loop
                    # Ensure deiconify runs on main thread
                    self.root.after(0, self.root.deiconify)
                    self.root.after(0, lambda: self.frames[ListenerManagerFrame.__name__].update_watch_button_state(False))
                    return # Exit thread
                except Exception as e:
                    import traceback
                    print(f"Error during watch loop for listener {listener.get('name', 'Unknown')}: {e}")
                    traceback.print_exc()
                    # Optionally disable faulty listener or add a cooldown
                    # For now, just continue to the next listener or next cycle

            if not self.watcher_active.is_set(): break # Check again before sleep

            if not processed_one_this_cycle:
                time.sleep(self.scan_interval) # Sleep only if no icon was processed in this full pass
        
        # Loop finished
        if self.root.state() == 'iconic': # If still minimized
            self.root.after(0, self.root.deiconify) # Ensure deiconify runs on main thread

    def on_closing(self):
        if self.watcher_active.is_set():
            self.stop_watching()
        if not self._check_unsaved_changes():
            return # User cancelled closing
        if keyboard:
            keyboard.remove_all_hotkeys()
        self.root.destroy()


# --- UI Frame Class ---
class ListenerManagerFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.configure(bg="#E0E0E0")

        # --- Menu Bar ---
        menubar = tk.Menu(self.controller.root)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="New Profile", command=self.controller.new_profile)
        filemenu.add_command(label="Load Profile...", command=self.controller.load_profile)
        filemenu.add_command(label="Save Profile", command=self.controller.save_profile)
        filemenu.add_command(label="Save Profile As...", command=self.controller.save_profile_as)
        filemenu.add_separator()
        filemenu.add_command(label="Set Capture Hotkey...", command=self.controller.set_new_capture_hotkey)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.controller.on_closing)
        menubar.add_cascade(label="File", menu=filemenu)
        self.controller.root.config(menu=menubar)

        # --- Top Controls ---
        top_controls_frame = tk.Frame(self, bg=self["bg"])
        top_controls_frame.pack(pady=10, padx=10, fill="x")

        self.watch_button = tk.Button(top_controls_frame, text="Start Watching", font=("Arial", 12, "bold"),
                                      bg="#A5D6A7", command=self.toggle_watch_state, width=15)
        self.watch_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.hotkey_label = tk.Label(top_controls_frame, text=f"Capture Hotkey: {self.controller.capture_hotkey.upper()}", bg=self["bg"])
        self.hotkey_label.pack(side=tk.LEFT, padx=10)
        
        tk.Button(top_controls_frame, text="Info/Help", command=self.show_help).pack(side=tk.RIGHT)


        # --- Listeners List (Treeview) ---
        list_frame = tk.Frame(self, bg=self["bg"])
        list_frame.pack(pady=5, padx=10, fill="both", expand=True)

        cols = ("#", "Name", "Keybind", "Active", "Confidence")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", selectmode="browse")
        
        self.tree.heading("#", text="#", anchor="w")
        self.tree.column("#", width=30, stretch=False, anchor="center")
        self.tree.heading("Name", text="Name", anchor="w")
        self.tree.column("Name", width=200, stretch=True)
        self.tree.heading("Keybind", text="Keybind", anchor="w")
        self.tree.column("Keybind", width=100, stretch=False, anchor="center")
        self.tree.heading("Active", text="Active", anchor="w")
        self.tree.column("Active", width=60, stretch=False, anchor="center")
        self.tree.heading("Confidence", text="Confidence", anchor="w")
        self.tree.column("Confidence", width=80, stretch=False, anchor="center")

        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.bind("<Double-1>", self.on_listener_double_click)

        # --- Bottom Controls ---
        bottom_controls_frame = tk.Frame(self, bg=self["bg"])
        bottom_controls_frame.pack(pady=10, padx=10, fill="x")

        tk.Button(bottom_controls_frame, text="Remove Selected", command=self.remove_selected_listener).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_controls_frame, text="Edit Selected", command=self.edit_selected_listener).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_controls_frame, text="Move Up", command=lambda: self.move_listener(-1)).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_controls_frame, text="Move Down", command=lambda: self.move_listener(1)).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_controls_frame, text="Toggle Active", command=self.toggle_selected_listener_active).pack(side=tk.LEFT, padx=5)

        self.refresh_listeners_list()

    def update_hotkey_display(self):
        self.hotkey_label.config(text=f"Capture Hotkey: {self.controller.capture_hotkey.upper()}")

    def toggle_watch_state(self):
        if self.controller.watcher_active.is_set():
            self.controller.stop_watching()
        else:
            self.controller.start_watching()
    
    def update_watch_button_state(self, is_watching):
        if is_watching:
            self.watch_button.config(text="Stop Watching", bg="#FFBBAA") # Reddish for stop
        else:
            self.watch_button.config(text="Start Watching", bg="#A5D6A7") # Greenish for start

    def refresh_listeners_list(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for idx, listener in enumerate(self.controller.listeners):
            active_str = "Yes" if listener.get("active", False) else "No"
            self.tree.insert("", "end", iid=listener["id"], values=(
                idx + 1,
                listener["name"],
                listener["keybind_raw"],
                active_str,
                f"{listener.get('confidence', 0.8):.2f}"
            ))

    def get_selected_listener_id(self):
        selected_item = self.tree.focus() # Gets the iid of the selected item
        if not selected_item:
            simpledialog.messagebox.showinfo("Action", "No listener selected.", parent=self.controller.root)
            return None
        return selected_item # This is the listener's ID

    def find_listener_index_by_id(self, listener_id):
        for i, listener in enumerate(self.controller.listeners):
            if listener["id"] == listener_id:
                return i
        return -1

    def remove_selected_listener(self):
        listener_id = self.get_selected_listener_id()
        if not listener_id: return

        idx = self.find_listener_index_by_id(listener_id)
        if idx != -1:
            # Optionally, ask for confirmation before deleting image file
            listener_to_remove = self.controller.listeners[idx]
            if simpledialog.messagebox.askyesno("Confirm Delete", f"Delete listener '{listener_to_remove['name']}'?\nAssociated image file will NOT be deleted from disk automatically by this action, but will not be used.", parent=self.controller.root):
                # Image file deletion can be tricky (e.g. if shared). Let's not delete it for now.
                # User can clean up images folder manually if needed.
                self.controller.listeners.pop(idx)
                self.refresh_listeners_list()
                self.controller.mark_profile_modified()

    def edit_selected_listener(self):
        listener_id = self.get_selected_listener_id()
        if not listener_id: return
        
        idx = self.find_listener_index_by_id(listener_id)
        if idx == -1: return

        listener = self.controller.listeners[idx]
        
        # Create a dialog for editing - for now, just key aspects
        new_name = simpledialog.askstring("Edit Name", "Listener Name:", initialvalue=listener["name"], parent=self.controller.root)
        if new_name is None: return # User cancelled name edit
        
        new_keybind = simpledialog.askstring("Edit Keybind", "Keybind:", initialvalue=listener["keybind_raw"], parent=self.controller.root)
        if new_keybind is None: return

        new_confidence = simpledialog.askfloat("Edit Confidence", "Confidence (0.1-1.0):", initialvalue=listener["confidence"], minvalue=0.1, maxvalue=1.0, parent=self.controller.root)
        if new_confidence is None: return
        
        new_post_press_delay = simpledialog.askfloat("Edit Post-Press Delay", "Delay (s):", initialvalue=listener.get("post_press_delay", 0.1), minvalue=0.01, maxvalue=2.0, parent=self.controller.root)
        if new_post_press_delay is None: new_post_press_delay = listener.get("post_press_delay", 0.1)

        new_max_seq_presses = simpledialog.askinteger("Edit Max Sequential Presses", "Max presses:", initialvalue=listener.get("max_sequential_presses",3), minvalue=0, parent=self.controller.root)
        if new_max_seq_presses is None: new_max_seq_presses = listener.get("max_sequential_presses",3)


        # Check if anything changed
        changed = (listener["name"] != new_name or
                   listener["keybind_raw"] != new_keybind.lower() or
                   listener["confidence"] != new_confidence or
                   listener.get("post_press_delay", 0.1) != new_post_press_delay or
                   listener.get("max_sequential_presses", 3) != new_max_seq_presses)

        if changed:
            listener["name"] = new_name
            listener["keybind_raw"] = new_keybind.lower()
            listener["confidence"] = new_confidence
            listener["post_press_delay"] = new_post_press_delay
            listener["max_sequential_presses"] = new_max_seq_presses
            self.controller.mark_profile_modified()
            self.refresh_listeners_list()


    def move_listener(self, direction): # -1 for up, 1 for down
        listener_id = self.get_selected_listener_id()
        if not listener_id: return

        idx = self.find_listener_index_by_id(listener_id)
        if idx == -1: return

        new_idx = idx + direction
        if 0 <= new_idx < len(self.controller.listeners):
            self.controller.listeners.insert(new_idx, self.controller.listeners.pop(idx))
            self.refresh_listeners_list()
            self.controller.mark_profile_modified()
            # Re-select the moved item
            if self.tree.exists(listener_id):
                self.tree.focus(listener_id)
                self.tree.selection_set(listener_id)


    def toggle_selected_listener_active(self):
        listener_id = self.get_selected_listener_id()
        if not listener_id: return

        idx = self.find_listener_index_by_id(listener_id)
        if idx != -1:
            listener = self.controller.listeners[idx]
            listener["active"] = not listener.get("active", False)
            self.refresh_listeners_list()
            self.controller.mark_profile_modified()
            # Re-select
            if self.tree.exists(listener_id):
                self.tree.focus(listener_id)
                self.tree.selection_set(listener_id)

    def on_listener_double_click(self, event):
        item_iid = self.tree.identify_row(event.y) # Get iid of clicked row
        if item_iid:
            self.tree.focus(item_iid) # Focus the clicked item
            self.tree.selection_set(item_iid) # Select it
            self.toggle_selected_listener_active() # Then toggle its state
            # Or call self.edit_selected_listener() if preferred

    def refresh_content(self): # Called when frame is shown
        self.refresh_listeners_list()
        self.update_watch_button_state(self.controller.watcher_active.is_set())
        self.update_hotkey_display()

    def show_help(self):
        help_text = """WoW Icon Listener Tool - Help

File Menu:
- New/Load/Save Profile: Manage your listener configurations. Profiles are saved as .json files, with captured icons in an 'images' subfolder.
- Set Capture Hotkey: Change the global hotkey used to initiate icon capture (default F12). Requires an application restart if watcher was active.

Main Window:
- Start/Stop Watching: Toggles the icon detection and key pressing. The main window will minimize while watching.
- Capture Hotkey Display: Shows the currently active hotkey. Press this key anywhere to start capturing an icon.

Listener List:
- Displays configured listeners. Priority is top-down.
- #: Priority order.
- Name: Custom name for the listener.
- Keybind: Key(s) to press (e.g., '1', 'ctrl+s').
- Active: 'Yes' if this listener is currently enabled for watching.
- Confidence: How closely the screen image must match the captured icon (0.1-1.0).

Buttons:
- Remove Selected: Deletes the selected listener from the list.
- Edit Selected: Modify parameters of the selected listener.
- Move Up/Down: Change priority of the selected listener.
- Toggle Active: Enable/disable the selected listener. (Or double-click list item)

How to Use:
1. (Optional) Set your preferred Capture Hotkey via File menu.
2. Press the global Capture Hotkey. The screen will dim.
3. Click and drag a rectangle around the WoW icon you want to track. Release the mouse.
4. Enter a name, the keybind to press, confidence, and other parameters in the dialogs.
5. The listener is added to the list. Adjust its priority with Move Up/Down.
6. Configure multiple listeners as needed.
7. Click "Start Watching". The tool will now monitor for active listeners' icons.

Detection Logic:
- When an icon is found for an active listener (respecting priority):
  - The associated keybind is pressed.
  - If the icon remains, the key is pressed again after a short 'Post-Press Delay'.
  - This repeats up to 'Max Sequential Presses' times or until the icon disappears.
  - After an action, the tool rescans from the highest priority listener.

Failsafe: Quickly move your mouse to the top-left corner of your primary screen to stop PyAutoGUI (and thus the watcher).
"""
        # Use a Toplevel for help to make it easily dismissible
        help_win = tk.Toplevel(self.controller.root)
        help_win.title("Help / Instructions")
        x, y = get_screen_center_for_window(600, 550, self.controller.root)
        help_win.geometry(f"600x550+{x}+{y}")
        help_win.transient(self.controller.root) # Keep on top of main

        text_area = scrolledtext.ScrolledText(help_win, wrap=tk.WORD, height=20, width=70, font=("Arial", 9))
        text_area.insert(tk.INSERT, help_text)
        text_area.config(state=tk.DISABLED, bg="#F0F0F0")
        text_area.pack(pady=10, padx=10, fill="both", expand=True)
        tk.Button(help_win, text="Close", command=help_win.destroy).pack(pady=5)
        help_win.focus_set()


# --- Main Execution ---
if __name__ == "__main__":
    app_root = tk.Tk()
    try:
        # Apply a modern theme if available
        style = ttk.Style(app_root)
        available_themes = style.theme_names()
        if 'clam' in available_themes: style.theme_use('clam')
        elif 'vista' in available_themes: style.theme_use('vista') # Windows
        elif 'aqua' in available_themes: style.theme_use('aqua') # MacOS
    except tk.TclError:
        print("TTK themes not available or failed to apply.")
    
    app_instance = WoWAutomationApp(app_root)
    app_root.mainloop()

