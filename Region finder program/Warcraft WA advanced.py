import tkinter as tk
from tkinter import ttk, simpledialog, filedialog, colorchooser, scrolledtext
import pyautogui
import time
import threading
import json
import os
import random
import shutil
from PIL import Image, ImageTk, ImageGrab # Pillow for pixel color and image ops

# --- Global Variables & Constants ---
DEFAULT_PROJECT_NAME = "WarcraftAutomation"

# --- Helper Functions ---
def get_screen_center_for_window(window_width, window_height, root):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    return x, y

# --- Main Application Class ---
class WarcraftAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Warcraft WA Advanced Automation Tool")
        x, y = get_screen_center_for_window(550, 700, root)
        self.root.geometry(f"+{x}+{y}")

        self.objects = {}
        self.current_listeners = [] # To store icon/keybind listeners
        self.current_project_path = None
        self.current_sequence_name = DEFAULT_PROJECT_NAME
        self.sequence_modified = False

        self.drag_select_window = None
        self.drag_start_x = None
        self.drag_start_y = None
        self.drag_rect_id = None

        self.container = tk.Frame(root)
        self.container.pack(fill="both", expand=True)

        self.frames = {}
        # Define frames for the new application
        # For now, we'll just have a MainFrame and an ObjectCreationFrame
        for F in (MainFrame, ObjectCreationFrame):
            page_name = F.__name__
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("MainFrame")

    def mark_sequence_modified(self, modified=True):
        self.sequence_modified = modified
        current_title = self.root.title()
        base_title = current_title.rstrip("*")

        active_frame_name = ""
        for name, frame_instance in self.frames.items():
            if frame_instance.winfo_ismapped():
                active_frame_name = name
                break

        new_title_base = active_frame_name if active_frame_name else base_title

        if modified:
            self.root.title(new_title_base + "*")
        else:
            self.root.title(new_title_base)

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()
        if hasattr(frame, 'refresh_content') and callable(getattr(frame, 'refresh_content')):
            frame.refresh_content()

        title_base = page_name
        if self.sequence_modified:
            self.root.title(title_base + "*")
        else:
            self.root.title(title_base)

        self.root.update_idletasks()
        self.root.geometry("")

    def get_object_names(self, object_type=None):
        if object_type:
            return [name for name, obj_data in self.objects.items() if obj_data.get("type") == object_type]
        return list(self.objects.keys())

    def add_object(self, name, obj_data):
        if name in self.objects:
            simpledialog.messagebox.showwarning("Warning", f"Object with name '{name}' already exists. Please choose a unique name.", parent=self.root)
            return False
        if not name:
            simpledialog.messagebox.showwarning("Warning", "Object name cannot be empty.", parent=self.root)
            return False
        self.objects[name] = obj_data
        print(f"Added object: {name} - {obj_data}")
        if self.frames["ObjectCreationFrame"].winfo_exists():
             self.frames["ObjectCreationFrame"].update_objects_display()
        # if self.frames["StepCreatorFrame"].winfo_exists(): # No StepCreatorFrame yet
        #      self.frames["StepCreatorFrame"].refresh_object_dropdowns()
        self.mark_sequence_modified()
        return True

    def create_region_drag_mode(self):
        if self.drag_select_window and self.drag_select_window.winfo_exists(): return
        self.drag_select_window = tk.Toplevel(self.root)
        self.drag_select_window.attributes('-fullscreen',True); self.drag_select_window.attributes('-alpha',0.3); self.drag_select_window.attributes('-topmost',True)
        drag_canvas = tk.Canvas(self.drag_select_window,bg="gray",cursor="crosshair",highlightthickness=0); drag_canvas.pack(fill="both",expand=True)
        tk.Label(drag_canvas,text="Click & drag. Release to confirm. ESC to cancel.",bg="lightyellow",fg="black").place(x=10,y=10)
        def on_b1_press(event): self.drag_start_x=event.x; self.drag_start_y=event.y; self.drag_rect_id=drag_canvas.create_rectangle(self.drag_start_x,self.drag_start_y,self.drag_start_x,self.drag_start_y,outline='red',width=2)
        def on_b1_motion(event):
            if self.drag_rect_id: drag_canvas.coords(self.drag_rect_id,self.drag_start_x,self.drag_start_y,event.x,event.y)
        def on_b1_release(event):
            if self.drag_start_x is None: return
            x1,y1=min(self.drag_start_x,event.x),min(self.drag_start_y,event.y); x2,y2=max(self.drag_start_x,event.x),max(self.drag_start_y,event.y)
            self.drag_select_window.destroy(); self.drag_select_window=None; self.drag_start_x,self.drag_start_y,self.drag_rect_id = None,None,None
            if abs(x1-x2)<5 or abs(y1-y2)<5: simpledialog.messagebox.showinfo("Info","Selection too small.",parent=self.root); return
            coords=(int(x1),int(y1),int(x2-x1),int(y2-y1)) # x, y, width, height
            
            # Prompt for icon name
            obj_name = simpledialog.askstring("Name Icon", "Enter a name for this icon:", parent=self.root)
            if not obj_name: return

            # Prompt for keybind
            keybind = simpledialog.askstring("Assign Keybind", f"Enter the keybind for '{obj_name}' (e.g., '1', 'f', 'alt+q'):", parent=self.root)
            if not keybind: return

            try:
                self.root.withdraw(); time.sleep(0.2); img = pyautogui.screenshot(region=coords); self.root.deiconify()
                base_img_filename=f"{obj_name.replace(' ','_').replace('.','_')}.png"; final_abs_img_path=""
                
                # Save image to a project-specific 'icons' directory
                if self.current_project_path:
                    icons_dir=os.path.join(self.current_project_path,"icons"); os.makedirs(icons_dir,exist_ok=True)
                    final_abs_img_path=os.path.join(icons_dir,base_img_filename)
                else:
                    # If no project, save to current working directory and inform user
                    final_abs_img_path=os.path.join(os.getcwd(),base_img_filename)
                    simpledialog.messagebox.showinfo("Icon Saved (No Project)",f"Icon: {final_abs_img_path}\nRelative after 'Save As...'.",parent=self.root)
                
                img.save(final_abs_img_path)
                
                obj_data={
                    "type":"icon",
                    "image_path":final_abs_img_path,
                    "capture_coords":coords,
                    "keybind":keybind,
                    "confidence":0.8 # Default confidence
                }
                if self.add_object(obj_name,obj_data):
                    simpledialog.messagebox.showinfo("Icon Created",f"Icon '{obj_name}' captured and linked to '{keybind}'.",parent=self.root)
                    # Add to listeners list
                    self.current_listeners.append({"name": obj_name, "object_data": obj_data, "active": False, "priority": len(self.current_listeners) + 1})
                    self.frames["MainFrame"].update_listeners_display()

            except Exception as e: self.root.deiconify(); simpledialog.messagebox.showerror("Error",f"Capture icon error: {e}",parent=self.root)
        def on_escape_drag(event=None):
            if self.drag_select_window: self.drag_select_window.destroy(); self.drag_select_window=None
            self.drag_start_x,self.drag_start_y,self.drag_rect_id=None,None,None; print("Drag selection cancelled.")
        drag_canvas.bind("<ButtonPress-1>",on_b1_press); drag_canvas.bind("<B1-Motion>",on_b1_motion)
        drag_canvas.bind("<ButtonRelease-1>",on_b1_release); self.drag_select_window.bind("<Escape>",on_escape_drag)
        self.drag_select_window.focus_force()

    def _check_unsaved_changes(self):
        if self.sequence_modified:
            response = simpledialog.messagebox.askyesnocancel("Unsaved Changes", f"Sequence '{self.current_sequence_name}' has unsaved changes. Save now?", parent=self.root)
            if response is True: return self.save_sequence()
            elif response is False: return True
            else: return False
        return True

    def new_sequence(self):
        if not self._check_unsaved_changes(): return
        self.objects={}; self.current_listeners=[]
        # self.frames["StepCreatorFrame"].clear_and_rebuild() # No StepCreatorFrame yet
        self.current_project_path = None
        self.current_sequence_name = DEFAULT_PROJECT_NAME
        self.mark_sequence_modified(False)
        self.show_frame("MainFrame")
        self.frames["MainFrame"].update_listeners_display()

    def save_sequence(self):
        if not self.current_project_path:
            return self.save_sequence_as()
        
        try:
            project_data = {
                "objects": self.objects,
                "listeners": self.current_listeners,
                "sequence_name": self.current_sequence_name
            }
            
            with open(os.path.join(self.current_project_path, f"{self.current_sequence_name}.json"), "w") as f:
                json.dump(project_data, f, indent=4)
            self.mark_sequence_modified(False)
            simpledialog.messagebox.showinfo("Save Successful", f"Sequence '{self.current_sequence_name}' saved.", parent=self.root)
            return True
        except Exception as e:
            simpledialog.messagebox.showerror("Error", f"Failed to save sequence: {e}", parent=self.root)
            return False

    def save_sequence_as(self):
        project_name = simpledialog.askstring("Save Sequence As", "Enter project name:", initialvalue=self.current_sequence_name, parent=self.root)
        if not project_name: return False

        save_dir = filedialog.askdirectory(parent=self.root, title="Select Directory to Save Project")
        if not save_dir: return False

        project_path = os.path.join(save_dir, project_name)
        os.makedirs(project_path, exist_ok=True)

        self.current_project_path = project_path
        self.current_sequence_name = project_name
        return self.save_sequence()

    def load_sequence(self):
        if not self._check_unsaved_changes(): return

        load_dir = filedialog.askdirectory(parent=self.root, title="Select Project Directory to Load")
        if not load_dir: return

        project_name = os.path.basename(load_dir)
        json_file_path = os.path.join(load_dir, f"{project_name}.json")

        if not os.path.exists(json_file_path):
            simpledialog.messagebox.showerror("Error", f"Project file '{project_name}.json' not found in selected directory.", parent=self.root)
            return

        try:
            with open(json_file_path, "r") as f:
                project_data = json.load(f)
            
            self.objects = project_data.get("objects", {})
            self.current_listeners = project_data.get("listeners", [])
            self.current_sequence_name = project_data.get("sequence_name", project_name)
            self.current_project_path = load_dir
            self.mark_sequence_modified(False)
            simpledialog.messagebox.showinfo("Load Successful", f"Sequence '{self.current_sequence_name}' loaded.", parent=self.root)
            self.show_frame("MainFrame")
            self.frames["MainFrame"].update_listeners_display()

        except Exception as e:
            simpledialog.messagebox.showerror("Error", f"Failed to load sequence: {e}", parent=self.root)

    def start_listening(self):
        # This will be the main loop for monitoring icons and pressing keys
        # Needs to run in a separate thread to not block the UI
        if hasattr(self, '_listener_thread') and self._listener_thread.is_alive():
            simpledialog.messagebox.showinfo("Info", "Listeners are already active.", parent=self.root)
            return

        active_listeners = [l for l in self.current_listeners if l["active"]]
        if not active_listeners:
            simpledialog.messagebox.showinfo("Info", "No active listeners to start.", parent=self.root)
            return

        self._stop_listening_flag = threading.Event()
        self._listener_thread = threading.Thread(target=self._listener_loop, args=(active_listeners, self._stop_listening_flag))
        self._listener_thread.daemon = True # Allow thread to exit with main program
        self._listener_thread.start()
        simpledialog.messagebox.showinfo("Info", "Started monitoring for active icons.", parent=self.root)

    def stop_listening(self):
        if hasattr(self, '_stop_listening_flag'):
            self._stop_listening_flag.set()
            simpledialog.messagebox.showinfo("Info", "Stopping icon monitoring.", parent=self.root)
        else:
            simpledialog.messagebox.showinfo("Info", "No active listeners to stop.", parent=self.root)

    def _listener_loop(self, listeners, stop_flag):
        # Sort listeners by priority (lower number = higher priority)
        listeners.sort(key=lambda x: x["priority"])

        while not stop_flag.is_set():
            found_and_pressed = False
            for listener in listeners:
                obj_data = listener["object_data"]
                if obj_data["type"] == "icon":
                    image_path = obj_data["image_path"]
                    confidence = obj_data.get("confidence", 0.8)
                    keybind = obj_data["keybind"]

                    try:
                        # Check if the icon is on screen
                        # pyautogui.locateOnScreen is resource intensive, consider optimizations
                        location = pyautogui.locateOnScreen(image_path, confidence=confidence)
                        if location:
                            print(f"Found {listener['name']} at {location}. Pressing {keybind}.")
                            # Simulate key press
                            if '+' in keybind: # Handle hotkeys like 'alt+q'
                                keys = keybind.split('+')
                                pyautogui.hotkey(*keys)
                            else:
                                pyautogui.press(keybind)
                            found_and_pressed = True
                            # Add a small delay to prevent rapid re-detection and spamming
                            time.sleep(0.5) 
                            break # Press only the highest priority active icon
                    except pyautogui.ImageNotFoundException:
                        pass # Icon not found, continue to next listener
                    except Exception as e:
                        print(f"Error monitoring {listener['name']}: {e}")
            
            if not found_and_pressed:
                time.sleep(0.1) # Small delay if no icon was found to reduce CPU usage

        print("Listener loop stopped.")


# --- UI Frames ---
class MainFrame(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        label = tk.Label(self, text="Warcraft Automation Main Menu", font=("Arial", 16))
        label.pack(pady=10, padx=10)

        button_frame = tk.Frame(self)
        button_frame.pack(pady=10)

        btn_create_icon = ttk.Button(button_frame, text="Capture New Icon", command=self.controller.create_region_drag_mode)
        btn_create_icon.grid(row=0, column=0, padx=5, pady=5)

        btn_manage_icons = ttk.Button(button_frame, text="Manage Icons/Listeners", command=lambda: self.controller.show_frame("ObjectCreationFrame"))
        btn_manage_icons.grid(row=0, column=1, padx=5, pady=5)

        btn_start_listeners = ttk.Button(button_frame, text="Start Monitoring", command=self.controller.start_listening)
        btn_start_listeners.grid(row=1, column=0, padx=5, pady=5)

        btn_stop_listeners = ttk.Button(button_frame, text="Stop Monitoring", command=self.controller.stop_listening)
        btn_stop_listeners.grid(row=1, column=1, padx=5, pady=5)

        menu_bar = tk.Menu(self.controller.root)
        self.controller.root.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Sequence", command=self.controller.new_sequence)
        file_menu.add_command(label="Save Sequence", command=self.controller.save_sequence)
        file_menu.add_command(label="Save Sequence As...", command=self.controller.save_sequence_as)
        file_menu.add_command(label="Load Sequence", command=self.controller.load_sequence)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.controller.root.quit)

        # Listener List Display
        self.listener_list_frame = ttk.LabelFrame(self, text="Active Listeners (Priority)")
        self.listener_list_frame.pack(pady=10, padx=10, fill="both", expand=True)

        self.listener_canvas = tk.Canvas(self.listener_list_frame)
        self.listener_canvas.pack(side="left", fill="both", expand=True)

        self.listener_scrollbar = ttk.Scrollbar(self.listener_list_frame, orient="vertical", command=self.listener_canvas.yview)
        self.listener_scrollbar.pack(side="right", fill="y")

        self.listener_canvas.configure(yscrollcommand=self.listener_scrollbar.set)
        self.listener_canvas.bind('<Configure>', lambda e: self.listener_canvas.configure(scrollregion = self.listener_canvas.bbox("all")))

        self.listeners_inner_frame = tk.Frame(self.listener_canvas)
        self.listener_canvas.create_window((0, 0), window=self.listeners_inner_frame, anchor="nw", width=self.listener_canvas.winfo_width())

        self.listeners_inner_frame.bind("<Configure>", lambda e: self.listener_canvas.configure(scrollregion = self.listener_canvas.bbox("all")))

        self.update_listeners_display()

    def update_listeners_display(self):
        for widget in self.listeners_inner_frame.winfo_children():
            widget.destroy()

        # Sort listeners by priority
        sorted_listeners = sorted(self.controller.current_listeners, key=lambda x: x["priority"])

        if not sorted_listeners:
            tk.Label(self.listeners_inner_frame, text="No listeners configured yet.").pack(pady=5)
            return

        for i, listener in enumerate(sorted_listeners):
            listener_frame = tk.Frame(self.listeners_inner_frame, bd=1, relief="solid")
            listener_frame.pack(fill="x", padx=5, pady=2)

            # Priority label
            priority_label = tk.Label(listener_frame, text=f"Prio {listener['priority']}:")
            priority_label.pack(side="left", padx=5)

            # Name and Keybind
            name_keybind_label = tk.Label(listener_frame, text=f"{listener['name']} (Key: {listener['object_data']['keybind']})")
            name_keybind_label.pack(side="left", expand=True, fill="x")

            # Active Toggle
            active_var = tk.BooleanVar(value=listener["active"])
            active_check = ttk.Checkbutton(listener_frame, text="Active", variable=active_var, 
                                           command=lambda l=listener, var=active_var: self.toggle_listener_active(l, var.get()))
            active_check.pack(side="right", padx=5)

            # Move Up/Down buttons
            btn_up = ttk.Button(listener_frame, text="↑", width=3, command=lambda l=listener: self.move_listener_priority(l, -1))
            btn_up.pack(side="right", padx=2)
            btn_down = ttk.Button(listener_frame, text="↓", width=3, command=lambda l=listener: self.move_listener_priority(l, 1))
            btn_down.pack(side="right", padx=2)

        self.listeners_inner_frame.update_idletasks()
        self.listener_canvas.config(scrollregion=self.listener_canvas.bbox("all"))

    def toggle_listener_active(self, listener, is_active):
        listener["active"] = is_active
        self.controller.mark_sequence_modified()
        print(f"Listener {listener['name']} active status set to {is_active}")

    def move_listener_priority(self, listener_to_move, direction):
        # direction: -1 for up, 1 for down
        current_priority = listener_to_move["priority"]
        new_priority = current_priority + direction

        # Find the listener that will swap with the current one
        for listener in self.controller.current_listeners:
            if listener["priority"] == new_priority:
                listener["priority"] = current_priority
                listener_to_move["priority"] = new_priority
                self.controller.mark_sequence_modified()
                self.update_listeners_display()
                return
        print(f"Cannot move {listener_to_move['name']} further in direction {direction}")

    def refresh_content(self):
        self.update_listeners_display()
        # Adjust canvas scroll region when frame is shown
        self.listeners_inner_frame.update_idletasks()
        self.listener_canvas.config(scrollregion=self.listener_canvas.bbox("all"))


class ObjectCreationFrame(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        label = tk.Label(self, text="Manage Captured Icons", font=("Arial", 16))
        label.pack(pady=10, padx=10)

        btn_back = ttk.Button(self, text="Back to Main Menu", command=lambda: self.controller.show_frame("MainFrame"))
        btn_back.pack(pady=5)

        # Display existing objects (icons)
        self.objects_display_frame = ttk.LabelFrame(self, text="Captured Icons")
        self.objects_display_frame.pack(pady=10, padx=10, fill="both", expand=True)

        self.objects_canvas = tk.Canvas(self.objects_display_frame)
        self.objects_canvas.pack(side="left", fill="both", expand=True)

        self.objects_scrollbar = ttk.Scrollbar(self.objects_display_frame, orient="vertical", command=self.objects_canvas.yview)
        self.objects_scrollbar.pack(side="right", fill="y")

        self.objects_canvas.configure(yscrollcommand=self.objects_scrollbar.set)
        self.objects_canvas.bind('<Configure>', lambda e: self.objects_canvas.configure(scrollregion = self.objects_canvas.bbox("all")))

        self.objects_inner_frame = tk.Frame(self.objects_canvas)
        self.objects_canvas.create_window((0, 0), window=self.objects_inner_frame, anchor="nw", width=self.objects_canvas.winfo_width())

        self.objects_inner_frame.bind("<Configure>", lambda e: self.objects_canvas.configure(scrollregion = self.objects_canvas.bbox("all")))

        self.update_objects_display()

    def update_objects_display(self):
        for widget in self.objects_inner_frame.winfo_children():
            widget.destroy()

        icons = self.controller.get_object_names(object_type="icon")
        if not icons:
            tk.Label(self.objects_inner_frame, text="No icons captured yet.").pack(pady=5)
            return

        for icon_name in icons:
            obj_data = self.controller.objects[icon_name]
            icon_frame = tk.Frame(self.objects_inner_frame, bd=1, relief="groove")
            icon_frame.pack(fill="x", padx=5, pady=2)

            tk.Label(icon_frame, text=f"Name: {icon_name}").pack(anchor="w")
            tk.Label(icon_frame, text=f"Keybind: {obj_data['keybind']}").pack(anchor="w")
            tk.Label(icon_frame, text=f"Path: {os.path.basename(obj_data['image_path'])}").pack(anchor="w")

            # Add a delete button
            delete_btn = ttk.Button(icon_frame, text="Delete", command=lambda name=icon_name: self.delete_object(name))
            delete_btn.pack(pady=2)

        self.objects_inner_frame.update_idletasks()
        self.objects_canvas.config(scrollregion=self.objects_canvas.bbox("all"))

    def delete_object(self, obj_name):
        if simpledialog.messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{obj_name}'?", parent=self.controller.root):
            if obj_name in self.controller.objects:
                # Remove associated image file
                image_path = self.controller.objects[obj_name].get("image_path")
                if image_path and os.path.exists(image_path):
                    try:
                        os.remove(image_path)
                        print(f"Deleted image file: {image_path}")
                    except Exception as e:
                        print(f"Error deleting image file {image_path}: {e}")

                del self.controller.objects[obj_name]
                # Also remove from current_listeners if it exists there
                self.controller.current_listeners = [l for l in self.controller.current_listeners if l["name"] != obj_name]
                # Re-assign priorities after deletion
                for i, listener in enumerate(self.controller.current_listeners):
                    listener["priority"] = i + 1

                self.controller.mark_sequence_modified()
                self.update_objects_display()
                self.controller.frames["MainFrame"].update_listeners_display()
                simpledialog.messagebox.showinfo("Deleted", f"'{obj_name}' has been deleted.", parent=self.controller.root)

    def refresh_content(self):
        self.update_objects_display()
        # Adjust canvas scroll region when frame is shown
        self.objects_inner_frame.update_idletasks()
        self.objects_canvas.config(scrollregion=self.objects_canvas.bbox("all"))


if __name__ == "__main__":
    root = tk.Tk()
    app = WarcraftAutomationApp(root)
    root.mainloop()