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
# For features like drag-to-select or global hotkeys if needed later
# For now, pynput might be an overkill if not using global hotkeys for recording.
# PyAutoGUI and Tkinter's own event handling will cover a lot.

DEFAULT_PROJECT_NAME = "UntitledSequence"

# --- Helper Functions ---
def get_screen_center_for_window(window_width, window_height, root):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    return x, y

# --- Main Application Class ---
class DesktopAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Python Desktop Automation Tool")
        # Allow window to resize dynamically to fit content
        # Position window at center of screen
        x, y = get_screen_center_for_window(400, 650, root)
        self.root.geometry(f"+{x}+{y}")


        # Core data structures
        self.objects = {}  # Stores all created objects (regions, images, pixels)
                           # e.g., {"ObjectName": {"type": "region", "coords": (x,y,w,h), ...}}

        self.current_steps = [] # List of dicts: {"object_name": str, "action": str, "params": {}}

        self.current_project_path = None # Path to the folder where sequence and its assets are saved
        self.current_sequence_name = DEFAULT_PROJECT_NAME
        self.loop_count = tk.IntVar(value=1) # 0 for indefinite
        self.sequence_modified = False # Track if changes have been made

        # For Grid Overlay
        self.grid_window = None
        self.grid_rows_var = tk.IntVar(value=10)
        self.grid_cols_var = tk.IntVar(value=10)
        self.selected_grid_cells = []

        # For Drag Selection
        self.drag_select_window = None
        self.drag_start_x = None
        self.drag_start_y = None
        self.drag_rect_id = None

        # For Pixel Monitor
        self.pixel_monitor_active = False
        self._pixel_listener = None # Using pynput if truly global click needed

        # --- UI Frames ---
        self.container = tk.Frame(root)
        self.container.pack(fill="both", expand=True)

        self.frames = {}
        for F in (MainFrame, ObjectCreationFrame, StepCreatorFrame, InstructionsFrame):
            page_name = F.__name__
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("MainFrame")

    def mark_sequence_modified(self, modified=True):
        self.sequence_modified = modified
        # Update window title to indicate unsaved changes
        title = self.root.title()
        if modified and not title.endswith("*"):
            self.root.title(title + "*")
        elif not modified and title.endswith("*"):
            self.root.title(title[:-1])

    def show_frame(self, page_name):
        """Show a frame for the given page name"""
        frame = self.frames[page_name]
        frame.tkraise()
        # If the frame has a 'refresh_content' method, call it
        if hasattr(frame, 'refresh_content') and callable(getattr(frame, 'refresh_content')):
            frame.refresh_content()
        self.root.title(page_name) # Optional: update window title

        # Force window to resize to fit content
        self.root.update_idletasks()
        self.root.geometry("")  # Reset any explicit geometry to allow natural sizing

    def get_object_names(self, object_type=None):
        """Returns a list of names for objects, optionally filtered by type."""
        if object_type:
            return [name for name, obj_data in self.objects.items() if obj_data.get("type") == object_type]
        return list(self.objects.keys())

    def add_object(self, name, obj_data):
        if name in self.objects:
            simpledialog.messagebox.showwarning("Warning", f"Object with name '{name}' already exists. Please choose a unique name.")
            return False
        if not name:
            simpledialog.messagebox.showwarning("Warning", "Object name cannot be empty.")
            return False
        self.objects[name] = obj_data
        print(f"Added object: {name} - {obj_data}")
        if self.frames["ObjectCreationFrame"].winfo_exists(): # Refresh object list if visible
             self.frames["ObjectCreationFrame"].update_objects_display()
        if self.frames["StepCreatorFrame"].winfo_exists(): # Refresh step creator if visible
             self.frames["StepCreatorFrame"].refresh_object_dropdowns()
        self.mark_sequence_modified() # Mark as modified
        return True

    def _start_pixel_monitor_listener(self):
        """Uses a simple Tkinter binding for pixel selection within a temporary window."""
        if self.pixel_monitor_active:
            return
        self.pixel_monitor_active = True

        self.pixel_capture_instruction_window = tk.Toplevel(self.root)
        self.pixel_capture_instruction_window.attributes('-topmost', True)
        self.pixel_capture_instruction_window.geometry("300x100+100+100")
        tk.Label(self.pixel_capture_instruction_window,
                 text="Pixel Monitor Active\nClick anywhere on your primary screen to select a pixel.\nPress ESC in this small window to cancel.").pack(pady=20)
        self.pixel_capture_instruction_window.focus_force()

        # This is a bit tricky. A truly global click usually needs pynput.
        # For a simpler prototype, we can make the user click on the main screen
        # while this small window is up, then use pyautogui.position()
        # However, a better UX is a temporary fullscreen overlay.
        # Let's try with a short delay and pyautogui.position() for simplicity first.
        # And an Escape binding on the instruction window to cancel.

        def _on_pixel_capture_escape(event=None):
            self.pixel_monitor_active = False
            if self.pixel_capture_instruction_window and self.pixel_capture_instruction_window.winfo_exists():
                self.pixel_capture_instruction_window.destroy()
            print("Pixel monitoring cancelled.")

        self.pixel_capture_instruction_window.bind("<Escape>", _on_pixel_capture_escape)

        # We need a way to detect the click *outside* this instruction window.
        # The most robust way is a global listener (pynput).
        # A simpler, but less ideal way for a prototype:
        # After the user clicks this, we assume their next click is the target.
        # This requires user discipline.
        # For a slightly better approach, we can use a temporary fullscreen transparent window.

        # For now, let's use a "Capture Now" button in the small window
        tk.Button(self.pixel_capture_instruction_window, text="Capture Pixel Under Mouse NOW",
                  command=self._capture_pixel_under_mouse).pack(pady=5)

    def _capture_pixel_under_mouse(self):
        if not self.pixel_monitor_active: return

        try:
            x, y = pyautogui.position()
            rgb = pyautogui.pixel(x, y)

            obj_name = simpledialog.askstring("Name Pixel Object", "Enter a name for this pixel object:", parent=self.root)
            if obj_name:
                obj_data = {"type": "pixel", "coords": (x, y), "rgb": rgb}
                if self.add_object(obj_name, obj_data):
                    simpledialog.messagebox.showinfo("Pixel Captured",
                                                     f"Pixel '{obj_name}' captured at ({x},{y}) with RGB: {rgb}",
                                                     parent=self.root)

        except Exception as e:
            simpledialog.messagebox.showerror("Error", f"Could not capture pixel: {e}", parent=self.root)
        finally:
            self.pixel_monitor_active = False
            if self.pixel_capture_instruction_window and self.pixel_capture_instruction_window.winfo_exists():
                self.pixel_capture_instruction_window.destroy()


    def create_region_grid_mode(self):
        """Opens the grid overlay for region selection."""
        if self.grid_window and self.grid_window.winfo_exists():
            self.grid_window.destroy()

        self.grid_window = tk.Toplevel(self.root)
        self.grid_window.attributes('-fullscreen', True)
        self.grid_window.attributes('-alpha', 0.4)  # Semi-transparent
        self.grid_window.attributes('-topmost', True)

        self.grid_canvas = tk.Canvas(self.grid_window, bg='gray', highlightthickness=0) # Darker transparent bg
        self.grid_canvas.pack(fill="both", expand=True)

        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()

        try:
            rows = self.grid_rows_var.get()
            cols = self.grid_cols_var.get()
            if rows <= 0 or cols <= 0: raise ValueError("Grid dimensions must be positive.")
        except (tk.TclError, ValueError) as e:
            simpledialog.messagebox.showerror("Error", f"Invalid grid dimensions: {e}. Using 10x10.", parent=self.root)
            self.grid_rows_var.set(10)
            self.grid_cols_var.set(10)
            rows, cols = 10, 10

        self.cell_width = self.screen_width / cols
        self.cell_height = self.screen_height / rows
        self.selected_grid_cells = [] # Reset selection

        self._draw_grid_on_canvas()
        self.grid_canvas.bind("<Button-1>", self._on_grid_cell_click)
        self.grid_window.bind("<Escape>", lambda e: self._confirm_grid_selection(cancelled=True))

        # Add a small confirmation bar/button
        confirm_bar = tk.Frame(self.grid_canvas, bg="lightgray", relief=tk.RAISED, borderwidth=1) # On canvas
        tk.Label(confirm_bar, text=f"{rows}x{cols} Grid. Click cells. ESC to cancel.", bg="lightgray").pack(side=tk.LEFT, padx=10)
        tk.Button(confirm_bar, text="Confirm Selection", command=self._confirm_grid_selection).pack(side=tk.LEFT, padx=10)
        self.grid_canvas.create_window(self.screen_width // 2, 30, window=confirm_bar, anchor="n") # Place bar at top-center

        self.grid_window.focus_force()

    def _draw_grid_on_canvas(self):
        self.grid_canvas.delete("grid_line") # Clear only lines
        self.grid_canvas.delete("cell_highlight") # Clear only highlights

        rows = self.grid_rows_var.get()
        cols = self.grid_cols_var.get()

        for r in range(rows):
            for c in range(cols):
                x1, y1 = c * self.cell_width, r * self.cell_height
                x2, y2 = x1 + self.cell_width, y1 + self.cell_height

                # Draw cell selection highlight
                if (r,c) in self.selected_grid_cells:
                    self.grid_canvas.create_rectangle(x1, y1, x2, y2,
                                                      fill="blue", outline="lightblue", stipple="gray50",
                                                      tags="cell_highlight")
                # Draw grid lines
                if r < rows:
                    self.grid_canvas.create_line(0, y2, self.screen_width, y2, fill="white", tags="grid_line", width=0.5) # Lighter lines
                if c < cols:
                    self.grid_canvas.create_line(x2, 0, x2, self.screen_height, fill="white", tags="grid_line", width=0.5)


    def _on_grid_cell_click(self, event):
        col = int(event.x // self.cell_width)
        row = int(event.y // self.cell_height)
        cell = (row, col)
        if cell in self.selected_grid_cells:
            self.selected_grid_cells.remove(cell)
        else:
            self.selected_grid_cells.append(cell)
        self._draw_grid_on_canvas() # Redraw highlights

    def _confirm_grid_selection(self, cancelled=False):
        if cancelled or not self.selected_grid_cells:
            if self.grid_window and self.grid_window.winfo_exists():
                self.grid_window.destroy()
            self.selected_grid_cells = []
            if not cancelled:
                 simpledialog.messagebox.showinfo("Info", "No cells selected.", parent=self.root)
            return

        min_r = min(r for r, c in self.selected_grid_cells)
        max_r = max(r for r, c in self.selected_grid_cells)
        min_c = min(c for r, c in self.selected_grid_cells)
        max_c = max(c for r, c in self.selected_grid_cells)

        x1 = min_c * self.cell_width
        y1 = min_r * self.cell_height
        # For PyAutoGUI region, width and height are needed
        width = (max_c - min_c + 1) * self.cell_width
        height = (max_r - min_r + 1) * self.cell_height

        coords = (int(x1), int(y1), int(width), int(height))

        # Determine if this is for a Region Object or Image Object based on context
        # For now, assume it's for a Region Object from ObjectCreationFrame
        creation_type = self.frames["ObjectCreationFrame"].current_creation_type

        obj_name = simpledialog.askstring(f"Name {creation_type.capitalize()} Object", f"Enter name for the selected {creation_type}:", parent=self.root)
        if obj_name:
            if creation_type == "region":
                obj_data = {"type": "region", "mode": "grid", "coords": coords, "cells": list(self.selected_grid_cells)}
                if self.add_object(obj_name, obj_data):
                    simpledialog.messagebox.showinfo("Region Created", f"Region '{obj_name}' created from grid.", parent=self.root)
            elif creation_type == "image":
                # Capture image from this region
                try:
                    # Ensure main window is not obscuring if coordinates are relative to it
                    self.root.withdraw() # Hide main window briefly
                    time.sleep(0.2) # Give time for window to hide
                    img = pyautogui.screenshot(region=coords)
                    self.root.deiconify() # Show main window again

                    base_img_filename = f"{obj_name.replace(' ', '_').replace('.', '_')}.png" # Sanitize name
                    final_abs_img_path = "" # Will hold the true absolute path of the image

                    if self.current_project_path:
                        images_dir = os.path.join(self.current_project_path, "images")
                        os.makedirs(images_dir, exist_ok=True)
                        final_abs_img_path = os.path.join(images_dir, base_img_filename)
                    else:
                        # If no project, save to CWD. This path will be updated on first "Save As".
                        final_abs_img_path = os.path.join(os.getcwd(), base_img_filename)
                        simpledialog.messagebox.showinfo("Image Saved (No Project)",
                                                         f"Image will be saved to:\n{final_abs_img_path}\nPath will be project-relative after 'Save As...'.",
                                                         parent=self.root)
                    img.save(final_abs_img_path)

                    # The path stored in the object in memory is ALWAYS absolute.
                    # The path stored in JSON will be made relative during save.
                    obj_data = {"type": "image", "mode": "grid",
                                "image_path": final_abs_img_path, # STORE ABSOLUTE PATH IN MEMORY
                                "capture_coords": coords, "confidence": 0.8}
                    if self.add_object(obj_name, obj_data):
                         simpledialog.messagebox.showinfo("Image Created", f"Image '{obj_name}' captured.", parent=self.root)
                         # self.mark_sequence_modified() is called by add_object

                except Exception as e:
                    self.root.deiconify()
                    simpledialog.messagebox.showerror("Error", f"Could not capture image: {e}", parent=self.root)

        if self.grid_window and self.grid_window.winfo_exists():
            self.grid_window.destroy()
        self.selected_grid_cells = []

    def create_region_drag_mode(self):
        """Activates drag selection mode."""
        if self.drag_select_window and self.drag_select_window.winfo_exists():
            return # Already active

        self.drag_select_window = tk.Toplevel(self.root)
        self.drag_select_window.attributes('-fullscreen', True)
        self.drag_select_window.attributes('-alpha', 0.3) # Semi-transparent
        self.drag_select_window.attributes('-topmost', True)
        # self.drag_select_window.overrideredirect(True) # Optional: no window decorations

        drag_canvas = tk.Canvas(self.drag_select_window, bg="gray", cursor="crosshair", highlightthickness=0)
        drag_canvas.pack(fill="both", expand=True)

        tk.Label(drag_canvas, text="Click and drag to select region. Release mouse to confirm. ESC to cancel.",
                 bg="lightyellow", fg="black").place(x=10, y=10)

        def on_b1_press(event):
            self.drag_start_x = event.x
            self.drag_start_y = event.y
            self.drag_rect_id = drag_canvas.create_rectangle(self.drag_start_x, self.drag_start_y,
                                                             self.drag_start_x, self.drag_start_y,
                                                             outline='red', width=2)
        def on_b1_motion(event):
            if self.drag_rect_id:
                drag_canvas.coords(self.drag_rect_id, self.drag_start_x, self.drag_start_y, event.x, event.y)

        def on_b1_release(event):
            if self.drag_start_x is None: return # Click without drag

            x1, y1 = min(self.drag_start_x, event.x), min(self.drag_start_y, event.y)
            x2, y2 = max(self.drag_start_x, event.x), max(self.drag_start_y, event.y)

            self.drag_select_window.destroy()
            self.drag_select_window = None
            self.drag_start_x, self.drag_start_y, self.drag_rect_id = None, None, None

            if abs(x1-x2) < 5 or abs(y1-y2) < 5 : # Too small
                simpledialog.messagebox.showinfo("Info", "Selection too small.", parent=self.root)
                return

            coords = (int(x1), int(y1), int(x2-x1), int(y2-y1)) # x, y, width, height

            creation_type = self.frames["ObjectCreationFrame"].current_creation_type
            obj_name = simpledialog.askstring(f"Name {creation_type.capitalize()} Object", f"Enter name for the selected {creation_type}:", parent=self.root)
            if obj_name:
                if creation_type == "region":
                    obj_data = {"type": "region", "mode": "drag", "coords": coords}
                    if self.add_object(obj_name, obj_data):
                         simpledialog.messagebox.showinfo("Region Created", f"Region '{obj_name}' created by dragging.", parent=self.root)
                elif creation_type == "image":
                    try:
                        self.root.withdraw()
                        time.sleep(0.2)
                        img = pyautogui.screenshot(region=coords)
                        self.root.deiconify()

                        base_img_filename = f"{obj_name.replace(' ', '_').replace('.', '_')}.png"
                        final_abs_img_path = ""

                        if self.current_project_path:
                            images_dir = os.path.join(self.current_project_path, "images")
                            os.makedirs(images_dir, exist_ok=True)
                            final_abs_img_path = os.path.join(images_dir, base_img_filename)
                        else:
                            final_abs_img_path = os.path.join(os.getcwd(), base_img_filename)
                            simpledialog.messagebox.showinfo("Image Saved (No Project)",
                                                             f"Image will be saved to:\n{final_abs_img_path}\nPath will be project-relative after 'Save As...'.",
                                                             parent=self.root)
                        img.save(final_abs_img_path)

                        obj_data = {"type": "image", "mode": "drag",
                                    "image_path": final_abs_img_path, # STORE ABSOLUTE PATH IN MEMORY
                                    "capture_coords": coords, "confidence": 0.8}
                        if self.add_object(obj_name, obj_data):
                             simpledialog.messagebox.showinfo("Image Created", f"Image '{obj_name}' captured.", parent=self.root)
                             # self.mark_sequence_modified() is called by add_object
                    except Exception as e:
                        self.root.deiconify()
                        simpledialog.messagebox.showerror("Error", f"Could not capture image: {e}", parent=self.root)

        def on_escape_drag(event=None):
            if self.drag_select_window:
                self.drag_select_window.destroy()
                self.drag_select_window = None
            self.drag_start_x, self.drag_start_y, self.drag_rect_id = None, None, None
            print("Drag selection cancelled.")

        drag_canvas.bind("<ButtonPress-1>", on_b1_press)
        drag_canvas.bind("<B1-Motion>", on_b1_motion)
        drag_canvas.bind("<ButtonRelease-1>", on_b1_release)
        self.drag_select_window.bind("<Escape>", on_escape_drag)
        self.drag_select_window.focus_force()


    def _check_unsaved_changes(self):
        if self.sequence_modified:
            response = simpledialog.messagebox.askyesnocancel("Unsaved Changes",
                                                              f"Sequence '{self.current_sequence_name}' has unsaved changes. Save now?",
                                                              parent=self.root)
            if response is True: # Yes
                return self.save_sequence() # Returns True on success, False on cancel
            elif response is False: # No
                return True # Proceed without saving
            else: # Cancel
                return False # Abort current operation
        return True # No unsaved changes or user chose to proceed

    def new_sequence(self):
        if not self._check_unsaved_changes():
            return

        self.objects = {}
        self.current_steps = []
        self.frames["StepCreatorFrame"].clear_and_rebuild_steps([]) # Clear UI steps
        self.loop_count.set(1)
        self.current_project_path = None
        self.current_sequence_name = DEFAULT_PROJECT_NAME
        self.mark_sequence_modified(False)

        # Refresh UIs
        self.frames["ObjectCreationFrame"].update_objects_display()
        self.frames["MainFrame"].refresh_content()
        # self.show_frame("MainFrame") # Optionally switch to main frame
        print("New sequence created.")


    def save_sequence(self):
        if not self.current_project_path:
            return self.save_sequence_as() # If no path, it's a "Save As"
        else:
            # Finalize steps from UI to self.current_steps
            self.frames["StepCreatorFrame"].finalize_steps_for_controller()

            project_dir = self.current_project_path
            sequence_filename = os.path.join(project_dir, f"{self.current_sequence_name}.json")
            project_images_dir = os.path.join(project_dir, "images") # Target images directory
            os.makedirs(project_images_dir, exist_ok=True)

            data_to_save = {
                "sequence_name": self.current_sequence_name,
                "loop_count": self.loop_count.get(),
                "objects": {},
                "steps": self.current_steps
            }

            for obj_name, obj_data_in_memory in self.objects.items():
                obj_data_for_json = obj_data_in_memory.copy() # Work on a copy for JSON

                if obj_data_for_json.get("type") == "image":
                    current_abs_image_path = obj_data_in_memory.get("image_path") # This should be absolute

                    if not current_abs_image_path or not os.path.isabs(current_abs_image_path):
                        print(f"Warning: Image object '{obj_name}' has invalid or non-absolute path: {current_abs_image_path}. Skipping asset handling.")
                        # Keep whatever path it has for JSON, might cause load issues
                        data_to_save["objects"][obj_name] = obj_data_for_json
                        continue

                    if not os.path.exists(current_abs_image_path):
                        print(f"Warning: Image file for '{obj_name}' not found at: {current_abs_image_path}. Storing path as is.")
                        # Path in JSON will be this (likely broken) absolute path.
                        data_to_save["objects"][obj_name] = obj_data_for_json
                        continue

                    img_basename = os.path.basename(current_abs_image_path)
                    # Target absolute path for this image within the project's images folder
                    target_abs_path_in_project_images = os.path.join(project_images_dir, img_basename)

                    # Normalize paths for reliable comparison
                    norm_current_path = os.path.normpath(current_abs_image_path)
                    norm_target_path = os.path.normpath(target_abs_path_in_project_images)

                    # If the image is not already in the project's images folder, copy it.
                    if norm_current_path != norm_target_path:
                        try:
                            shutil.copy2(current_abs_image_path, target_abs_path_in_project_images)
                            print(f"Copied image for '{obj_name}' to project: {target_abs_path_in_project_images}")
                            # Update the in-memory object to point to the new copy in the project
                            self.objects[obj_name]["image_path"] = target_abs_path_in_project_images
                        except Exception as e:
                            print(f"Error copying image {current_abs_image_path} to project: {e}")
                            simpledialog.messagebox.showerror("Save Error", f"Could not copy image asset {img_basename} for {obj_name}", parent=self.root)
                            # Continue saving, but the JSON will point to the original absolute path
                            # which might be outside the project structure.

                    # For JSON, store the path relative to the project's images folder
                    obj_data_for_json["image_path"] = os.path.join("images", img_basename)

                data_to_save["objects"][obj_name] = obj_data_for_json

            try:
                with open(sequence_filename, 'w') as f:
                    json.dump(data_to_save, f, indent=4)
                simpledialog.messagebox.showinfo("Save Sequence", f"Sequence '{self.current_sequence_name}' saved.", parent=self.root)
                self.mark_sequence_modified(False)
                self.frames["MainFrame"].refresh_content() # Update loaded sequence label
                return True
            except Exception as e:
                simpledialog.messagebox.showerror("Save Error", f"Could not save sequence: {e}", parent=self.root)
                return False

    def save_sequence_as(self):
        # Finalize steps from UI to self.current_steps
        self.frames["StepCreatorFrame"].finalize_steps_for_controller()

        # Ask for a directory to save the project
        project_dir = filedialog.askdirectory(title="Select Project Folder for Sequence", parent=self.root)
        if not project_dir:
            return False # User cancelled

        # Ask for sequence name (will be folder name and json file name)
        # Default to current name if not "UntitledSequence"
        default_name = self.current_sequence_name if self.current_sequence_name != DEFAULT_PROJECT_NAME else "MyNewSequence"
        seq_name = simpledialog.askstring("Sequence Name", "Enter a name for this sequence:",
                                          initialvalue=default_name, parent=self.root)
        if not seq_name:
            return False # User cancelled

        self.current_project_path = os.path.join(project_dir, seq_name) # Project is a folder named after sequence
        self.current_sequence_name = seq_name
        # Now call the regular save_sequence which will use the new path
        return self.save_sequence()


    def load_sequence(self):
        if not self._check_unsaved_changes():
            return

        filepath = filedialog.askopenfilename(
            title="Load Sequence File",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            parent=self.root
        )
        if not filepath:
            return # User cancelled

        try:
            with open(filepath, 'r') as f:
                loaded_data = json.load(f)

            # Set project context before processing objects
            self.current_project_path = os.path.dirname(filepath) # The folder containing the .json
            self.current_sequence_name = loaded_data.get("sequence_name", os.path.splitext(os.path.basename(filepath))[0])

            # Process loaded objects for image paths
            temp_objects = {}
            for obj_name, obj_data in loaded_data.get("objects", {}).items():
                obj_copy = obj_data.copy()
                if obj_copy.get("type") == "image" and obj_copy.get("image_path"):
                    relative_img_path = obj_copy["image_path"]
                    # Convert relative path to absolute based on the project directory
                    abs_image_path = os.path.join(self.current_project_path, relative_img_path)
                    if os.path.exists(abs_image_path):
                        obj_copy["image_path"] = abs_image_path # Store absolute path in memory
                    else:
                        print(f"Warning: Image asset not found for '{obj_name}' at expected location: {abs_image_path}")
                        # Keep the relative path, it might be resolvable if user moves project
                        # Or could mark as missing / ask user to locate
                temp_objects[obj_name] = obj_copy
            self.objects = temp_objects

            self.current_steps = loaded_data.get("steps", [])
            self.loop_count.set(loaded_data.get("loop_count", 1))

            # Update UI
            self.frames["ObjectCreationFrame"].update_objects_display()
            self.frames["StepCreatorFrame"].clear_and_rebuild_steps(self.current_steps)
            self.frames["MainFrame"].refresh_content() # Update loaded sequence label
            self.mark_sequence_modified(False)

            simpledialog.messagebox.showinfo("Load Sequence", f"Sequence '{self.current_sequence_name}' loaded.", parent=self.root)

        except Exception as e:
            simpledialog.messagebox.showerror("Load Error", f"Could not load sequence: {e}", parent=self.root)
            # Reset to a clean state if load fails badly
            self.new_sequence() # Or some other error handling

    def run_sequence(self):
        # Ensure the controller's step list is up-to-date with the UI
        # This is crucial if changes were made in StepCreatorFrame without saving
        self.frames["StepCreatorFrame"].finalize_steps_for_controller()

        if not self.current_steps:
            simpledialog.messagebox.showinfo("Run Sequence", "No steps in the current sequence to run.", parent=self.root)
            return

        try:
            loops = self.loop_count.get()
            if loops < 0:
                simpledialog.messagebox.showerror("Error", "Loop count cannot be negative.", parent=self.root)
                return
        except tk.TclError:
            simpledialog.messagebox.showerror("Error", "Invalid loop count. Must be an integer.", parent=self.root)
            return

        is_infinite = (loops == 0)
        current_loop = 0

        self.root.iconify() # Minimize main window during run
        time.sleep(0.5) # Give time for window to minimize

        pyautogui.FAILSAFE = True # Default failsafe
        # pyautogui.PAUSE = 0.1 # Default pause between actions (can be overridden by Wait steps)

        print(f"--- Running Sequence: {self.current_sequence_name} ---")
        if is_infinite: print("Looping indefinitely. Press Ctrl+C in terminal to stop (or trigger PyAutoGUI failsafe).")
        else: print(f"Looping {loops} times.")

        try:
            while is_infinite or current_loop < loops:
                current_loop += 1
                if not is_infinite: print(f"Executing Loop {current_loop}/{loops}")
                else: print(f"Executing Loop {current_loop}")

                for i, step in enumerate(self.current_steps):
                    print(f"  Step {i+1}: Action: {step['action']}, Object: {step.get('object_name', 'N/A')}, Params: {step.get('params', {})}")

                    obj_name = step.get("object_name")
                    action = step.get("action")
                    params = step.get("params", {})
                    target_object = self.objects.get(obj_name) if obj_name else None

                    try:
                        # --- OBJECT-SPECIFIC ACTIONS ---
                        if target_object:
                            obj_type = target_object.get("type")
                            obj_coords = target_object.get("coords") # (x,y,w,h) for region/image, (x,y) for pixel

                            if action == "Click": # Generic click, behavior depends on type
                                if obj_type == "region" and obj_coords:
                                    # Click center of region
                                    center_x = obj_coords[0] + obj_coords[2] / 2
                                    center_y = obj_coords[1] + obj_coords[3] / 2
                                    pyautogui.click(center_x, center_y)
                                elif obj_type == "pixel" and obj_coords:
                                    pyautogui.click(obj_coords[0], obj_coords[1])
                                elif obj_type == "image" and target_object.get("image_path"):
                                    loc = pyautogui.locateCenterOnScreen(target_object["image_path"], confidence=target_object.get("confidence", 0.8))
                                    if loc:
                                        pyautogui.click(loc)
                                    else:
                                        print(f"    WARN: Image '{obj_name}' not found for click.")
                                        # Potentially raise error or have option to skip/stop
                                else:
                                     print(f"    WARN: Cannot perform 'Click' on object '{obj_name}' of type '{obj_type}'.")

                            elif action == "Wait for Image" and obj_type == "image": # Wait for THIS image object
                                start_time = time.time()
                                timeout = params.get("timeout_s", 10)
                                found = False
                                print(f"    Waiting for image '{obj_name}' (timeout: {timeout}s)")
                                while time.time() - start_time < timeout:
                                    if pyautogui.locateOnScreen(target_object["image_path"], confidence=target_object.get("confidence", 0.8)):
                                        print(f"    Image '{obj_name}' found.")
                                        found = True
                                        break
                                    time.sleep(0.5) # Check interval
                                if not found:
                                    print(f"    TIMEOUT: Image '{obj_name}' not found after {timeout}s.")
                                    # Add option to stop sequence on timeout

                            elif action == "Wait for Pixel Color" and obj_type == "pixel":
                                expected_rgb = tuple(params.get("expected_rgb", target_object.get("rgb"))) # Use stored or param
                                if not expected_rgb:
                                    print(f"    ERROR: No RGB color defined for pixel '{obj_name}' or wait parameters.")
                                    continue

                                timeout = params.get("timeout_s", 10)
                                start_time = time.time()
                                found_color = False
                                print(f"    Waiting for pixel '{obj_name}' at {target_object['coords']} to be RGB {expected_rgb} (timeout: {timeout}s)")
                                while time.time() - start_time < timeout:
                                    current_rgb = pyautogui.pixel(target_object['coords'][0], target_object['coords'][1])
                                    if current_rgb == expected_rgb:
                                        print(f"    Pixel color matched.")
                                        found_color = True
                                        break
                                    time.sleep(0.5)
                                if not found_color:
                                    print(f"    TIMEOUT: Pixel color not matched after {timeout}s. Last color: {current_rgb}")

                        # --- GLOBAL ACTIONS (Object may or may not be used) ---
                        if action == "Wait":
                            duration = params.get("duration_s", 1.0)
                            min_dur = params.get("min_s")
                            max_dur = params.get("max_s")
                            if min_dur is not None and max_dur is not None:
                                wait_time = random.uniform(min_dur, max_dur)
                                print(f"    Random Wait: {wait_time:.2f}s")
                                time.sleep(wait_time)
                            else:
                                print(f"    Static Wait: {duration}s")
                                time.sleep(duration)

                        elif action == "Keyboard Input":
                            text = params.get("text_to_type", "")
                            keys_to_press = params.get("keys_to_press") # e.g. "enter", ["ctrl", "c"]
                            if text:
                                pyautogui.typewrite(text, interval=0.05) # Small interval for reliability
                            if keys_to_press:
                                if isinstance(keys_to_press, list):
                                    pyautogui.hotkey(*keys_to_press)
                                else:
                                    pyautogui.press(keys_to_press)

                        # Add more actions here...

                    except pyautogui.FailSafeException:
                        print("!!! FAILSAFE TRIGGERED (mouse to top-left corner) !!!")
                        self.root.deiconify()
                        return # Stop all execution
                    except Exception as e:
                        print(f"    ERROR executing step {i+1}: {e}")
                        # Add option to stop sequence on error
                        # For now, continue to next step or loop
                        # If critical error, might need to break loop

                if not is_infinite and current_loop >= loops:
                    break # Exit while loop

        except KeyboardInterrupt:
            print("\n--- Execution Interrupted by User (Ctrl+C) ---")
        finally:
            self.root.deiconify() # Show main window again
            print(f"--- Sequence Finished: {self.current_sequence_name} ---")


# --- UI Frame Classes ---
class BaseFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.configure(bg="#E0E0E0") # Light grey background for frames

class MainFrame(BaseFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)

        tk.Label(self, text="Automation Tool", font=("Arial", 16, "bold"), bg=self["bg"]).pack(pady=20)

        btn_frame_new = tk.Frame(self, bg=self["bg"])
        btn_frame_new.pack(pady=10, fill="x", padx=20)

        # Add a "New Sequence" button
        tk.Button(btn_frame_new, text="New Sequence", width=20, command=controller.new_sequence).pack(pady=5, fill="x")
        tk.Button(btn_frame_new, text="Instructions", width=20, command=lambda: controller.show_frame("InstructionsFrame")).pack(pady=5, fill="x")

        tk.Label(btn_frame_new, text="Create & Edit", font=("Arial", 12, "underline"), bg=self["bg"]).pack(pady=(10,0)) # Renamed section
        tk.Button(btn_frame_new, text="Object Creation", width=20, command=lambda: controller.show_frame("ObjectCreationFrame")).pack(pady=5, fill="x")
        tk.Button(btn_frame_new, text="Step Creator", width=20, command=lambda: controller.show_frame("StepCreatorFrame")).pack(pady=5, fill="x")

        sep = ttk.Separator(self, orient="horizontal")
        sep.pack(fill="x", pady=15, padx=20)

        tk.Label(self, text="File Operations", font=("Arial", 12, "underline"), bg=self["bg"]).pack(pady=(10,0)) # Renamed section

        file_ops_frame = tk.Frame(self, bg=self["bg"]) # New frame for better layout
        file_ops_frame.pack(pady=5, padx=20, fill="x")

        tk.Button(file_ops_frame, text="Load Sequence", command=controller.load_sequence).pack(side=tk.LEFT, expand=True, fill="x", padx=(0,5))
        # Add Save and Save As buttons here if desired, or keep them in StepCreator
        # For now, keep Save in StepCreator, add Save As here for convenience
        tk.Button(file_ops_frame, text="Save Sequence As...", command=controller.save_sequence_as).pack(side=tk.LEFT, expand=True, fill="x", padx=(5,0))

        loaded_seq_frame = tk.Frame(self, bg=self["bg"]) # Frame for label
        loaded_seq_frame.pack(pady=5, padx=20, fill="x")
        tk.Label(loaded_seq_frame, text="Current:", bg=self["bg"]).pack(side=tk.LEFT)
        self.loaded_seq_label = tk.Label(loaded_seq_frame, text="No sequence loaded", bg="white", relief=tk.SUNKEN, anchor="w")
        self.loaded_seq_label.pack(side=tk.LEFT, expand=True, fill="x")

        loop_frame = tk.Frame(self, bg=self["bg"])
        loop_frame.pack(pady=10, padx=20, fill="x")
        tk.Label(loop_frame, text="Loops (0 for indefinite):", bg=self["bg"]).pack(side=tk.LEFT)
        tk.Entry(loop_frame, textvariable=controller.loop_count, width=5, justify="center").pack(side=tk.LEFT, padx=5)

        tk.Button(self, text="Run Sequence", width=20, font=("Arial", 12, "bold"), bg="#A5D6A7", command=controller.run_sequence).pack(pady=20, padx=20, fill="x")

    def refresh_content(self):
        if self.controller.current_sequence_name == DEFAULT_PROJECT_NAME and not self.controller.current_project_path:
            self.loaded_seq_label.config(text="No sequence loaded")
        else:
            self.loaded_seq_label.config(text=self.controller.current_sequence_name)

class ObjectCreationFrame(BaseFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.current_creation_type = "region" # To know if grid/drag is for region or image

        tk.Label(self, text="Object Creation", font=("Arial", 16, "bold"), bg=self["bg"]).pack(pady=10)

        # --- Region Creation ---
        region_frame = tk.LabelFrame(self, text="Region Creation", padx=10, pady=10, bg=self["bg"])
        region_frame.pack(pady=10, padx=10, fill="x")

        grid_dim_frame = tk.Frame(region_frame, bg=self["bg"])
        grid_dim_frame.pack(fill="x")
        tk.Label(grid_dim_frame, text="Grid (W x H):", bg=self["bg"]).pack(side=tk.LEFT, padx=(0,5))
        tk.Entry(grid_dim_frame, textvariable=controller.grid_cols_var, width=3).pack(side=tk.LEFT)
        tk.Label(grid_dim_frame, text="x", bg=self["bg"]).pack(side=tk.LEFT)
        tk.Entry(grid_dim_frame, textvariable=controller.grid_rows_var, width=3).pack(side=tk.LEFT, padx=(0,10))

        tk.Button(grid_dim_frame, text="Grid Mode", command=lambda: self.set_creation_type_and_run("region", controller.create_region_grid_mode)).pack(side=tk.LEFT, expand=True, fill="x")
        tk.Button(region_frame, text="Drag Mode", command=lambda: self.set_creation_type_and_run("region", controller.create_region_drag_mode)).pack(pady=5, fill="x")
        tk.Button(region_frame, text="Pixel Monitor", command=controller._start_pixel_monitor_listener).pack(pady=5, fill="x")

        # --- Image Creation ---
        image_frame = tk.LabelFrame(self, text="Image Creation", padx=10, pady=10, bg=self["bg"])
        image_frame.pack(pady=10, padx=10, fill="x")
        # Grid Dim for image capture shares the same variables for simplicity
        tk.Button(image_frame, text="Grid Mode (Capture)", command=lambda: self.set_creation_type_and_run("image", controller.create_region_grid_mode)).pack(pady=5, fill="x")
        tk.Button(image_frame, text="Drag Mode (Capture)", command=lambda: self.set_creation_type_and_run("image", controller.create_region_drag_mode)).pack(pady=5, fill="x")

        # --- Sound Creation (Placeholder) ---
        sound_frame = tk.LabelFrame(self, text="Sound Creation", padx=10, pady=10, bg=self["bg"])
        sound_frame.pack(pady=10, padx=10, fill="x")
        tk.Button(sound_frame, text="Sound Recording (Future)", state=tk.DISABLED).pack(pady=5, fill="x")

        # --- View All Objects ---
        self.objects_list_frame = tk.LabelFrame(self, text="Created Objects", padx=10, pady=10, bg=self["bg"])
        self.objects_list_frame.pack(pady=10, padx=10, fill="both", expand=True)
        self.objects_text = scrolledtext.ScrolledText(self.objects_list_frame, height=5, wrap=tk.WORD, state=tk.DISABLED)
        self.objects_text.pack(fill="both", expand=True)
        # tk.Button(self, text="View All Objects", command=self.view_all_objects).pack(pady=5, padx=10, fill="x")

        tk.Button(self, text="Back to Main Menu", command=lambda: controller.show_frame("MainFrame")).pack(pady=15, side=tk.BOTTOM)

    def set_creation_type_and_run(self, c_type, func_to_run):
        self.current_creation_type = c_type
        func_to_run()

    def update_objects_display(self):
        self.objects_text.config(state=tk.NORMAL)
        self.objects_text.delete(1.0, tk.END)
        if not self.controller.objects:
            self.objects_text.insert(tk.END, "No objects created yet.")
        else:
            for name, data in self.controller.objects.items():
                obj_type = data.get("type", "N/A")
                details = ""
                if obj_type == "region" or obj_type == "image":
                    details = f"Coords: {data.get('coords')}"
                    if data.get('image_path'): details += f", Path: {os.path.basename(data['image_path'])}"
                elif obj_type == "pixel":
                    details = f"Coords: {data.get('coords')}, RGB: {data.get('rgb')}"

                self.objects_text.insert(tk.END, f"- {name} ({obj_type.capitalize()}): {details}\n")
        self.objects_text.config(state=tk.DISABLED)

    def refresh_content(self):
        self.update_objects_display()


class StepCreatorFrame(BaseFrame):
    ACTION_CONFIG = {
        "region": ["Click", "Wait for Image in Region (Future)", "Type into Region (Future)"],
        "pixel": ["Click", "Wait for Pixel Color"],
        "image": ["Click", "Wait for Image"], # Click = click located image
        "_global_": ["Wait", "Keyboard Input"] # Actions not tied to an object type
    }

    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.step_widgets = [] # To keep track of rows of widgets

        tk.Label(self, text="Step Creator", font=("Arial", 16, "bold"), bg=self["bg"]).pack(pady=10)

        # Frame for the list of steps, with a scrollbar
        self.canvas_steps = tk.Canvas(self, borderwidth=0, background="#ffffff")
        self.steps_area_frame = tk.Frame(self.canvas_steps, background="#ffffff") # Frame to hold actual step rows
        self.scrollbar_steps = tk.Scrollbar(self, orient="vertical", command=self.canvas_steps.yview)
        self.canvas_steps.configure(yscrollcommand=self.scrollbar_steps.set)

        self.scrollbar_steps.pack(side="right", fill="y", padx=(0,5), pady=5)
        self.canvas_steps.pack(side="left", fill="both", expand=True, padx=(5,0), pady=5)
        self.canvas_steps.create_window((0,0), window=self.steps_area_frame, anchor="nw", tags="self.steps_area_frame")

        self.steps_area_frame.bind("<Configure>", lambda e: self.canvas_steps.configure(scrollregion=self.canvas_steps.bbox("all")))
        self.canvas_steps.bind_all("<MouseWheel>", self._on_mousewheel) # For Windows/Linux mouse wheel scrolling

        self.add_step_row() # Start with one empty step

        btn_row = tk.Frame(self, bg=self["bg"])
        btn_row.pack(pady=10, fill="x", padx=10)
        tk.Button(btn_row, text="+ Add Step", command=self.add_step_row).pack(
            side=tk.LEFT, padx=5, pady=5, fill="x", expand=True
        )
        # Update Save Sequence button
        tk.Button(btn_row, text="Save Sequence", command=self.controller.save_sequence, state=tk.NORMAL).pack(
            side=tk.LEFT, padx=5, pady=5, fill="x", expand=True
        )

        tk.Button(self, text="Back to Main Menu", command=lambda: controller.show_frame("MainFrame")).pack(pady=15, side=tk.BOTTOM)

    def _on_mousewheel(self, event):
        # Check if mouse is over the canvas_steps before scrolling
        if self.canvas_steps.winfo_containing(event.x_root, event.y_root) == self.canvas_steps:
            self.canvas_steps.yview_scroll(int(-1*(event.delta/120)), "units")


    def add_step_row(self, step_data=None):
        row_frame = tk.Frame(self.steps_area_frame, bg="#F0F0F0", pady=2) # Light background for each step row
        row_frame.pack(fill="x", expand=True)

        step_num = len(self.step_widgets) + 1
        tk.Label(row_frame, text=f"{step_num}.", width=3, bg=row_frame["bg"]).pack(side=tk.LEFT, padx=(5,0))

        # Object Dropdown
        obj_var = tk.StringVar()
        obj_dropdown = ttk.Combobox(row_frame, textvariable=obj_var, values=self.controller.get_object_names(), width=15, state="readonly")
        obj_dropdown.pack(side=tk.LEFT, padx=5)

        # Action Dropdown
        action_var = tk.StringVar()
        action_dropdown = ttk.Combobox(row_frame, textvariable=action_var, width=20, state="readonly")
        action_dropdown.pack(side=tk.LEFT, padx=5)

        # Params Button (opens dialog to set params)
        params_button = tk.Button(row_frame, text="Params", width=8,
                                  command=lambda ov=obj_var, av=action_var, idx=len(self.step_widgets): self.edit_step_params(ov, av, idx))
        params_button.pack(side=tk.LEFT, padx=5)

        # Delete Button
        del_button = tk.Button(row_frame, text="X", fg="red", width=2,
                               command=lambda rf=row_frame, idx=len(self.step_widgets): self.delete_step_row(rf, idx))
        del_button.pack(side=tk.LEFT, padx=5)

        # Store widgets and associated data
        current_step_entry = {
            "frame": row_frame, "obj_var": obj_var, "obj_dropdown": obj_dropdown,
            "action_var": action_var, "action_dropdown": action_dropdown,
            "params": {}, "params_button": params_button, "delete_button": del_button
        }
        self.step_widgets.append(current_step_entry)

        # Link object selection to action population
        obj_var.trace_add("write", lambda *args, ov=obj_var, ad=action_dropdown: self.update_action_dropdown(ov, ad))

        if step_data: # If loading a step
            obj_var.set(step_data.get("object_name", ""))
            action_var.set(step_data.get("action", ""))
            current_step_entry["params"] = step_data.get("params", {})
        else: # New step, populate initial actions if "Global" applicable
            self.update_action_dropdown(obj_var, action_dropdown) # Populate global actions if obj is empty
            self.controller.mark_sequence_modified() # Mark as modified for new steps only

        self._renumber_steps()

    def delete_step_row(self, row_frame, index_to_delete):
        row_frame.destroy()
        # Important: Remove from self.step_widgets and self.controller.current_steps
        # This needs careful index management if not just popping last one.
        # For simplicity, let's assume deletion is complex and requires rebuilding or careful splicing.
        # A simpler way: mark as "deleted" and filter out later, or rebuild list.

        # For now, let's just remove from widgets. The actual data in controller.current_steps
        # should be updated when "Save Sequence" is pressed or structure is finalized.
        # This is a simplification.

        # Find the entry to remove from self.step_widgets
        removed_widget_entry = None
        for i, entry in enumerate(self.step_widgets):
            if entry["frame"] == row_frame:
                removed_widget_entry = self.step_widgets.pop(i)
                break

        # If also managing controller.current_steps directly:
        if index_to_delete < len(self.controller.current_steps): # Ensure index is valid
            del self.controller.current_steps[index_to_delete]

        self.controller.mark_sequence_modified() # Mark as modified

        self._renumber_steps()
        self.steps_area_frame.update_idletasks() # Crucial for scrollbar to update
        self.canvas_steps.config(scrollregion=self.canvas_steps.bbox("all"))


    def _renumber_steps(self):
        for i, widgets in enumerate(self.step_widgets):
            # Find the label in the frame's children
            for child in widgets["frame"].winfo_children():
                if isinstance(child, tk.Label) and child.cget("text").endswith("."):
                    child.config(text=f"{i+1}.")
                    break

    def finalize_steps_for_controller(self):
        """Collect all steps from UI and update controller's current_steps."""
        steps = []
        for step_widget in self.step_widgets:
            obj_var = step_widget["obj_var"]
            action_var = step_widget["action_var"]
            params = step_widget["params"]

            # Only include steps that have both an object and action selected
            if obj_var.get() or action_var.get():  # Allow global actions with no object
                step_data = {
                    "object_name": obj_var.get() if obj_var.get() != "Global" else None,
                    "action": action_var.get(),
                    "params": params
                }
                steps.append(step_data)

        self.controller.current_steps = steps
        print(f"Finalized steps for controller: {steps}")
        return steps

    def clear_and_rebuild_steps(self, steps_data_list_from_file):
        """Clear all step rows and rebuild from the provided steps data."""
        # Clear existing UI rows
        for widget_entry in self.step_widgets:
            if widget_entry["frame"].winfo_exists():
                widget_entry["frame"].destroy()
        self.step_widgets = []

        # Set the controller's steps directly from the loaded file data
        self.controller.current_steps = list(steps_data_list_from_file) if steps_data_list_from_file else []

        # Now, build the UI based on this newly set self.controller.current_steps
        for step_data in self.controller.current_steps:
            self.add_step_row(step_data) # add_step_row uses step_data to set UI if provided

        # If no steps were loaded (or in the file), ensure at least one empty UI row
        if not self.step_widgets:
            self.add_step_row() # This adds an empty UI row; finalize will pick it up if it's filled

        self.steps_area_frame.update_idletasks()
        self.canvas_steps.config(scrollregion=self.canvas_steps.bbox("all"))
        self.controller.mark_sequence_modified(False) # Loading a sequence makes it "unmodified" initially

    def update_action_dropdown(self, obj_var, action_dropdown):
        selected_obj_name = obj_var.get()
        actions = []
        if selected_obj_name:
            obj_data = self.controller.objects.get(selected_obj_name)
            if obj_data:
                obj_type = obj_data.get("type")
                actions.extend(self.ACTION_CONFIG.get(obj_type, []))

        # Always add global actions, or only if no object type selected?
        # For now, let's make global actions always available, maybe with a prefix.
        # Or better: if object is selected, show its actions. If no object, show globals.
        # The mockup has object always there. So, global actions are just general purpose.

        current_actions = list(self.ACTION_CONFIG["_global_"]) # Start with global
        if selected_obj_name:
            obj_data = self.controller.objects.get(selected_obj_name)
            if obj_data:
                obj_type = obj_data.get("type")
                obj_specific_actions = self.ACTION_CONFIG.get(obj_type, [])
                # Prepend object-specific, then global, avoid duplicates
                current_actions = obj_specific_actions + [ga for ga in current_actions if ga not in obj_specific_actions]


        action_dropdown['values'] = current_actions
        if current_actions:
            # Keep current action if valid, else set to first, or clear
            current_action_val = action_dropdown.get()
            if current_action_val not in current_actions:
                 action_dropdown.set(current_actions[0] if current_actions else "")
        else:
            action_dropdown.set("")


    def edit_step_params(self, obj_var, action_var, step_idx):
        # This should open a dialog based on the selected action to get its parameters
        action = action_var.get()
        obj_name = obj_var.get()
        current_params = self.step_widgets[step_idx].get("params", {})

        # Example for "Wait" action
        if action == "Wait":
            dialog = WaitParamsDialog(self, "Wait Action Parameters", current_params)
            if dialog.result and dialog.result != current_params: # Check if params actually changed
                self.step_widgets[step_idx]["params"] = dialog.result
                self.controller.mark_sequence_modified()
                print(f"Params for step {step_idx+1} (Wait): {dialog.result}")
        elif action == "Keyboard Input":
            dialog = KeyboardParamsDialog(self, "Keyboard Input Parameters", current_params)
            if dialog.result:
                self.step_widgets[step_idx]["params"] = dialog.result
                self.controller.mark_sequence_modified()
                print(f"Params for step {step_idx+1} (Keyboard): {dialog.result}")
        elif action == "Wait for Pixel Color":
            # Needs the pixel object to suggest current color
            target_pixel_obj = self.controller.objects.get(obj_name)
            dialog = PixelColorWaitParamsDialog(self, "Pixel Color Wait Parameters", current_params, target_pixel_obj)
            if dialog.result:
                self.step_widgets[step_idx]["params"] = dialog.result
                self.controller.mark_sequence_modified()
        elif action == "Wait for Image":
            # Needs image object for confidence etc.
            target_image_obj = self.controller.objects.get(obj_name)
            dialog = ImageWaitParamsDialog(self, "Image Wait Parameters", current_params, target_image_obj)
            if dialog.result:
                self.step_widgets[step_idx]["params"] = dialog.result
                self.controller.mark_sequence_modified()
        else:
            simpledialog.messagebox.showinfo("Parameters", f"No specific parameters for '{action}'.\nObject: {obj_name}", parent=self)
            self.step_widgets[step_idx]["params"] = {} # Clear params if none defined

    def refresh_content(self):
        # Rebuild step rows if objects changed significantly, or just update dropdowns
        # For now, just ensure object dropdowns are fresh if they exist.
        # A full rebuild is safer if objects can be deleted/renamed.
        all_obj_names = self.controller.get_object_names()
        for step_widget_entry in self.step_widgets:
            current_obj_val = step_widget_entry["obj_var"].get()
            step_widget_entry["obj_dropdown"]['values'] = all_obj_names
            if current_obj_val not in all_obj_names and current_obj_val != "":
                step_widget_entry["obj_var"].set("") # Clear if object no longer exists
            # Trigger action dropdown update
            self.update_action_dropdown(step_widget_entry["obj_var"], step_widget_entry["action_dropdown"])

        # If self.controller.current_steps is the source of truth, rebuild UI from it
        # This is safer if loading sequences.
        # For now, this basic refresh is okay for interactive building.

    def refresh_object_dropdowns(self):
        """Called when objects are added/deleted to update all object dropdowns."""
        all_obj_names = self.controller.get_object_names()
        for step_entry in self.step_widgets:
            current_selection = step_entry["obj_var"].get()
            step_entry["obj_dropdown"]["values"] = all_obj_names
            if current_selection not in all_obj_names: # If previously selected object is gone
                step_entry["obj_var"].set("") # Clear it, will trigger action_dropdown update
            else: # Trigger action update even if object selection is same, in case object def changed
                self.update_action_dropdown(step_entry["obj_var"], step_entry["action_dropdown"])

    def finalize_steps_for_controller(self):
        """Updates controller.current_steps from the UI before saving or running."""
        self.controller.current_steps = []
        for sw in self.step_widgets:
            obj_name = sw["obj_var"].get()
            action = sw["action_var"].get()
            params = sw.get("params", {})
            if action : # Only add steps that have an action selected
                self.controller.current_steps.append({
                    "object_name": obj_name if obj_name else None, # Store None if no object selected
                    "action": action,
                    "params": params
                })
        print("Finalized steps for controller:", self.controller.current_steps)


class InstructionsFrame(BaseFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        tk.Label(self, text="Instructions", font=("Arial", 16, "bold"), bg=self["bg"]).pack(pady=10)

        instructions_text = """
Welcome to the Python Desktop Automation Tool!

1.  **Object Creation**:
    *   Go to "Object Creation" from the Main Menu.
    *   **Region**: Define a rectangular area on screen.
        *   Grid Mode: Set grid size, click "Grid Mode", select cells, confirm.
        *   Drag Mode: Click "Drag Mode", click and drag on screen.
    *   **Pixel**: Define a single point and its color.
        *   Pixel Monitor: Click, then click "Capture Pixel..." button, move mouse to target, click it.
    *   **Image**: Capture a part of the screen as an image.
        *   Grid/Drag Mode (Capture): Similar to region, but captures the selection as an image file.
    *   All objects must be named uniquely.

2.  **Step Creator**:
    *   Go to "Step Creator" from the Main Menu.
    *   Click "+ Add Step" to add a new action to your sequence.
    *   For each step:
        *   Select an **Object** from the dropdown (objects you created).
        *   Select an **Action** to perform (e.g., Click, Wait, Type).
        *   Click "Params" to set specific details for the action (e.g., wait duration, text to type).
    *   You can delete steps using the "X" button.

3.  **Running a Sequence**:
    *   From the Main Menu:
        *   Set the number of **Loops** (0 for infinite).
        *   Click "Run Sequence".
    *   To stop an infinitely looping sequence or any sequence early, quickly move your mouse to the top-left corner of the screen (PyAutoGUI's default failsafe).

4.  **Saving/Loading (Future)**:
    *   Functionality to save your created objects and steps into a sequence file, and load them back, will be added.

Tips:
*   Name your objects descriptively (e.g., "LoginButton", "UsernameField").
*   Test individual steps or small sequences frequently.
"""
        text_area = scrolledtext.ScrolledText(self, wrap=tk.WORD, height=20, width=60, font=("Arial", 10))
        text_area.insert(tk.INSERT, instructions_text)
        text_area.config(state=tk.DISABLED, bg=self["bg"]) # Read-only
        text_area.pack(pady=10, padx=10, fill="both", expand=True)

        tk.Button(self, text="Back to Main Menu", command=lambda: controller.show_frame("MainFrame")).pack(pady=15)


# --- Parameter Dialogs for Steps ---
class BaseParamsDialog(simpledialog.Dialog):
    def __init__(self, parent, title, existing_params=None):
        self.existing_params = existing_params if existing_params else {}
        self.result = None
        super().__init__(parent, title)

    def body(self, master):
        # Override in subclasses
        pass

    def apply(self):
        # Override in subclasses to populate self.result
        pass

class WaitParamsDialog(BaseParamsDialog):
    def body(self, master):
        self.wait_type_var = tk.StringVar(value=self.existing_params.get("type", "static"))
        self.duration_var = tk.StringVar(value=str(self.existing_params.get("duration_s", 1.0)))
        self.min_dur_var = tk.StringVar(value=str(self.existing_params.get("min_s", 0.5)))
        self.max_dur_var = tk.StringVar(value=str(self.existing_params.get("max_s", 2.0)))

        tk.Radiobutton(master, text="Static Wait (seconds):", variable=self.wait_type_var, value="static", command=self.toggle_fields).grid(row=0, column=0, sticky="w")
        self.static_entry = tk.Entry(master, textvariable=self.duration_var, width=10)
        self.static_entry.grid(row=0, column=1)

        tk.Radiobutton(master, text="Random Wait (min-max seconds):", variable=self.wait_type_var, value="random", command=self.toggle_fields).grid(row=1, column=0, sticky="w")
        self.min_entry = tk.Entry(master, textvariable=self.min_dur_var, width=5)
        self.min_entry.grid(row=1, column=1, sticky="w")
        tk.Label(master, text="to").grid(row=1, column=1) # Poor man's layout
        self.max_entry = tk.Entry(master, textvariable=self.max_dur_var, width=5)
        self.max_entry.grid(row=1, column=1, sticky="e")

        self.toggle_fields()
        return self.static_entry # initial focus

    def toggle_fields(self):
        if self.wait_type_var.get() == "static":
            self.static_entry.config(state="normal")
            self.min_entry.config(state="disabled")
            self.max_entry.config(state="disabled")
        else: # random
            self.static_entry.config(state="disabled")
            self.min_entry.config(state="normal")
            self.max_entry.config(state="normal")

    def validate(self):
        try:
            if self.wait_type_var.get() == "static":
                val = float(self.duration_var.get())
                if val < 0: raise ValueError("Duration must be non-negative.")
            else:
                min_val = float(self.min_dur_var.get())
                max_val = float(self.max_dur_var.get())
                if min_val < 0 or max_val < 0: raise ValueError("Durations must be non-negative.")
                if min_val > max_val: raise ValueError("Min duration cannot exceed Max duration.")
            return 1
        except ValueError as e:
            simpledialog.messagebox.showerror("Invalid Input", str(e), parent=self)
            return 0

    def apply(self):
        if self.wait_type_var.get() == "static":
            self.result = {"type": "static", "duration_s": float(self.duration_var.get())}
        else:
            self.result = {"type": "random", "min_s": float(self.min_dur_var.get()), "max_s": float(self.max_dur_var.get())}

class KeyboardParamsDialog(BaseParamsDialog):
    def body(self, master):
        tk.Label(master, text="Text to type:").grid(row=0, column=0, sticky="w")
        self.text_var = tk.StringVar(value=self.existing_params.get("text_to_type", ""))
        self.text_entry = tk.Entry(master, textvariable=self.text_var, width=30)
        self.text_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(master, text="Or Special Keys (e.g., enter, ctrl,c):").grid(row=1, column=0, sticky="w")
        self.keys_var = tk.StringVar(value=", ".join(self.existing_params.get("keys_to_press", [])) if isinstance(self.existing_params.get("keys_to_press"), list) else self.existing_params.get("keys_to_press", ""))
        self.keys_entry = tk.Entry(master, textvariable=self.keys_var, width=30)
        self.keys_entry.grid(row=1, column=1, padx=5, pady=5)
        tk.Label(master, text="(Separate multiple keys with comma for hotkeys)").grid(row=2, columnspan=2, sticky="w", padx=5)

        return self.text_entry

    def apply(self):
        self.result = {}
        text_val = self.text_var.get()
        keys_val_str = self.keys_var.get()

        if text_val:
            self.result["text_to_type"] = text_val

        if keys_val_str:
            # If comma detected, treat as list for hotkey, else single key
            if ',' in keys_val_str:
                self.result["keys_to_press"] = [k.strip() for k in keys_val_str.split(',')]
            else:
                self.result["keys_to_press"] = keys_val_str.strip()


class PixelColorWaitParamsDialog(BaseParamsDialog):
    def __init__(self, parent, title, existing_params=None, pixel_obj=None):
        self.pixel_obj = pixel_obj
        super().__init__(parent, title, existing_params)

    def body(self, master):
        default_rgb_str = ""
        if self.existing_params.get("expected_rgb"):
            default_rgb_str = ",".join(map(str, self.existing_params["expected_rgb"]))
        elif self.pixel_obj and self.pixel_obj.get("rgb"):
            default_rgb_str = ",".join(map(str, self.pixel_obj["rgb"]))

        tk.Label(master, text="Expected RGB (R,G,B):").grid(row=0, column=0, sticky="w")
        self.rgb_var = tk.StringVar(value=default_rgb_str)
        self.rgb_entry = tk.Entry(master, textvariable=self.rgb_var, width=15)
        self.rgb_entry.grid(row=0, column=1)
        tk.Button(master, text="Pick Color", command=self.pick_color).grid(row=0, column=2, padx=5)

        tk.Label(master, text="Timeout (seconds):").grid(row=1, column=0, sticky="w")
        self.timeout_var = tk.StringVar(value=str(self.existing_params.get("timeout_s", 10.0)))
        self.timeout_entry = tk.Entry(master, textvariable=self.timeout_var, width=10)
        self.timeout_entry.grid(row=1, column=1)
        return self.rgb_entry

    def pick_color(self):
        # colorchooser returns ((r,g,b), '#rrggbb') or (None, None)
        initial_color_hex = None
        try:
            r,g,b = map(int, self.rgb_var.get().split(','))
            initial_color_hex = f"#{r:02x}{g:02x}{b:02x}"
        except: pass

        color_code = colorchooser.askcolor(title="Choose Expected Pixel Color", initialcolor=initial_color_hex, parent=self)
        if color_code and color_code[0]: # Check if a color was chosen (color_code[0] is RGB tuple)
            r, g, b = map(int, color_code[0]) # askcolor returns float 0-255 for RGB
            self.rgb_var.set(f"{r},{g},{b}")


    def validate(self):
        try:
            rgb_parts = [int(p.strip()) for p in self.rgb_var.get().split(',')]
            if len(rgb_parts) != 3 or not all(0 <= val <= 255 for val in rgb_parts):
                raise ValueError("RGB must be three numbers between 0-255, comma-separated.")
            timeout = float(self.timeout_var.get())
            if timeout < 0: raise ValueError("Timeout must be non-negative.")
            return 1
        except ValueError as e:
            simpledialog.messagebox.showerror("Invalid Input", str(e), parent=self)
            return 0

    def apply(self):
        rgb_parts = [int(p.strip()) for p in self.rgb_var.get().split(',')]
        self.result = {
            "expected_rgb": tuple(rgb_parts),
            "timeout_s": float(self.timeout_var.get())
        }

class ImageWaitParamsDialog(BaseParamsDialog):
    def __init__(self, parent, title, existing_params=None, image_obj=None):
        self.image_obj = image_obj # To get default confidence
        super().__init__(parent, title, existing_params)

    def body(self, master):
        default_confidence = 0.8
        if self.existing_params.get("confidence"):
            default_confidence = self.existing_params["confidence"]
        elif self.image_obj and self.image_obj.get("confidence"):
            default_confidence = self.image_obj["confidence"]

        tk.Label(master, text="Confidence (0.0-1.0):").grid(row=0, column=0, sticky="w")
        self.confidence_var = tk.StringVar(value=str(default_confidence))
        self.confidence_entry = tk.Entry(master, textvariable=self.confidence_var, width=10)
        self.confidence_entry.grid(row=0, column=1)

        tk.Label(master, text="Timeout (seconds):").grid(row=1, column=0, sticky="w")
        self.timeout_var = tk.StringVar(value=str(self.existing_params.get("timeout_s", 10.0)))
        self.timeout_entry = tk.Entry(master, textvariable=self.timeout_var, width=10)
        self.timeout_entry.grid(row=1, column=1)
        return self.confidence_entry

    def validate(self):
        try:
            confidence = float(self.confidence_var.get())
            if not (0.0 <= confidence <= 1.0):
                raise ValueError("Confidence must be between 0.0 and 1.0.")
            timeout = float(self.timeout_var.get())
            if timeout < 0: raise ValueError("Timeout must be non-negative.")
            return 1
        except ValueError as e:
            simpledialog.messagebox.showerror("Invalid Input", str(e), parent=self)
            return 0

    def apply(self):
        self.result = {
            "confidence": float(self.confidence_var.get()),
            "timeout_s": float(self.timeout_var.get())
        }


# --- Main Execution ---
if __name__ == "__main__":
    app_root = tk.Tk()
    # Set a theme if available for a more modern look (optional)
    try:
        style = ttk.Style(app_root)
        # Available themes: 'winnative', 'clam', 'alt', 'default', 'classic', 'vista', 'xpnative'
        if 'clam' in style.theme_names(): # Clam is widely available
            style.theme_use('clam')
        elif 'vista' in style.theme_names(): # Good on Windows
             style.theme_use('vista')
    except tk.TclError:
        print("TTK themes not available or fa=0iled to apply.")

    app_instance = DesktopAutomationApp(app_root)

    # Before mainloop, finalize steps for controller if there are any initial steps in UI
    # This is more for loading sequences, but good practice
    if hasattr(app_instance.frames["StepCreatorFrame"], 'finalize_steps_for_controller'):
        app_instance.frames["StepCreatorFrame"].finalize_steps_for_controller()

    app_root.mainloop()