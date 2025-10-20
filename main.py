import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
import sys
import subprocess
from PIL import Image, ImageTk
import win32com.client

class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Python Taskbar Pinner")
        # MODIFIED: Increased height for the new name field
        self.geometry("550x600")
        self.resizable(False, False)

        # Member Variables
        self.script_path = tk.StringVar()
        self.image_path = tk.StringVar()
        # NEW: StringVar to hold the custom shortcut name
        self.shortcut_name_var = tk.StringVar()
        self.hide_console = tk.BooleanVar(value=True)
        self.image_preview = None

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Python Script Selection ---
        script_frame = self.create_file_drop_frame(main_frame, "1. Drag & Drop Python Script (.py) or Click to Select", self.select_script, self.script_path, has_clear=False)
        script_frame.pack(fill=tk.X, pady=5)
        
        # --- Image Selection ---
        image_frame = self.create_file_drop_frame(main_frame, "2. Drag & Drop Icon Image (Optional)", self.select_image, self.image_path, has_clear=True)
        image_frame.pack(fill=tk.X, pady=5)
        
        # NEW: Shortcut Name Section ---
        name_frame = ttk.LabelFrame(main_frame, text="3. Shortcut Name (Optional)", padding="10")
        name_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(name_frame, text="Leave blank to use the script's name.").pack(anchor="w", pady=(0, 5))
        ttk.Entry(name_frame, textvariable=self.shortcut_name_var).pack(fill=tk.X, expand=True)
        
        # --- Image Preview ---
        self.preview_canvas = tk.Canvas(main_frame, width=128, height=128, bg="white", relief="groove")
        self.preview_canvas.pack(pady=10)
        self.preview_canvas.create_text(64, 64, text="Icon Preview\n(Defaults to Python icon)", fill="grey", justify=tk.CENTER)

        # --- Options ---
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.pack(fill=tk.X, pady=10)
        
        ttk.Checkbutton(
            options_frame, text="Hide Console Window (Recommended for GUI apps)", variable=self.hide_console
        ).pack(anchor="w")

        # --- Create Button ---
        create_button = ttk.Button(main_frame, text="Create Shortcut and Show Me", command=self.create_shortcut)
        create_button.pack(pady=20, ipady=5)

        # --- Status Bar ---
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor="w", padding="2 5")
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.update_status("Ready. Select your Python script.")

    def create_file_drop_frame(self, parent, label_text, command, string_var, has_clear=False):
        frame = ttk.LabelFrame(parent, text=label_text, padding="10")
        entry = ttk.Entry(frame, textvariable=string_var, state="readonly")
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        button_browse = ttk.Button(frame, text="...", command=command, width=4)
        button_browse.pack(side=tk.LEFT, padx=(0, 5))

        if has_clear:
            button_clear = ttk.Button(frame, text="Clear", command=self.clear_image, width=6)
            button_clear.pack(side=tk.LEFT)

        frame.drop_target_register(DND_FILES)
        entry.drop_target_register(DND_FILES)
        frame.dnd_bind('<<Drop>>', lambda e: self.handle_drop(e, string_var))
        entry.dnd_bind('<<Drop>>', lambda e: self.handle_drop(e, string_var))
        return frame
        
    def handle_drop(self, event, target_var):
        filepath = event.data.strip('{}')
        if os.path.exists(filepath):
            target_var.set(filepath)
            if target_var == self.image_path: self.update_image_preview()

    def update_status(self, message, is_error=False):
        self.status_var.set(message)
        self.status_bar.config(foreground="red" if is_error else "black")

    def select_script(self):
        path = filedialog.askopenfilename(title="Select Python Script", filetypes=[("Python Files", "*.py *.pyw")])
        if path: self.script_path.set(path); self.update_status("Python script selected.")

    def select_image(self):
        path = filedialog.askopenfilename(title="Select Icon Image", filetypes=[("Image Files", "*.png *.jpg *.jpeg *.ico")])
        if path: self.image_path.set(path); self.update_image_preview(); self.update_status("Icon image selected.")
    
    def clear_image(self):
        self.image_path.set("")
        self.update_image_preview()
        self.update_status("Image cleared. Will use default Python icon.")

    def update_image_preview(self):
        self.preview_canvas.delete("all")
        path = self.image_path.get()
        if not path:
            self.preview_canvas.create_text(64, 64, text="Icon Preview\n(Defaults to Python icon)", fill="grey", justify=tk.CENTER)
            return
        try:
            img = Image.open(path)
            img.thumbnail((128, 128))
            self.image_preview = ImageTk.PhotoImage(img)
            self.preview_canvas.create_image(64, 64, image=self.image_preview)
        except Exception:
            self.preview_canvas.create_text(64, 64, text="Invalid Image", fill="red")

    # MODIFIED: Core logic updated to handle custom shortcut name
    def create_shortcut(self):
        script = self.script_path.get()
        image = self.image_path.get()

        # 1. Validation
        if not script:
            messagebox.showerror("Error", "Please select a Python script.")
            return
        if not os.path.exists(script):
            messagebox.showerror("Error", "The selected Python script does not exist.")
            return

        try:
            script_dir = os.path.dirname(script)
            
            # 2. Determine Shortcut Name
            # NEW: Check if the user provided a custom name, otherwise derive from script
            custom_name = self.shortcut_name_var.get().strip()
            if custom_name:
                shortcut_name = custom_name
            else:
                script_filename = os.path.basename(script)
                shortcut_name = os.path.splitext(script_filename)[0]
            
            # 3. Determine Python Executable
            python_exe_name = 'pythonw.exe' if self.hide_console.get() else 'python.exe'
            python_exe_path = os.path.join(os.path.dirname(sys.executable), python_exe_name)
            if not os.path.exists(python_exe_path): python_exe_path = sys.executable

            icon_location = python_exe_path # Default to python executable for icon

            # 4. Process image ONLY if one was provided
            if image:
                if not os.path.exists(image):
                    messagebox.showerror("Error", "The selected image file does not exist.")
                    return
                # Convert Image to .ico using the final shortcut name
                icon_path = os.path.join(script_dir, f"{shortcut_name}_icon.ico")
                img = Image.open(image)
                img.save(icon_path, format='ICO', sizes=[(32,32), (48,48), (64,64), (256,256)])
                icon_location = icon_path

            # 5. Create the Shortcut
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            shortcut_path = os.path.join(desktop_path, f"{shortcut_name}.lnk")

            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(shortcut_path)
            shortcut.TargetPath = python_exe_path
            shortcut.Arguments = f'"{script}"'
            shortcut.WorkingDirectory = script_dir
            shortcut.IconLocation = icon_location
            shortcut.save()
            
            self.update_status("Shortcut created! Showing you the file now...")
            
            messagebox.showinfo(
                "Success! Last Step...",
                "A File Explorer window will now open with your new shortcut highlighted.\n\n"
                "To finish, just right-click it and choose 'Pin to taskbar'."
            )
            
            subprocess.run(f'explorer /select,"{shortcut_path}"', shell=True)

        except Exception as e:
            self.update_status(f"An error occurred: {e}", is_error=True)
            messagebox.showerror("Error", f"Failed to create shortcut:\n{e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()