import sys
import os
import subprocess
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QLineEdit, QPushButton,
                             QFileDialog, QCheckBox, QFrame, QMessageBox)
from PyQt6.QtCore import Qt, QSettings
from PyQt6.QtGui import QPixmap, QDragEnterEvent, QDropEvent
from PIL import Image
import win32com.client


class FileDropWidget(QFrame):
    """A custom widget that combines a QLineEdit, buttons, and drag-and-drop functionality."""
    def __init__(self, placeholder_text, browse_filter, show_clear_button=False, callback=None):
        super().__init__()
        self.callback = callback
        self.browse_filter = browse_filter
        
        self.setAcceptDrops(True)
        self.setObjectName("fileDropWidget")
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(6)
        
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText(placeholder_text)
        self.path_edit.setReadOnly(True)
        self.path_edit.setFixedHeight(42)
        layout.addWidget(self.path_edit)
        
        browse_btn = QPushButton("Browse")
        browse_btn.setObjectName("browseButton")
        browse_btn.setFixedHeight(42)
        browse_btn.setFixedWidth(100)
        browse_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        browse_btn.clicked.connect(self.browse)
        layout.addWidget(browse_btn)
        
        if show_clear_button:
            self.clear_btn = QPushButton("Clear")
            self.clear_btn.setObjectName("clearButton")
            self.clear_btn.setFixedHeight(42)
            self.clear_btn.setFixedWidth(80)
            self.clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            self.clear_btn.clicked.connect(self.clear)
            layout.addWidget(self.clear_btn)
        else:
            self.clear_btn = None
            
        self.default_style = ""
        self.hover_style = ""

    def set_styles(self, default, hover):
        self.default_style = default
        self.hover_style = hover
        self.setStyleSheet(self.default_style)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet(self.hover_style)
            
    def dragLeaveEvent(self, event):
        self.setStyleSheet(self.default_style)
        
    def dropEvent(self, event: QDropEvent):
        self.setStyleSheet(self.default_style)
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files and self.callback:
            self.callback(files[0])

    def browse(self):
        filepath, _ = QFileDialog.getOpenFileName(self, self.path_edit.placeholderText(), "", self.browse_filter)
        if filepath and self.callback:
            self.callback(filepath)
            
    def clear(self):
        if self.callback:
            self.callback("")

    def setText(self, text):
        self.path_edit.setText(text)

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Python Taskbar Pinner")
        self.setFixedSize(600, 720)
        
        self.settings = QSettings("TaskbarPinner", "App")
        self.dark_mode = self.settings.value("dark_mode", True, type=bool)
        
        self.script_path = ""
        self.image_path = ""
        
        self.init_ui()
        self.apply_theme()
        
    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # --- HEADER ---
        header = QWidget()
        header.setObjectName("header")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(30, 15, 30, 15)
        
        title = QLabel("Python Taskbar Pinner")
        title.setObjectName("title")
        header_layout.addWidget(title)
        header_layout.addStretch()
        
        self.theme_btn = QPushButton("☀️" if self.dark_mode else "🌙")
        self.theme_btn.setObjectName("themeButton")
        self.theme_btn.setFixedSize(40, 40)
        self.theme_btn.clicked.connect(self.toggle_theme)
        self.theme_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        header_layout.addWidget(self.theme_btn)
        
        main_layout.addWidget(header)
        
        # --- CONTENT AREA ---
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(30, 20, 30, 30)
        content_layout.setSpacing(7)
        
        # Script Selection
        script_label = QLabel("Select Python Script")
        script_label.setObjectName("sectionLabel")
        content_layout.addWidget(script_label)
        self.script_selector = FileDropWidget(
            "Drag & drop a .py file or click Browse", "Python Files (*.py *.pyw)", callback=self.handle_script_selection)
        content_layout.addWidget(self.script_selector)
        
        # Image Selection
        image_label = QLabel("Select Icon Image (Optional)")
        image_label.setObjectName("sectionLabel")
        content_layout.addWidget(image_label)
        self.image_selector = FileDropWidget(
            "Drag & drop an image or click Browse", "Image Files (*.png *.jpg *.ico)", True, self.handle_image_selection)
        content_layout.addWidget(self.image_selector)
        
        # Shortcut Name
        name_label = QLabel("Shortcut Name (Optional)")
        name_label.setObjectName("sectionLabel")
        content_layout.addWidget(name_label)
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Leave blank to use script's name")
        self.name_input.setFixedHeight(42)
        content_layout.addWidget(self.name_input)
        
        content_layout.addSpacing(20)
        
        # Icon Preview
        preview_label = QLabel("Icon Preview")
        preview_label.setObjectName("previewLabel")
        preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        content_layout.addWidget(preview_label, 0, Qt.AlignmentFlag.AlignCenter)

        self.preview_frame = QLabel()
        self.preview_frame.setObjectName("preview")
        self.preview_frame.setFixedSize(140, 140)
        self.preview_frame.setAlignment(Qt.AlignmentFlag.AlignCenter)
        content_layout.addWidget(self.preview_frame, 0, Qt.AlignmentFlag.AlignCenter)
        
        content_layout.addSpacing(20)

        # Hide Console Checkbox
        self.hide_console_cb = QCheckBox("Hide console window (recommended for GUI apps)")
        self.hide_console_cb.setCursor(Qt.CursorShape.PointingHandCursor)
        content_layout.addWidget(self.hide_console_cb, 0, Qt.AlignmentFlag.AlignCenter)
        
        content_layout.addStretch(1)
        create_btn = QPushButton("Create Shortcut")
        create_btn.setObjectName("createButton")
        create_btn.clicked.connect(self.create_shortcut)
        create_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        create_btn.setFixedHeight(50)
        content_layout.addWidget(create_btn)
        
        main_layout.addWidget(content)
        
        # Status Bar
        self.status_bar = QLabel("Ready. Select your Python script.")
        self.status_bar.setObjectName("statusBar")
        self.status_bar.setFixedHeight(45)
        self.status_bar.setAlignment(Qt.AlignmentFlag.AlignVCenter)
        self.status_bar.setContentsMargins(30, 0, 30, 0)
        main_layout.addWidget(self.status_bar)
        
        self.update_preview()

    def apply_theme(self):
        if self.dark_mode:
            # --- Dark Theme Colors ---
            bg_main = "#1a1a1a"
            bg_secondary = "#252525"
            bg_input = "#2a2a2a"
            text_primary = "#e8e8e8"
            text_secondary = "#a0a0a0"
            border = "#3a3a3a"
            accent = "#4a9eff"
            btn_gray = "#3f3f3f"
            btn_gray_hover = "#4a4a4a"
            red = "#c94040"
            red_hover = "#b33030"
            
            drop_default = f"background-color: {bg_input}; border: 1px solid {border}; border-radius: 8px;"
            drop_hover = f"background-color: {bg_input}; border: 1px solid {accent}; border-radius: 8px;"
            
            stylesheet = f"""
                QMainWindow {{ background-color: {bg_main}; }}
                QWidget {{ color: {text_primary}; font-family: 'Segoe UI', system-ui, sans-serif; font-size: 10pt; }}
                #header {{ background-color: {bg_secondary}; border-bottom: 1px solid {border}; }}
                #title {{ font-size: 14pt; font-weight: 500; }}
                QPushButton#themeButton {{ background-color: {bg_input}; border: 1px solid {border}; border-radius: 8px; font-size: 16pt; padding: 0; }}
                QPushButton#themeButton:hover {{ background-color: {btn_gray_hover}; }}
                #sectionLabel {{ font-size: 11pt; font-weight: 600; margin-bottom: 4px; }}
                #fileDropWidget {{ {drop_default} }}
                #fileDropWidget QLineEdit {{ background-color: transparent; border: none; padding-left: 10px; }}
                QLineEdit {{ background-color: {bg_input}; border: 1px solid {border}; border-radius: 6px; padding: 0 15px; }}
                QLineEdit:focus {{ border: 1px solid {accent}; }}
                QPushButton {{ background-color: {bg_input}; border: 1px solid {border}; border-radius: 6px; font-weight: 500; }}
                QPushButton#browseButton {{ background-color: {btn_gray}; }}
                QPushButton#browseButton:hover {{ background-color: {btn_gray_hover}; }}
                QPushButton#clearButton {{ background-color: {red}; border-color: {red}; color: white; }}
                QPushButton#clearButton:hover {{ background-color: {red_hover}; border-color: {red_hover}; }}
                QPushButton#createButton {{ background-color: {accent}; border: none; color: #ffffff; font-size: 11pt; font-weight: 600; border-radius: 8px; }}
                QPushButton#createButton:hover {{ background-color: #3a8eef; }}
                #previewLabel {{ color: {text_secondary}; font-weight: 600; margin-bottom: 5px; }}
                #preview {{ background-color: {bg_input}; border: 1px solid {border}; border-radius: 8px; color: {text_secondary}; font-size: 9pt; }}
                QCheckBox {{ spacing: 10px; }}
                QCheckBox::indicator {{ width: 20px; height: 20px; border-radius: 4px; border: 2px solid {border}; background-color: {bg_input}; }}
                QCheckBox::indicator:hover {{ border-color: {accent}; }}
                QCheckBox::indicator:checked {{ background-color: {accent}; border-color: {accent}; image: url(data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTEzLjMzMzMgNEw2IDExLjMzMzNMMi42NjY2NyA4IiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCIvPgo8L3N2Zz4K); }}
                #statusBar {{ background-color: {bg_secondary}; border-top: 1px solid {border}; color: {text_secondary}; font-size: 9pt; }}
            """
        else:
            # --- Light Theme Colors ---
            bg_main = "#ffffff"
            bg_secondary = "#f5f5f5"
            bg_input = "#ffffff"
            text_primary = "#1a1a1a"
            text_secondary = "#666666"
            border = "#e0e0e0"
            accent = "#2563eb"
            btn_gray = "#f0f0f0"
            btn_gray_hover = "#e2e8f0"
            red = "#e53e3e"
            red_hover = "#c53030"

            drop_default = f"background-color: {bg_input}; border: 1px solid {border}; border-radius: 8px;"
            drop_hover = f"background-color: {bg_input}; border: 1px solid {accent}; border-radius: 8px;"
            
            stylesheet = f"""
                QMainWindow {{ background-color: {bg_main}; }}
                QWidget {{ color: {text_primary}; font-family: 'Segoe UI', system-ui, sans-serif; font-size: 10pt; }}
                #header {{ background-color: {bg_secondary}; border-bottom: 1px solid {border}; }}
                #title {{ font-size: 14pt; font-weight: 500; }}
                QPushButton#themeButton {{ background-color: {bg_input}; border: 1px solid {border}; border-radius: 8px; font-size: 16pt; padding: 0; }}
                QPushButton#themeButton:hover {{ background-color: {btn_gray_hover}; }}
                #sectionLabel {{ font-size: 11pt; font-weight: 600; margin-bottom: 4px; }}
                #fileDropWidget {{ {drop_default} }}
                #fileDropWidget QLineEdit {{ background-color: transparent; border: none; padding-left: 10px; }}
                QLineEdit {{ background-color: {bg_input}; border: 1px solid {border}; border-radius: 6px; padding: 0 15px; }}
                QLineEdit:focus {{ border: 1px solid {accent}; }}
                QPushButton {{ background-color: {bg_input}; border: 1px solid {border}; border-radius: 6px; font-weight: 500; }}
                QPushButton#browseButton {{ background-color: {btn_gray}; }}
                QPushButton#browseButton:hover {{ background-color: {btn_gray_hover}; }}
                QPushButton#clearButton {{ background-color: {red}; border-color: {red}; color: black; }}
                QPushButton#clearButton:hover {{ background-color: {red_hover}; border-color: {red_hover}; }}
                QPushButton#createButton {{ background-color: {accent}; border: none; color: #ffffff; font-size: 11pt; font-weight: 600; border-radius: 8px; }}
                QPushButton#createButton:hover {{ background-color: #1d4ed8; }}
                #previewLabel {{ color: {text_secondary}; font-weight: 600; margin-bottom: 5px; }}
                #preview {{ background-color: {bg_secondary}; border: 1px solid {border}; border-radius: 8px; color: {text_secondary}; font-size: 9pt; }}
                QCheckBox {{ spacing: 10px; }}
                QCheckBox::indicator {{ width: 20px; height: 20px; border-radius: 4px; border: 2px solid {border}; background-color: {bg_input}; }}
                QCheckBox::indicator:hover {{ border-color: {accent}; }}
                QCheckBox::indicator:checked {{ background-color: {accent}; border-color: {accent}; image: url(data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTEzLjMzMzMgNEw2IDExLjMzMzNMMi42NjY2NyA4IiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCIvPgo8L3N2Zz4K); }}
                #statusBar {{ background-color: {bg_secondary}; border-top: 1px solid {border}; color: {text_secondary}; font-size: 9pt; }}
            """
        
        self.setStyleSheet(stylesheet)
        self.script_selector.set_styles(drop_default, drop_hover)
        self.image_selector.set_styles(drop_default, drop_hover)
        
    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self.settings.setValue("dark_mode", self.dark_mode)
        self.theme_btn.setText("☀️" if self.dark_mode else "🌙")
        self.apply_theme()
        self.update_preview()
        
    def handle_script_selection(self, filepath):
        if not filepath: return
        
        if filepath.endswith(('.py', '.pyw')):
            self.script_path = filepath
            self.script_selector.setText(filepath)
            self.update_status("Python script selected.")
        else:
            self.update_status("Please select a .py or .pyw file.", is_error=True)

    def handle_image_selection(self, filepath):
        if filepath == "":
            self.image_path = ""
            self.image_selector.setText("")
            self.update_preview()
            self.update_status("Image cleared. Will use default icon.")
            return

        if filepath.lower().endswith(('.png', '.jpg', '.jpeg', '.ico', '.bmp')):
            self.image_path = filepath
            self.image_selector.setText(filepath)
            self.update_preview()
            self.update_status("Icon image selected.")
        else:
            self.update_status("Please select a valid image file.", is_error=True)
            
    def update_preview(self):
        if not self.image_path:
            self.preview_frame.setPixmap(QPixmap())
            self.preview_frame.setText("Icon Preview\n(Defaults to Python icon)")
            return
            
        try:
            pixmap = QPixmap(self.image_path)
            if not pixmap.isNull():
                scaled = pixmap.scaled(120, 120, Qt.AspectRatioMode.KeepAspectRatio, 
                                      Qt.TransformationMode.SmoothTransformation)
                self.preview_frame.setPixmap(scaled)
                self.preview_frame.setText("")
            else:
                self.preview_frame.setPixmap(QPixmap())
                self.preview_frame.setText("Invalid\nImage")
        except Exception:
            self.preview_frame.setPixmap(QPixmap())
            self.preview_frame.setText("Invalid\nImage")
            
    def update_status(self, message, is_error=False):
        self.status_bar.setText(message)
        base_style = self.status_bar.styleSheet()
        if is_error:
            color = "#ff6b6b" if self.dark_mode else "#d32f2f"
            self.status_bar.setStyleSheet(f"{base_style} color: {color};")
        else:
            self.apply_theme()
            self.status_bar.setText(message)
            
    def create_shortcut(self):
        script = self.script_path
        image = self.image_path
        
        if not script:
            QMessageBox.critical(self, "Error", "Please select a Python script.")
            return
            
        if not os.path.exists(script):
            QMessageBox.critical(self, "Error", "The selected Python script does not exist.")
            return
            
        try:
            script_dir = os.path.dirname(script)
            
            custom_name = self.name_input.text().strip()
            shortcut_name = custom_name if custom_name else os.path.splitext(os.path.basename(script))[0]
                
            python_exe_name = 'pythonw.exe' if self.hide_console_cb.isChecked() else 'python.exe'
            python_exe_path = os.path.join(os.path.dirname(sys.executable), python_exe_name)
            if not os.path.exists(python_exe_path):
                python_exe_path = sys.executable
                
            icon_location = python_exe_path
            
            if image:
                if not os.path.exists(image):
                    QMessageBox.warning(self, "Warning", "The selected image file does not exist. Using default icon.")
                else:
                    icon_path = os.path.join(script_dir, f"{shortcut_name}_icon.ico")
                    try:
                        img = Image.open(image)
                        img.save(icon_path, format='ICO', sizes=[(32,32), (48,48), (64,64), (256,256)])
                        icon_location = icon_path
                    except Exception as e:
                        QMessageBox.warning(self, "Warning", f"Could not convert image: {e}. Using Python's default icon.")
                        icon_location = python_exe_path
                
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            shortcut_path = os.path.join(desktop_path, f"{shortcut_name}.lnk")
            
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(shortcut_path)
            shortcut.TargetPath = python_exe_path
            shortcut.Arguments = f'"{script}"'
            shortcut.WorkingDirectory = script_dir
            shortcut.IconLocation = icon_location
            shortcut.save()
            
            self.update_status("Shortcut created successfully!")
            
            QMessageBox.information(self, "Success!",
                "Shortcut created on your Desktop.\nYou can now right-click it and choose 'Pin to taskbar'.")
            
            subprocess.run(f'explorer /select,"{shortcut_path}"', shell=True)
            
        except Exception as e:
            self.update_status(f"Error: {e}", is_error=True)
            QMessageBox.critical(self, "Error", f"Failed to create shortcut:\n{e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())
