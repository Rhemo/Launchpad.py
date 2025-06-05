import sys
import subprocess
import shutil
import re
import webbrowser
import time
import json
import os
import os.path
import importlib.util
import importlib
import platform
import winreg
import psutil
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QPushButton, QGridLayout, QWidget, 
                            QLabel, QLineEdit, QListWidget, QVBoxLayout, QTextEdit, QInputDialog, 
                            QMessageBox, QListWidgetItem, QDialog, QFormLayout, QComboBox, QCheckBox, 
                            QDialogButtonBox, QHBoxLayout, QFileDialog, QSpacerItem, QSizePolicy, 
                            QAction, QMenu, QToolBar, QTextBrowser, QToolButton, QColorDialog)
from PyQt5.QtCore import Qt, QSize, pyqtSignal, QTimer, QEvent
from PyQt5.QtGui import QTextCharFormat, QFont, QTextCursor, QTextListFormat, QColor, QSyntaxHighlighter

class CommandHighlighter(QSyntaxHighlighter):
    def __init__(self, parent, default_commands, custom_commands):
        super().__init__(parent)
        self.default_commands = default_commands
        self.custom_commands = custom_commands
        # Command format (blue)
        self.command_format = QTextCharFormat()
        self.command_format.setForeground(QColor("#1e90ff"))  # Blue for commands
        self.command_format.setFontWeight(QFont.Bold)
        # Placeholder format (dark green)
        self.placeholder_format = QTextCharFormat()
        self.placeholder_format.setForeground(QColor("#008000"))  # Dark Green for placeholders
        self.placeholder_format.setFontWeight(QFont.Bold)

    def highlightBlock(self, text):
        # Highlight commands
        commands = self.default_commands + self.custom_commands
        for command in commands:
            index = text.lower().find(command.lower())
            while index >= 0:
                length = len(command)
                if (index == 0 or not text[index-1].isalnum()) and (index + length >= len(text) or not text[index + length].isalnum()):
                    self.setFormat(index, length, self.command_format)
                index = text.lower().find(command.lower(), index + length)
        # Highlight placeholders like %username%, <domain>, <whid> (case-insensitive)
        placeholder_pattern = r'<%[a-zA-Z_]+%>|<[a-zA-Z_]+(?:\|validate_service)?>'  # Matches %username%, <domain>, <whid>, <WHID>
        for match in re.finditer(placeholder_pattern, text, re.IGNORECASE):
            start, end = match.span()
            self.setFormat(start, end - start, self.placeholder_format)

def install_required_module(module_name):
    """Install a required Python module using pip."""
    print(f"Attempting to install {module_name}...")
    try:
        result = subprocess.check_call([sys.executable, '-m', 'pip', 'install', module_name])
        print(f"Installation of {module_name} completed with return code: {result}")
    except subprocess.CalledProcessError as e:
        print(f"Failed to install {module_name}: {e}")
        raise

def verify_module(module_name, import_names, test_attrs):
    """Verify module installation in a fresh interpreter."""
    script = (
        "import importlib\n"
        "importlib.invalidate_caches()\n"
        "try:\n"
        f"    {'; '.join(f'module_{i} = importlib.import_module(\"{name}\")' for i, name in enumerate(import_names))}\n"
        f"    {'; '.join(f'getattr(module_{i}, \"{attr}\")' for i, attr in enumerate(test_attrs))}\n"
        "    print('SUCCESS')\n"
        "except (ImportError, AttributeError) as e:\n"
        "    print('FAIL: ' + str(e))"
    )
    try:
        result = subprocess.check_output([sys.executable, '-c', script], text=True, stderr=subprocess.STDOUT)
        if 'SUCCESS' in result:
            print(f"{module_name} successfully installed and verified.")
            return True
        else:
            print(f"Verification failed: {result.strip()}")
            return False
    except subprocess.CalledProcessError as e:
        print(f"Error verifying {module_name}: {e.output}")
        return False

def check_and_install_modules():
    """
    Check for required modules and install if missing or broken.
    Currently checks for: PyQt5 and its dependencies
    """
    required_modules = {
        'PyQt5': ('PyQt5', ['PyQt5.QtWidgets', 'PyQt5.QtCore'], ['QApplication', 'Qt']),
        'pywin32': ('pywin32', ['win32com.client'], ['Dispatch']),
        'psutil': ('psutil', ['psutil'], ['Process'])
    }
    
    print("Verifying installations...")
    try:
        for module_name, (pip_name, import_names, test_attrs) in required_modules.items():
            print(f"Checking {module_name}...")
            for import_name, test_attr in zip(import_names, test_attrs):
                if import_name in sys.modules:
                    del sys.modules[import_name]
                module = importlib.import_module(import_name)
                getattr(module, test_attr)
            print(f"{module_name} is installed and functional.")
        print("All required modules are verified.")
        return
    except (ImportError, AttributeError) as e:
        print(f"Module verification failed: {str(e)}")
        print("Attempting repair...")
        try:
            print("Cleaning up old installations...")
            packages_to_remove = ['PyQt5', 'PyQt5-Qt5', 'PyQt5-sip', 'PyQtWebEngine']
            for package in packages_to_remove:
                try:
                    subprocess.check_call([sys.executable, '-m', 'pip', 'uninstall', '-y', package])
                except:
                    print(f"Note: {package} was not installed")
            print("Purging pip cache...")
            subprocess.check_call([sys.executable, '-m', 'pip', 'cache', 'purge'])
            print("Installing required packages...")
            install_required_module('PyQt5-sip')
            install_required_module('PyQt5')
            print("Verifying new installation...")
            for module_name, (pip_name, import_names, test_attrs) in required_modules.items():
                if not verify_module(module_name, import_names, test_attrs):
                    raise Exception(f"Verification of {module_name} failed after installation")
            print("All required modules successfully installed and verified.")
            print("Starting new process to apply module changes...")
            subprocess.run([sys.executable] + sys.argv)
            sys.exit(0)
        except Exception as install_error:
            print(f"Error during repair: {str(install_error)}")
            if sys.version_info >= (3, 13):
                print("Python 3.13 detected; PyQt5 may not be compatible. Try Python 3.12 or run again.")
            sys.exit(1)

check_and_install_modules()

def get_installed_browsers():
    """Detect installed browsers and the default browser."""
    browsers = []
    default_browser = None
    browser_paths = {}

    if platform.system() == "Windows":
        browser_checks = [
            ("Chrome", r"Google\Chrome\Application\chrome.exe"),
            ("Firefox", r"Mozilla Firefox\firefox.exe"),
            ("Edge", r"Microsoft\Edge\Application\msedge.exe"),
            ("Opera", r"Opera\launcher.exe"),
            ("Safari", r"Safari\Safari.exe")
        ]
        for name, path in browser_checks:
            for base_dir in [os.getenv("ProgramFiles"), os.getenv("ProgramFiles(x86)"), os.getenv("LocalAppData")]:
                if base_dir:
                    full_path = os.path.join(base_dir, path)
                    if os.path.exists(full_path):
                        browsers.append(name)
                        browser_paths[name] = full_path
                        break

        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice") as key:
                prog_id = winreg.QueryValueEx(key, "ProgId")[0].lower()
                if any(s in prog_id for s in ["chrome", "googlechrome"]):
                    default_browser = "Chrome"
                elif any(s in prog_id for s in ["firefox", "mozilla"]):
                    default_browser = "Firefox"
                elif any(s in prog_id for s in ["edge", "msedge", "microsoft-edge"]):
                    default_browser = "Edge"
                elif "opera" in prog_id:
                    default_browser = "Opera"
                elif "safari" in prog_id:
                    default_browser = "Safari"
        except Exception:
            pass

        if not default_browser:
            try:
                default_browser_name = webbrowser.get().name.lower()
                if any(s in default_browser_name for s in ["chrome", "google-chrome"]):
                    default_browser = "Chrome"
                elif any(s in default_browser_name for s in ["firefox", "mozilla"]):
                    default_browser = "Firefox"
                elif any(s in default_browser_name for s in ["edge", "msedge", "microsoft-edge"]):
                    default_browser = "Edge"
                elif "opera" in default_browser_name:
                    default_browser = "Opera"
                elif "safari" in default_browser_name:
                    default_browser = "Safari"
            except Exception:
                pass

        if not default_browser and browsers:
            default_browser = browsers[0]

    if default_browser and default_browser in browsers:
        browsers = [default_browser] + [b for b in browsers if b != default_browser]
    elif default_browser and default_browser not in browsers:
        browsers.insert(0, default_browser)
        for name, path in browser_checks:
            if name == default_browser:
                for base_dir in [os.getenv("ProgramFiles"), os.getenv("ProgramFiles(x86)"), os.getenv("LocalAppData")]:
                    if base_dir:
                        full_path = os.path.join(base_dir, path)
                        if os.path.exists(full_path):
                            browser_paths[name] = full_path
                            break
                break

    if not browsers and default_browser:
        browsers = [default_browser]
    
    return browsers, browser_paths

class ClickableLabel(QLabel):
    clicked = pyqtSignal()
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setStyleSheet("""
            QLabel {
                color: white;
                font-size: 9px;
                font-weight: bold;
                background-color: rgba(0, 0, 0, 50);
                border: 1px solid #777;
                padding: 2px;
            }
        """)
        self.setFixedSize(32, 20)
        self.setAlignment(Qt.AlignCenter)
    def mousePressEvent(self, event):
        self.clicked.emit()
        event.accept()

class EditLinkDialog(QDialog):
    def __init__(self, parent=None, link=None, is_new=True):
        super().__init__(parent)
        self.setWindowTitle("Add Link" if is_new else "Edit Link")
        self.link = link
        self.is_new = is_new
        link_key = self.parent().get_link_settings_key(self.link["name"]) if self.link else None
        link_settings = self.parent().user_preferences["link_settings"].get(link_key, {}) if link_key else {}
        self.selected_color = QColor(link_settings.get("color", "#0078d4"))
        self.is_setting_custom = False
        self.color_map = {
            "Default": "#0078d4", "Sapphire": "#0d47a1", "Deep Purple": "#7b1fa2",
            "Crimson": "#d81b60", "Action Green": "#28a745", "Amber": "#f57c00",
            "Slate Blue": "#546e7a", "Olive": "#827717", "Coral": "#f06292",
            "Teal": "#00838f", "Indigo": "#283593"
        }
        self.setup_ui()

    def setup_ui(self):
        layout = QFormLayout()
        self.name_input = QLineEdit(self.link["name"] if self.link else "")
        self.name_input.setPlaceholderText("e.g., My Portal")
        self.name_input.setMaxLength(30)
        layout.addRow("Button Name:", self.name_input)
        self.url_input = QLineEdit(self.link["url"] if self.link else "")
        self.url_input.setPlaceholderText("e.g., https://example.com or outlook.exe")
        layout.addRow("URL or Command:", self.url_input)
        self.tooltip_input = QLineEdit(self.link["tooltip"] if self.link else "")
        self.tooltip_input.setPlaceholderText("e.g., Open My Portal")
        layout.addRow("Tooltip:", self.tooltip_input)

        # Color selection dropdown
        self.color_combo = QComboBox()
        self.color_combo.addItems([
            "Default", "Sapphire", "Deep Purple", "Crimson", "Action Green",
            "Amber", "Slate Blue", "Olive", "Coral", "Teal", "Indigo", "Custom"
        ])
        if self.link:
            link_color = self.selected_color.name().lower()
            found = False
            for name, hex_code in self.color_map.items():
                if hex_code.lower() == link_color:
                    self.color_combo.setCurrentText(name)
                    found = True
                    break
            if not found:
                self.color_combo.setCurrentText("Custom")
        else:
            self.color_combo.setCurrentText("Default")
        layout.addRow("Button Color:", self.color_combo)

        # Color picker button
        self.color_picker_button = QPushButton("Pick Custom Color")
        self.color_picker_button.clicked.connect(self.pick_color)
        self.update_color_picker_button_style()
        layout.addRow(self.color_picker_button)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        if not self.is_new:
            delete_btn = buttons.addButton("Delete", QDialogButtonBox.DestructiveRole)
            delete_btn.clicked.connect(self.delete_link)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)

        # Connect dropdown change to update selected color
        self.color_combo.currentTextChanged.connect(self.on_color_combo_changed)

    def on_color_combo_changed(self, color_name):
        if color_name == "Custom" and not self.is_setting_custom:
            self.pick_color()
        elif color_name != "Custom":  # Only update if not Custom
            self.selected_color = QColor(self.color_map.get(color_name, "#0078d4"))
            self.update_color_picker_button_style()

    def pick_color(self):
        color = QColorDialog.getColor(self.selected_color, self, "Select Custom Color")
        if color.isValid():
            self.is_setting_custom = True
            self.selected_color = color
            self.color_combo.setCurrentText("Custom")
            self.update_color_picker_button_style()
            self.color_picker_button.repaint()  # Force UI update
            self.is_setting_custom = False

    def update_color_picker_button_style(self):
        hex_color = self.selected_color.name()
        self.color_picker_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {hex_color};
                color: white;
                padding: 5px;
                border: 1px solid #555;
                border-radius: 3px;
            }}
            QPushButton:hover {{ background-color: {self.darken_color(hex_color, 0.8)}; }}
        """)

    def darken_color(self, hex_color, factor=0.8):
        hex_color = hex_color.lstrip('#')
        rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        darkened = tuple(int(c * factor) for c in rgb)
        return f"#{darkened[0]:02x}{darkened[1]:02x}{darkened[2]:02x}"

    def validate_and_accept(self):
        name = self.name_input.text().strip()
        url = self.url_input.text().strip()
        tooltip = self.tooltip_input.text().strip()
        if not name or not url:
            QMessageBox.warning(self, "Input Error", "Name and URL/command are required.")
            return
        self.link_data = {
            "name": name,
            "url": url,
            "tooltip": tooltip or f"Open {name}",
            "icon": self.link["icon"] if self.link else "icon-default.png",
        }
        link_key = self.parent().get_link_settings_key(name)
        link_settings = self.parent().user_preferences["link_settings"].get(link_key, {})
        link_settings["color"] = self.selected_color.name()
        link_settings["favorite"] = link_settings.get("favorite", False)
        self.parent().user_preferences["link_settings"][link_key] = link_settings
        try:
            self.parent().save_user_preferences()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to save preferences: {e}")
            return
        self.accept()

    def delete_link(self):
        if QMessageBox.question(self, "Confirm Delete", f"Are you sure you want to delete '{self.link['name']}'?") == QMessageBox.Yes:
            # Remove user preferences for this link using mode-specific key
            link_key = self.parent().get_link_settings_key(self.link["name"])
            if link_key in self.parent().user_preferences["link_settings"]:
                del self.parent().user_preferences["link_settings"][link_key]
                self.parent().save_user_preferences()
            self.link_data = None
            self.accept()

    def get_data(self):
        return self.link_data

class SettingsDialog(QDialog):
    def __init__(self, current_settings, parent=None):
        super().__init__(parent)
        self.setWindowTitle("JSON Storage Settings")
        self.current_settings = current_settings
        self.setup_ui()

    def setup_ui(self):
        layout = QFormLayout()
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["Local", "Shared"])
        self.mode_combo.setCurrentText(self.current_settings.get("mode", "Local"))
        layout.addRow("Storage Mode:", self.mode_combo)
        self.path_input = QLineEdit(self.current_settings.get("network_path", ""))
        self.path_input.setPlaceholderText("e.g., \\\\server\\share\\path")
        layout.addRow("Network Share Path:", self.path_input)
        self.oncall_email_input = QLineEdit(self.current_settings.get("oncall_email", "page-ots-oncall-nbwi1@amazon.com"))
        self.oncall_email_input.setPlaceholderText("e.g., oncall@team.com")
        layout.addRow("On-Call Email:", self.oncall_email_input)
        self.clear_cache_button = QPushButton("Clear Help Cache")
        self.clear_cache_button.clicked.connect(self.clear_help_cache)
        layout.addRow(self.clear_cache_button)
        self.path_input.setEnabled(self.mode_combo.currentText() == "Shared")
        self.mode_combo.currentTextChanged.connect(self.toggle_path_input)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def toggle_path_input(self, mode):
        self.path_input.setEnabled(mode == "Shared")

    def clear_help_cache(self):
        for dialog in QApplication.topLevelWidgets():
            if isinstance(dialog, NewCommandDialog):
                dialog.cmdlet_syntax_cache.clear()
                dialog.help_repo = dialog.load_help_repository()
        QMessageBox.information(self, "Success", "Help cache cleared.")

    def validate_and_accept(self):
        mode = self.mode_combo.currentText()
        oncall_email = self.oncall_email_input.text().strip()
        if not oncall_email:
            QMessageBox.warning(self, "Input Error", "On-Call Email is required.")
            return
        if mode == "Shared":
            network_path = self.path_input.text().strip()
            if not network_path:
                QMessageBox.warning(self, "Input Error", "Network share path is required in Shared mode.")
                return
            try:
                if not os.path.exists(network_path):
                    QMessageBox.warning(self, "Input Error", f"Network path '{network_path}' does not exist or is inaccessible.")
                    return
            except Exception as e:
                QMessageBox.warning(self, "Input Error", f"Error accessing network path: {str(e)}")
                return
            self.settings = {"mode": "Shared", "network_path": network_path, "oncall_email": oncall_email}
        else:
            self.settings = {"mode": "Local", "network_path": "", "oncall_email": oncall_email}
        self.accept()

    def get_settings(self):
        return self.settings
        
class PageOnCallDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Page OnCall")
        self.setup_ui()
    
    def setup_ui(self):
        layout = QFormLayout()
        default_email = self.parent().settings.get("oncall_email", "page-ots-oncall-nbwi1@amazon.com")
        self.to_input = QLineEdit(default_email)
        self.to_input.setPlaceholderText("e.g., oncall@team.com")
        layout.addRow("To:", self.to_input)
        self.subject_input = QLineEdit()
        self.subject_input.setPlaceholderText("e.g., Urgent OnCall Request")
        layout.addRow("Subject:", self.subject_input)
        self.body_input = QTextEdit()
        self.body_input.setPlaceholderText("Enter additional details for the on-call team...")
        layout.addRow("Message:", self.body_input)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)
    
    def validate_and_accept(self):
        to = self.to_input.text().strip()
        subject = self.subject_input.text().strip()
        body = self.body_input.toPlainText().strip()
        if not to or not subject or not body:
            QMessageBox.warning(self, "Input Error", "To, Subject, and Message are required.")
            return
        self.email_data = {
            "to": to,
            "subject": subject,
            "body": body
        }
        self.accept()
    
    def get_email_data(self):
        return self.email_data

class MultiInputDialog(QDialog):
    def __init__(self, placeholders, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Enter Placeholder Values")
        self.placeholders = placeholders
        layout = QFormLayout()
        self.inputs = {}
        for ph in placeholders:
            display_name = ph.split("|")[0]
            input_field = QLineEdit()
            layout.addRow(f"Enter value for {display_name}:", input_field)
            self.inputs[ph] = input_field
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def validate_and_accept(self):
        for ph, input_field in self.inputs.items():
            value = input_field.text().strip()
            display_name = ph.split("|")[0]
            if not value:
                QMessageBox.warning(self, "Input Error", f"Value for {display_name} is required.")
                return
            if ph.endswith("|validate_service"):
                try:
                    result = subprocess.run(["powershell", "-Command", f"Get-Service -Name '{value}'"], capture_output=True, text=True, check=True)
                except subprocess.CalledProcessError:
                    QMessageBox.warning(self, "Validation Error", f"Invalid service name: {value}")
                    return
        self.accept()

    def get_values(self):
        return {ph.split("|")[0]: self.inputs[ph].text().strip() for ph in self.inputs}

class LinkButtonWidget(QWidget):
    def __init__(self, link, launch_callback, edit_callback, main_app):
        super().__init__()
        self.link = link
        self.launch_callback = launch_callback
        self.edit_callback = edit_callback
        self.main_app = main_app
        self.browsers, self.browser_paths = get_installed_browsers()
        self.main_app.browser_paths = self.browser_paths
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)
        self.main_button = ClickableLabel(self.link["name"])
        self.main_button.setToolTip(self.link["tooltip"])
        self.main_button.setWordWrap(True)
        # Fetch color from user preferences using mode-specific key, default to #0078d4
        link_key = self.main_app.get_link_settings_key(self.link["name"])
        link_settings = self.main_app.user_preferences["link_settings"].get(link_key, {})
        button_color = link_settings.get("color", "#0078d4")
        hover_color = self.darken_color(button_color, 0.8)
        self.main_button.setStyleSheet(f"""
            QLabel {{
                background-color: {button_color};
                color: white;
                font-size: 14px;
                font-weight: bold;
                padding: 8px;
                border-radius: 5px;
                border: 1px solid #555;
                text-align: center;
                margin: 0px;
            }}
            QLabel:hover {{ background-color: {hover_color}; }}
        """)
        self.main_button.setFixedSize(150, 56)
        self.main_button.clicked.connect(self.launch_default)
        layout.addWidget(self.main_button)

        local_extensions = ('.exe', '.txt', '.pdf', '.docx', '.xlsx', '.xlsm', '.bat', '.dotx', '.py')
        url = self.link["url"].strip()
        url_normalized = os.path.normpath(url).lower()
        self.is_url = url_normalized.startswith(("http://", "https://"))
        url_basename = os.path.basename(url_normalized)
        self.is_local_file = not self.is_url and (any(url_basename.endswith(ext) for ext in local_extensions) or any(url_normalized.endswith(ext) for ext in local_extensions))

        self.browser_combo = QComboBox()
        if self.is_local_file:
            self.browser_combo.addItems(["Launch"])
            self.browser_combo.setEnabled(False)
        else:
            self.browser_combo.addItems(self.browsers)
            saved_browser = self.main_app.get_browser_choice(self.link["name"])
            if saved_browser in self.browsers:
                self.browser_combo.setCurrentText(saved_browser)
            self.browser_combo.currentTextChanged.connect(self.save_browser_choice)

        self.browser_combo.setStyleSheet("""
            QComboBox {
                background-color: #444;
                color: white;
                font-size: 14px;
                padding: 0px 2px 2px 5px;
                border: 1px solid #555;
                border-radius: 3px;
                min-height: 20px;
                width: 150px;
            }
            QComboBox::drop-down {
                border: none;
            }
            QComboBox QAbstractItemView {
                background-color: #333;
                color: white;
                font-size: 14px;
                selection-background-color: #555;
            }
        """)
        self.browser_combo.setFixedSize(150, 20)
        layout.addWidget(self.browser_combo)

        # Edit label
        self.edit_label = ClickableLabel("EDIT", self.main_button)
        self.edit_label.setVisible(False)
        self.edit_label.clicked.connect(self.edit_link)
        self.edit_label.setGeometry(117, 4, 28, 20)

        # Favorite star label
        is_favorite = link_settings.get("favorite", False)
        self.favorite_label = ClickableLabel("☆" if not is_favorite else "★", self.main_button)
        self.favorite_label.setStyleSheet("""
            QLabel {
                color: #ffeb3b;
                font-size: 14px;
                background-color: rgba(0, 0, 0, 50);
                border: 1px solid #777;
                padding: 2px;
            }
            QLabel:hover {
                background-color: rgba(0, 0, 0, 80);
            }
        """)
        self.favorite_label.setFixedSize(28, 20)
        self.favorite_label.setAlignment(Qt.AlignCenter)
        self.favorite_label.clicked.connect(self.toggle_favorite)
        self.favorite_label.setGeometry(4, 4, 28, 20)
        self.favorite_label.setVisible(False)

        self.setLayout(layout)

    def toggle_favorite(self):
        # Toggle the favorite status in user preferences using mode-specific key
        link_key = self.main_app.get_link_settings_key(self.link["name"])
        link_settings = self.main_app.user_preferences["link_settings"].get(link_key, {})
        link_settings["favorite"] = not link_settings.get("favorite", False)
        self.main_app.user_preferences["link_settings"][link_key] = link_settings
        # Update the star icon
        self.favorite_label.setText("★" if link_settings["favorite"] else "☆")
        # Save the updated user preferences
        self.main_app.save_user_preferences()
        # Find the search bar and get its current query
        search_widget = self.main_app.launchpad_widget.findChild(QLineEdit)
        current_query = search_widget.text() if search_widget else ""
        # Refresh the Launchpad tab with the current query
        self.main_app.filter_links(current_query)

    def darken_color(self, hex_color, factor=0.8):
        hex_color = hex_color.lstrip('#')
        rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        darkened = tuple(int(c * factor) for c in rgb)
        return f"#{darkened[0]:02x}{darkened[1]:02x}{darkened[2]:02x}"

    def launch_default(self):
        url = self.link["url"]
        if "<whid>" in url.lower() or "<WHID>" in url:
            whid, ok = QInputDialog.getText(self, "WHID Input", "Enter WHID for this link:")
            if not ok or not whid.strip():
                QMessageBox.warning(self, "Input Error", "A valid WHID is required.")
                return
            whid_input = whid.strip()
            if "<WHID>" in url:
                url = url.replace("<WHID>", whid_input.upper())
            if "<whid>" in url:
                url = url.replace("<whid>", whid_input.lower())

        if self.is_local_file:
            self.launch_callback(url, None)
        else:
            selected_browser = self.browser_combo.currentText()
            self.launch_callback(url, selected_browser if selected_browser else None)

    def save_browser_choice(self, browser):
        if browser and not self.is_local_file:
            self.main_app.save_browser_choice(self.link["name"], browser)

    def edit_link(self):
        self.edit_callback(self.link)

    def enterEvent(self, event):
        self.edit_label.setVisible(True)
        self.favorite_label.setVisible(True)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.edit_label.setVisible(False)
        self.favorite_label.setVisible(False)
        super().leaveEvent(event)

class ActionButtonWidget(QWidget):
    def __init__(self, text, callback, is_oncall=False):
        super().__init__()
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)
        button = QPushButton(text)
        if is_oncall:
            button.setStyleSheet("""
                QPushButton {
                    background-color: #d32f2f;
                    color: white;
                    font-size: 14px;
                    font-weight: bold;
                    padding: 8px;
                    border-radius: 5px;
                    min-height: 40px;
                    border: 1px solid #555;
                    text-align: center;
                    margin: 0px;
                }
                QPushButton:hover { background-color: #b71c1c; }
            """)
        else:
            button.setStyleSheet("""
                QPushButton {
                    background-color: #28a745;
                    color: white;
                    font-size: 14px;
                    font-weight: bold;
                    padding: 8px;
                    border-radius: 5px;
                    min-height: 40px;
                    border: 1px solid #555;
                    text-align: center;
                    margin: 0px;
                }
                QPushButton:hover { background-color: #218838; }
            """)
        button.setFixedSize(150, 56)
        button.clicked.connect(callback)
        layout.addWidget(button)
        self.setLayout(layout)

class AddExampleDialog(QDialog):
    def __init__(self, parent=None, main_app=None):
        super().__init__(parent)
        self.setWindowTitle("Add PowerShell Cmdlet/Example")
        self.main_app = main_app
        self.setup_ui()

    def setup_ui(self):
        layout = QFormLayout()

        # Cmdlet/Pipeline input
        self.cmdlet_input = QLineEdit()
        self.cmdlet_input.setPlaceholderText("e.g., Get-Content or Get-Content | ForEach-Object")
        layout.addRow("Cmdlet/Pipeline:", self.cmdlet_input)

        # Aliases input
        self.aliases_input = QLineEdit()
        self.aliases_input.setPlaceholderText("e.g., cat,gc,type (comma-separated, optional)")
        layout.addRow("Aliases:", self.aliases_input)

        # Examples input
        self.example_input = QTextEdit()
        self.example_input.setPlaceholderText("Enter one example per line, e.g.,\nGet-Content -Path .\\test.txt\nGet-Content | ForEach-Object { $_ }")
        layout.addRow("Examples:", self.example_input)

        # New fields for full pipeline examples
        self.full_pipeline_description = QLineEdit()
        self.full_pipeline_description.setPlaceholderText("e.g., Filter processes by CPU and select properties")
        layout.addRow("Pipeline Description:", self.full_pipeline_description)

        self.full_pipeline_code = QTextEdit()
        self.full_pipeline_code.setPlaceholderText("e.g., Get-Process | Where-Object {$_.CPU -gt 100} | Select-Object Name, CPU")
        layout.addRow("Pipeline Code:", self.full_pipeline_code)

        # Dialog buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

    def validate_and_accept(self):
        cmdlet = self.cmdlet_input.text().strip()
        aliases = [a.strip() for a in self.aliases_input.text().split(",") if a.strip()]
        examples = [line.strip() for line in self.example_input.toPlainText().split("\n") if line.strip()]
        full_pipeline_description = self.full_pipeline_description.text().strip()
        full_pipeline_code = self.full_pipeline_code.toPlainText().strip()

        # Validation checks
        if not cmdlet:
            QMessageBox.warning(self, "Input Error", "Cmdlet/Pipeline is required.")
            return
        if not examples and not full_pipeline_code:
            QMessageBox.warning(self, "Input Error", "At least one example or pipeline is required.")
            return
        if full_pipeline_code and not full_pipeline_description:
            QMessageBox.warning(self, "Input Error", "Pipeline description is required for pipeline code.")
            return

        # Prepare data dictionary
        self.example_data = {
            "cmdlet": cmdlet,
            "aliases": aliases,
            "examples": examples,
            "full_pipeline": {
                "description": full_pipeline_description,
                "code": full_pipeline_code
            } if full_pipeline_code else None
        }
        self.accept()

    def get_data(self):
        return self.example_data

class NewCommandDialog(QDialog):
    def __init__(self, parent=None, main_app=None, command=None, default_commands=None, custom_commands=None):
        super().__init__(parent)
        self.setWindowTitle("New Command" if not command else "Edit Command")
        self.default_commands = default_commands or []
        self.custom_commands = custom_commands or []
        self.cmdlet_syntax_cache = {}  # Cache for cmdlet syntax
        self.cmdlet_syntax_cache.clear()  # Remove after testing        
        self.command = command  # Store command for use
        self.main_app = main_app  # Reference to MainApp for get_file_path
        # Load help repository
        self.help_repo = self.load_help_repository()
        self.setup_ui()
        # Debounce timer for delayed help updates
        self.debounce_timer = QTimer(self)
        self.debounce_timer.setSingleShot(True)
        self.debounce_timer.timeout.connect(self.update_param_hint)
        self.steps_input.textChanged.connect(self.start_debounce)

    def load_help_repository(self):
        # Load help examples from powershell_help.json using MainApp's file path logic
        if not self.main_app:
            print("Error: MainApp reference not provided")
            return {}
        help_file = self.main_app.get_file_path("powershell_help.json")
        try:
            if not os.path.exists(help_file):
                raise FileNotFoundError(f"powershell_help.json not found at {help_file}")
            with open(help_file, 'r', encoding='utf-8-sig') as f:  # Handle BOM
                content = f.read().strip()
                if not content:
                    print(f"Error: Empty powershell_help.json at {help_file}")
                    return {}
                # Replace invalid control characters
                content = ''.join(c for c in content if c.isprintable() or c in '\n\r\t')
                help_data = json.loads(content)
                cmdlets = help_data.get("cmdlets", {})
                if not cmdlets:
                    print("Warning: No cmdlets found in powershell_help.json")
                return cmdlets
        except FileNotFoundError:
            default_repo = {
                "cmdlets": {
                    "Get-Content": {
                        "aliases": ["cat", "gc", "type"],
                        "commands": ["Get-Content -Path .\\log.txt"]
                    },
                    "Where-Object": {
                        "aliases": ["where", "?"],
                        "commands": ["Get-Process | Where-Object { $_.CPU -gt 1000 }"]
                    }
                }
            }
            os.makedirs(os.path.dirname(help_file), exist_ok=True)
            with open(help_file, 'w', encoding='utf-8') as f:
                json.dump(default_repo, f, indent=4)
            return default_repo["cmdlets"]
        except json.JSONDecodeError as e:
            print(f"Error: Invalid JSON in powershell_help.json: {e}")
            return {}
        except Exception as e:
            print(f"Error loading help repository: {e}")
            return {}
        
    def start_debounce(self):
        # Start or restart 500ms debounce timer on text change
        self.debounce_timer.start(500)

    def setup_ui(self):
        layout = QFormLayout()
        self.title_input = QLineEdit(self.command["title"] if self.command else "")
        layout.addRow("Title:", self.title_input)
        self.shell_combo = QComboBox()
        self.shell_combo.addItems(["CMD", "PowerShell", "Terminal"])
        if self.command:
            self.shell_combo.setCurrentText(self.command["shell"])
        layout.addRow("Shell:", self.shell_combo)
        self.elevated_check = QCheckBox()
        if self.command:
            self.elevated_check.setChecked(self.command["elevated"])
        layout.addRow("Run Elevated:", self.elevated_check)
        self.steps_input = QTextEdit()
        self.steps_input.setPlaceholderText(
            "Enter steps, one per line, in the format: content, delay\n"
            "Examples:\n"
            "ipconfig /all, 500ms\n"
            "output: Checking network..., 0ms\n"
            "ping <domain>, 1000ms\n"
            "echo %USERNAME%, 0ms\n"
            "Delays are optional (default is 0ms). Use %VAR% for env vars."
        )
        if self.command:
            steps_text = ""
            for step in self.command["steps"]:
                if step["type"] == "command":
                    steps_text += f"{step['content']}, {step['delay']}ms\n"
                else:
                    steps_text += f"output: {step['content']}, {step['delay']}ms\n"
            self.steps_input.setText(steps_text.strip())
        self.highlighter = CommandHighlighter(self.steps_input.document(), self.default_commands, self.custom_commands)
        layout.addRow("Steps:", self.steps_input)
        self.param_hint_label = QTextEdit()
        self.param_hint_label.setReadOnly(True)
        self.param_hint_label.setLineWrapMode(QTextEdit.WidgetWidth)
        self.param_hint_label.setText("Example: (type a cmdlet to see usage)")
        self.param_hint_label.setStyleSheet("""
            QTextEdit {
                color: #ffeb3b;
                font-size: 12px;
                background-color: #333;
                padding: 2px;
                border: 1px solid #555;
            }
            QTextEdit:focus {
                border: 1px solid #777;
            }
        """)
        hint_container = QWidget()
        hint_layout = QVBoxLayout()
        hint_layout.setContentsMargins(0, 0, 0, 0)
        hint_layout.addWidget(self.param_hint_label)
        hint_container.setLayout(hint_layout)
        layout.addRow(hint_container)
        self.pause_between_check = QCheckBox()
        self.pause_between_check.setChecked(self.command["pause_between_commands"] if self.command and "pause_between_commands" in self.command else False)
        layout.addRow("Pause Between Commands:", self.pause_between_check)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)
        self.steps_input.installEventFilter(self)

    def eventFilter(self, obj, event):
        if obj == self.steps_input:
            if event.type() == QEvent.FocusIn or event.type() == QEvent.KeyPress:
                if not self.debounce_timer.isActive():
                    self.update_param_hint()
        return super().eventFilter(obj, event)

    def update_param_hint(self):
        MAX_EXAMPLES = 15
        text = self.steps_input.toPlainText()
        lines = text.split("\n")
        if not lines:
            self.param_hint_label.setHtml("Example: (type a cmdlet)")
            return
        current_line = next((line.strip() for line in reversed(lines) if line.strip()), "")
        if not current_line:
            self.param_hint_label.setHtml("Example: (type a cmdlet)")
            return
        cmd_parts = [part.strip() for part in current_line.split(",", 1)[0].split("|")]
        if not cmd_parts:
            self.param_hint_label.setHtml("Example: (type a cmdlet)")
            return
        last_part = cmd_parts[-1].strip()
        pipeline_cmdlets = ["where-object", "format-table", "select-object", "sort-object", "foreach-object", "select-string"]
        if not last_part or last_part == "|":
            self.param_hint_label.setHtml(f"Pipeline options: <b>{', '.join(pipeline_cmdlets)}</b>")
            return
        full_pipeline = " | ".join(cmd_parts).strip()
        main_cmdlet = cmd_parts[0].split()[0].strip() if cmd_parts[0].split() else ""
        last_cmdlet = last_part.split(" ")[0].strip() if last_part else ""
        if not main_cmdlet:
            self.param_hint_label.setHtml("Example: (type a valid cmdlet)")
            return
        # Load aliases dynamically from help_repo
        cmdlet_aliases = {cmd: data.get("aliases", []) for cmd, data in self.help_repo.items()}
        canonical_main = main_cmdlet
        canonical_last = last_cmdlet
        for cmd, aliases in cmdlet_aliases.items():
            cmd_lower = cmd.lower()
            if main_cmdlet.lower() == cmd_lower or main_cmdlet.lower() in [a.lower() for a in aliases]:
                canonical_main = cmd
            if last_cmdlet.lower() == cmd_lower or last_cmdlet.lower() in [a.lower() for a in aliases]:
                canonical_last = cmd
        cache_key = f"{full_pipeline.lower()}"
        examples = self.cmdlet_syntax_cache.get(cache_key, [])
        if not examples:
            if len(cmd_parts) > 1:
                if canonical_main in self.help_repo and "pipelines" in self.help_repo[canonical_main]:
                    pipeline_examples = self.help_repo[canonical_main]["pipelines"].get(canonical_last, [])
                    if pipeline_examples:
                        examples = pipeline_examples
                    else:
                        examples = [f"No examples found for {full_pipeline}"]
                else:
                    examples = [f"No examples found for {full_pipeline}"]
            else:
                if canonical_main in self.help_repo:
                    examples = self.help_repo[canonical_main].get("commands", [])
                else:
                    examples = [f"No examples found for {full_pipeline}"]
            self.cmdlet_syntax_cache[cache_key] = examples
        if examples and examples[0].startswith("No examples found"):
            self.param_hint_label.setHtml(examples[0])
        else:
            example_text = "<br>".join(ex for ex in examples[:MAX_EXAMPLES] if ex.strip())
            # Add full pipeline examples
            if canonical_main in self.help_repo and "fullPipelines" in self.help_repo[canonical_main]:
                full_pipelines = self.help_repo[canonical_main]["fullPipelines"]
                if full_pipelines:
                    example_text += "<br><b>Full Pipeline Examples:</b><br>"
                    for fp in full_pipelines:
                        example_text += f"<b>{fp['description']}:</b> {fp['code']}<br>"
            self.param_hint_label.setHtml(example_text or f"No examples found for {full_pipeline}")
        self.param_hint_label.setLineWrapMode(QTextEdit.WidgetWidth)

    def validate_and_accept(self):
        command = self.get_command()
        if command:
            self.command = command
            self.accept()

    def get_command(self):
        steps_text = self.steps_input.toPlainText().strip()
        steps = []
        for line in steps_text.split("\n"):
            line = line.strip()
            if not line:
                continue
            try:
                parts = [part.strip() for part in line.split(",", 1)]
                if len(parts) < 1:
                    QMessageBox.warning(self, "Input Error", f"Invalid step format: {line}\nUse: content, delay")
                    return None
                content = parts[0]
                delay = 0
                if len(parts) > 1 and parts[1]:
                    delay_match = re.match(r'(\d+)ms', parts[1])
                    if delay_match:
                        delay = int(delay_match.group(1))
                    else:
                        QMessageBox.warning(self, "Input Error", f"Invalid delay format: {parts[1]}\nUse: Xms (e.g., 500ms)")
                        return None
                if content.lower().startswith("output:"):
                    step_type = "output"
                    content = content[7:].strip()
                else:
                    step_type = "command"
                if not content:
                    QMessageBox.warning(self, "Input Error", f"Empty step content in line: {line}")
                    return None
                steps.append({
                    "type": step_type,
                    "content": content,
                    "delay": delay
                })
            except Exception as e:
                QMessageBox.warning(self, "Input Error", f"Error parsing step: {line}\n{e}")
                return None
        title = self.title_input.text().strip()
        if not title:
            QMessageBox.warning(self, "Input Error", "Title cannot be empty.")
            return None
        if not steps:
            QMessageBox.warning(self, "Input Error", "At least one step is required.")
            return None
        return {
            "title": title,
            "steps": steps,
            "shell": self.shell_combo.currentText(),
            "elevated": self.elevated_check.isChecked(),
            "pause_between_commands": self.pause_between_check.isChecked()
        }

class NewGuideDialog(QDialog):
    def __init__(self, parent=None, guide=None):
        super().__init__(parent)
        self.setWindowTitle("New Guide" if not guide else "Edit Guide")
        self.guide = guide
        self.layout = QVBoxLayout()
        self.setGeometry(100, 100, 800, 600)
        self.setup_ui()
        self.setLayout(self.layout)
    def setup_ui(self):
        form_layout = QFormLayout()
        self.steps_input.setPlaceholderText(
            "Enter steps, one per line, in the format: content, delay\n"
            "Examples:\n"
            "ipconfig /all, 500ms\n"
            "output: Checking network..., 0ms\n"
            "ping <domain>, 1000ms\n"
            "echo %USERNAME%, 0ms\n"
            "Delays are optional (default is 0ms). Use %VAR% for env vars."
        )
        form_layout.addRow("Title:", self.title_input)
        self.desc_label = QLabel("Description:")
        form_layout.addRow(self.desc_label)
        self.desc_input = QTextEdit()
        self.desc_input.setAcceptRichText(True)
        font = QFont()
        font.setPointSize(16)
        self.desc_input.setFont(font)
        self.desc_input.setStyleSheet("QTextEdit { font-size: 16px; }")
        if self.guide:
            desc = self.guide["description"]
            if not desc.startswith("<"):
                desc = f'<p style="font-size: 16px;">{desc}</p>'
            self.desc_input.setHtml(desc)
        self.desc_toolbar = QToolBar()
        self.desc_toolbar.setStyleSheet("""
            QToolBar {
                background: #444;
                border: 1px solid #555;
                padding: 2px;
            }
            QToolButton {
                color: white;
                font-size: 16px;
                font-weight: bold;
                background: #666;
                border: 1px solid #777;
                padding: 4px;
                margin: 2px;
                min-width: 24px;
                min-height: 24px;
            }
            QToolButton:hover {
                background: #888;
            }
        """)
        self.add_formatting_buttons(self.desc_toolbar, self.desc_input)
        form_layout.addRow(self.desc_toolbar)
        form_layout.addRow(self.desc_input)
        self.steps_label = QLabel("Steps:")
        form_layout.addRow(self.steps_label)
        self.steps_input = QTextEdit()
        self.steps_input.setAcceptRichText(True)
        self.steps_input.setFont(font)
        self.steps_input.setStyleSheet("QTextEdit { font-size: 16px; }")
        if self.guide:
            if isinstance(self.guide["steps"], list):
                steps_html = '<ol style="font-size: 16px;">' + "".join(f"<li>{step}</li>" for step in self.guide["steps"]) + "</ol>"
                self.steps_input.setHtml(steps_html)
            else:
                steps = self.guide["steps"]
                if not steps.startswith("<"):
                    steps = f'<p style="font-size: 16px;">{steps}</p>'
                self.steps_input.setHtml(steps)
        self.steps_toolbar = QToolBar()
        self.steps_toolbar.setStyleSheet("""
            QToolBar {
                background: #444;
                border: 1px solid #555;
                padding: 2px;
            }
            QToolButton {
                color: white;
                font-size: 16px;
                font-weight: bold;
                background: #666;
                border: 1px solid #777;
                padding: 4px;
                margin: 2px;
                min-width: 24px;
                min-height: 24px;
            }
            QToolButton:hover {
                background: #888;
            }
        """)
        self.add_formatting_buttons(self.steps_toolbar, self.steps_input)
        form_layout.addRow(self.steps_toolbar)
        form_layout.addRow(self.steps_input)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        form_layout.addWidget(buttons)
        self.layout.addLayout(form_layout)
    def add_formatting_buttons(self, toolbar, text_edit):
        bold_btn = QAction("B", self)
        bold_btn.setToolTip("Toggle Bold")
        bold_btn.triggered.connect(lambda: self.toggle_format(text_edit, "bold"))
        toolbar.addAction(bold_btn)
        italic_btn = QAction("I", self)
        italic_btn.setToolTip("Toggle Italic")
        italic_btn.triggered.connect(lambda: self.toggle_format(text_edit, "italic"))
        toolbar.addAction(italic_btn)
        for button in toolbar.findChildren(QToolButton):
            if button.defaultAction() == italic_btn:
                italic_font = QFont()
                italic_font.setItalic(True)
                button.setFont(italic_font)
                break
        bullet_btn = QAction("•", self)
        bullet_btn.setToolTip("Toggle Bullet List")
        bullet_btn.triggered.connect(lambda: self.toggle_bullet_list(text_edit))
        toolbar.addAction(bullet_btn)
        number_btn = QAction("1.", self)
        number_btn.setToolTip("Toggle Numbered List")
        number_btn.triggered.connect(lambda: self.toggle_numbered_list(text_edit))
        toolbar.addAction(number_btn)
    def toggle_format(self, text_edit, format_type):
        cursor = text_edit.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.WordUnderCursor)
        char_format = cursor.charFormat()
        if format_type == "bold":
            char_format.setFontWeight(QFont.Bold if char_format.fontWeight() != QFont.Bold else QFont.Normal)
        elif format_type == "italic":
            char_format.setFontItalic(not char_format.fontItalic())
        cursor.setCharFormat(char_format)
        text_edit.setTextCursor(cursor)
    def toggle_bullet_list(self, text_edit):
        cursor = text_edit.textCursor()
        text_edit.setFocus()
        list_format = cursor.block().textList()
        cursor.beginEditBlock()
        if list_format:
            current_style = list_format.format().style()
            if current_style == QTextListFormat.ListDisc:
                list_format.remove(cursor.block())
            else:
                cursor.createList(QTextListFormat.ListDisc)
        else:
            cursor.createList(QTextListFormat.ListDisc)
        cursor.endEditBlock()
        text_edit.setTextCursor(cursor)
    def toggle_numbered_list(self, text_edit):
        cursor = text_edit.textCursor()
        text_edit.setFocus()
        list_format = cursor.block().textList()
        cursor.beginEditBlock()
        if list_format:
            current_style = list_format.format().style()
            if current_style == QTextListFormat.ListDecimal:
                list_format.remove(cursor.block())
            else:
                cursor.createList(QTextListFormat.ListDecimal)
        else:
            cursor.createList(QTextListFormat.ListDecimal)
        cursor.endEditBlock()
        text_edit.setTextCursor(cursor)
    def validate_and_accept(self):
        guide = self.get_guide()
        if guide:
            self.guide_data = guide
            self.accept()
    def get_guide(self):
        title = self.title_input.text().strip()
        description = self.desc_input.toHtml().strip()
        steps = self.steps_input.toHtml().strip()
        if not title:
            QMessageBox.warning(self, "Input Error", "Title cannot be empty.")
            return None
        if not description or description == "<p></p>":
            QMessageBox.warning(self, "Input Error", "Description cannot be empty.")
            return None
        if not steps or steps == "<p></p>":
            QMessageBox.warning(self, "Input Error", "Steps cannot be empty.")
            return None
        return {
            "title": title,
            "description": description,
            "steps": steps
        }

class ImportCommandsDialog(QDialog):
    def __init__(self, commands, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Import Commands")
        self.layout = QVBoxLayout()
        self.command_list = QListWidget()
        self.command_list.setSelectionMode(QListWidget.MultiSelection)
        for command in commands:
            item = QListWidgetItem(command["title"])
            item.setData(Qt.UserRole, command)
            self.command_list.addItem(item)
        self.layout.addWidget(self.command_list)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        self.layout.addWidget(buttons)
        self.setLayout(self.layout)
    def get_selected_commands(self):
        selected_commands = []
        for item in self.command_list.selectedItems():
            command = item.data(Qt.UserRole)
            selected_commands.append(command)
        return selected_commands

class ImportGuidesDialog(QDialog):
    def __init__(self, guides, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Import Guides")
        self.layout = QVBoxLayout()
        self.guide_list = QListWidget()
        self.guide_list.setSelectionMode(QListWidget.MultiSelection)
        for guide in guides:
            item = QListWidgetItem(guide["title"])
            item.setData(Qt.UserRole, guide)
            self.guide_list.addItem(item)
        self.layout.addWidget(self.guide_list)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        self.layout.addWidget(buttons)
        self.setLayout(self.layout)
    def get_selected_guides(self):
        selected_guides = []
        for item in self.guide_list.selectedItems():
            guide = item.data(Qt.UserRole)
            selected_guides.append(guide)
        return selected_guides
        
class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("LaunchPad Help")
        self.setGeometry(200, 200, 1000, 600)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()
        self.content = QTextBrowser()
        self.content.setStyleSheet("""
            QTextBrowser {
                background-color: #333;
                color: #cccccc;
                border: 1px solid #444;
                font-size: 16px;
                padding: 5px;
            }
        """)
        self.content.setOpenExternalLinks(True)
        self.content.setHtml("""
            <h2>Welcome to LaunchPad</h2>
            <p>LaunchPad is your IT Hub for quick access to tools, commands, and guides.</p>
            <h3>Getting Started</h3>
            <ul>
                <li><b>Launchpad Tab</b>: Create links to frequently used sites and local apps, use placeholder commands <code>&lt;whid&gt;</code> or <code>&lt;WHID&gt;</code> to force user to enter a whid in either lower or uppercase. Use the search bar to filter links. Sort in alphabetical or reverse alphabetical </li>
                <li><b>Commands Tab</b>: Create and Run command sequences, use "pause between commands" and/or use delay Xms (e.g. 1000ms) between commands. Use placeholders like <code>&lt;domain&gt;</code> or <code>&lt;whid&gt;</code> which forces user input when ran for a domain / IP or whid.</li>
                <li><b>How-To Guides Tab</b>: View, create, or edit step-by-step guides for IT tasks and/or new hires. URLs are parsed and automatically made into links.</li>
                <li><b>Page OnCall</b>: Send an email to the on-call team via Outlook. Edit the recipient if necessary, subject, and message in the dialog. Default oncall email can be set in Settings -> Configure JSON Storage</li>
            </ul>
            <h3>Tips</h3>
            <ul>
                <li>Use the "Settings" menu to switch between Local and Shared storage modes.</li>
                <li>Edit links by hovering over them and clicking "EDIT".</li>
                <li>Check the "How-To Guides" for detailed instructions on common tasks.</li>
            </ul>
            <h3>Version Information</h3>
            <p><b>Version</b>: 2.0.0<br>
               <b>Release Date</b>: May 2025<br>
               <b>Developed by</b>: Enda Rensing (enrensing@amazon.com)</p>
        """)
        layout.addWidget(self.content)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)
        self.setLayout(layout)


class ChangelogDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Changelog")
        self.setGeometry(200, 200, 1000, 600)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()
        self.content = QTextBrowser()
        self.content.setStyleSheet("""
            QTextBrowser {
                background-color: #333;
                color: #cccccc;
                border: 1px solid #444;
                font-size: 16px;
                padding: 5px;
            }
        """)
        self.content.setOpenExternalLinks(True)
        self.content.setHtml("""
            <h2>LaunchPad Changelog</h2>
            <h3>Version 2.0.0 (May 24, 2025)</h3>
            <ul>
                <li><b>Environment Variable Support:</b> Added support for environment variables (e.g., %USERNAME%) in command steps for dynamic execution.</li>
                <li><b>Syntax Highlighting:</b> Implemented syntax highlighting in the Commands tab for recognized commands (darker blue: #1e90ff) and placeholders (dark green: #008000, bold).</li>
                <li><b>Dynamic Command Detection:</b> Updated default_commands to dynamically fetch available commands from the system PATH and CMD shell.</li>
                <li><b>UI Improvements:</b> Removed the default Windows "ding" sound from Help > About and Help > Version Info dialogs.</li>
                <li><b>Versioning:</b> Bumped version to 2.0.0 to reflect major enhancements.</li>
            </ul>
            <h3>Version 1.0.0 (May 2025)</h3>
            <ul>
                <li>Initial release with Launchpad, Commands, and How-To Guides tabs.</li>
                <li>Support for <WHID> placeholders in links and commands.</li>
                <li>Page OnCall and Taka Helper integration.</li>
            </ul>
        """)
        layout.addWidget(self.content)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)
        self.setLayout(layout)

class MainApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("IT Hub")
        self.setGeometry(100, 100, 800, 600)
        self.setMinimumWidth(320)
        # Fetch commands from PATH (Windows)
        self.default_commands = []
        common_commands = ["net", "ipconfig", "nslookup", "netstat", "ping", "tracert", "sc", "cmd", "powershell", "wt", "wmic", "echo"]
        for cmd in common_commands:
            self.default_commands.append(cmd.lower())  # Add explicitly
        # Fetch CMD built-ins
        try:
            cmd_output = subprocess.check_output("cmd /c help", shell=True, text=True).splitlines()
            for line in cmd_output:
                cmd = line.strip().split()[0].lower()
                if cmd and cmd not in self.default_commands and len(cmd) > 1:
                    self.default_commands.append(cmd)
        except subprocess.CalledProcessError:
            pass
        # Fetch PowerShell cmdlets
        try:
            ps_output = subprocess.check_output('powershell -Command "Get-Command -CommandType Cmdlet | ForEach-Object { $_.Name }"', shell=True, text=True).splitlines()
            for cmdlet in ps_output:
                cmd = cmdlet.strip().lower()
                if cmd and cmd not in self.default_commands:
                    self.default_commands.append(cmd)
        except subprocess.CalledProcessError:
            pass

        self.custom_commands = []
        self.app_data_dir = os.path.join(os.getenv('LOCALAPPDATA'), 'LaunchPad')
        os.makedirs(self.app_data_dir, exist_ok=True)
        self.settings_file = os.path.join(self.app_data_dir, "settings.json")
        self.user_preferences_file = os.path.join(self.app_data_dir, "user_preferences.json")
        self.browser_paths = {}
        self.load_settings()
        self.load_user_preferences()
        if self.settings.get("mode") == "Shared":
            network_path = self.settings.get("network_path", "")
            self.commands_file = os.path.join(network_path, "commands.json")
            self.links_file = os.path.join(network_path, "launchpad_links.json")
            self.guides_file = os.path.join(network_path, "howto_guides.json")
        else:
            self.commands_file = os.path.join(self.app_data_dir, "commands.json")
            self.links_file = os.path.join(self.app_data_dir, "launchpad_links.json")
            self.guides_file = os.path.join(self.app_data_dir, "howto_guides.json")
        self.load_commands()
        self.load_links()
        self.load_guides()
        self.sort_order = "Alphabetical (A-Z)"  # Default to A-Z sorting
        self.search_query = ""
        self.help_files_updated = self.user_preferences.get("help_files_updated", False)
        menu_bar = self.menuBar()
        settings_menu = QMenu("Settings", self)
        configure_action = QAction("Configure JSON Storage", self)
        configure_action.triggered.connect(self.open_settings_dialog)
        settings_menu.addAction(configure_action)
        refresh_action = QAction("Refresh Data", self)
        refresh_action.triggered.connect(self.refresh_shared_files)
        settings_menu.addAction("Add PowerShell Example", self.open_add_example_dialog)        
        settings_menu.addAction(refresh_action)
        menu_bar.addMenu(settings_menu)
        help_menu = QMenu("Help", self)
        user_guide_action = QAction("User Guide", self)
        user_guide_action.triggered.connect(self.show_user_guide)
        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about)
        version_action = QAction("Version Info", self)
        version_action.triggered.connect(self.show_version_info)
        changelog_action = QAction("Changelog", self)
        changelog_action.triggered.connect(self.show_changelog)
        help_menu.addAction(user_guide_action)
        help_menu.addAction(about_action)
        help_menu.addAction(version_action)
        help_menu.addAction(changelog_action)
        menu_bar.addMenu(help_menu)
        self.tabs = QTabWidget()
        self.launchpad_widget = self.launchpad_tab()
        self.tabs.addTab(self.launchpad_widget, "Launchpad")
        self.tabs.addTab(self.commands_tab(), "Commands")
        self.tabs.addTab(self.howto_guides_tab(), "How-To Guides")
        self.setCentralWidget(self.tabs)
        import ctypes
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 6)  # 6 = minimize
        self.setStyleSheet("""
            QMainWindow {
                background-color: #2b2b2b;
            }
            QTabWidget::pane {
                border: 1px solid #444;
                background: #333;
                margin-top: -1px;
            }
            QTabBar::tab {
                padding: 10px 20px;
                font-size: 16px;
                font-weight: bold;
                min-width: 140px;
                min-height: 30px;
                border: 1px solid #444;
            }
            QTabBar::tab:nth-child(3) {
                min-width: 250px;
            }
        """)
        self.tabs.currentChanged.connect(self.update_tab_styles)
        self.update_tab_styles(0)
        self.filter_links("")  # Ensure initial render with colors

    def get_link_settings_key(self, link_name):
        mode = self.settings.get("mode", "Local")
        return f"{mode}:{link_name}"

    def load_settings(self):
            if os.path.exists(self.settings_file):
                try:
                    with open(self.settings_file, 'r') as f:
                        self.settings = json.load(f)
                except Exception as e:
                    print(f"Failed to load settings: {str(e)}")
                    self.settings = {"mode": "Local", "network_path": "", "oncall_email": "page-ots-oncall-nbwi1@amazon.com"}
            else:
                self.settings = {"mode": "Local", "network_path": "", "oncall_email": "page-ots-oncall-nbwi1@amazon.com"}
                self.save_settings()

    def save_settings(self):
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f, indent=4)
        except Exception as e:
            print(f"Failed to save settings: {str(e)}")

    def load_user_preferences(self):
        if os.path.exists(self.user_preferences_file):
            try:
                with open(self.user_preferences_file, 'r') as f:
                    self.user_preferences = json.load(f)
            except Exception as e:
                print(f"Failed to load user preferences: {str(e)}")
                self.user_preferences = {"link_settings": {}, "powershell_help_updated": False}
        else:
            self.user_preferences = {"link_settings": {}, "powershell_help_updated": False}
            self.save_user_preferences()
        # Migrate browser choices from link_browser_choices.json if it exists
        browser_choices_file = os.path.join(self.app_data_dir, "link_browser_choices.json")
        if os.path.exists(browser_choices_file):
            try:
                with open(browser_choices_file, 'r') as f:
                    browser_choices = json.load(f)
                for link_key, browser in browser_choices.items():
                    if link_key in self.user_preferences["link_settings"]:
                        self.user_preferences["link_settings"][link_key]["browser"] = browser
                    else:
                        self.user_preferences["link_settings"][link_key] = {"color": "#0078d4", "favorite": False, "browser": browser}
                self.save_user_preferences()
                os.remove(browser_choices_file)  # Delete the old file
            except Exception as e:
                print(f"Failed to migrate browser choices: {str(e)}")
        # Ensure powershell_help_updated exists
        if "powershell_help_updated" not in self.user_preferences:
            self.user_preferences["powershell_help_updated"] = False
            self.save_user_preferences()

    def save_user_preferences(self):
        if self.acquire_lock(self.user_preferences_file):
            try:
                with open(self.user_preferences_file, 'w') as f:
                    json.dump(self.user_preferences, f, indent=4)
            except Exception as e:
                print(f"Debug: Save failed: {str(e)}")
            finally:
                self.release_lock(self.user_preferences_file)
        else:
            print(f"Debug: Lock failed for {self.user_preferences_file}")

    def get_browser_choice(self, link_name):
        link_key = self.get_link_settings_key(link_name)
        link_settings = self.user_preferences["link_settings"].get(link_key, {})
        return link_settings.get("browser", "")

    def save_browser_choice(self, link_name, browser):
        link_key = self.get_link_settings_key(link_name)
        link_settings = self.user_preferences["link_settings"].get(link_key, {})
        link_settings["browser"] = browser
        if "color" not in link_settings:
            link_settings["color"] = "#0078d4"
        if "favorite" not in link_settings:
            link_settings["favorite"] = False
        self.user_preferences["link_settings"][link_key] = link_settings
        self.save_user_preferences()

    def get_file_path(self, filename):
        if self.settings.get("mode") == "Shared" and self.settings.get("network_path"):
            return os.path.join(self.settings["network_path"], filename)
        else:
            app_data = os.getenv("LOCALAPPDATA")
            return os.path.join(app_data, "LaunchPad", filename)

    def acquire_lock(self, filename):
        lock_file = f"{filename}.lock"
        max_attempts = 30
        for _ in range(max_attempts):
            if not os.path.exists(lock_file):
                try:
                    with open(lock_file, 'w') as f:
                        f.write(str(os.getpid()))
                    return True
                except:
                    pass
            time.sleep(1)
        QMessageBox.warning(self, "Error", f"Could not acquire lock for {filename}. Another user may be editing the file.")
        return False

    def release_lock(self, filename):
        lock_file = f"{filename}.lock"
        if os.path.exists(lock_file):
            try:
                os.remove(lock_file)
            except:
                pass

    def load_commands(self):
        default_commands = [
            {"title": "Open Elevated CMD", "steps": [{"type": "command", "content": "cmd /k", "delay": 0}], "shell": "CMD", "elevated": True, "pause_between_commands": False},
            {"title": "Open Elevated PowerShell", "steps": [{"type": "command", "content": "powershell", "delay": 0}], "shell": "PowerShell", "elevated": True, "pause_between_commands": False},
            {"title": "Open Elevated Terminal", "steps": [{"type": "command", "content": "wt", "delay": 0}], "shell": "Terminal", "elevated": True, "pause_between_commands": False},
            {"title": "Reset AD Password", "steps": [{"type": "command", "content": "password.amazon.com", "delay": 0}], "shell": "CMD", "elevated": False, "pause_between_commands": False},
            {"title": "Clear DNS Cache", "steps": [{"type": "command", "content": "ipconfig /flushdns", "delay": 500}, {"type": "command", "content": "nslookup <domain>", "delay": 0}], "shell": "CMD", "elevated": False, "pause_between_commands": False},
            {"title": "Check Open Ports", "steps": [{"type": "command", "content": "netstat -a | findstr LISTENING", "delay": 500}, {"type": "output", "content": "Review output for active listeners.", "delay": 0}, {"type": "output", "content": "Note: Use netstat -an for faster numerical output.", "delay": 0}], "shell": "CMD", "elevated": False, "pause_between_commands": True}
        ]
        if os.path.exists(self.commands_file):
            try:
                if self.acquire_lock(self.commands_file):
                    with open(self.commands_file, 'r') as f:
                        self.commands = json.load(f)
                    for command in self.commands:
                        if "pause_between_commands" not in command:
                            command["pause_between_commands"] = False
                        if "pause_at_end" in command:
                            command["pause_between_commands"] = command.pop("pause_at_end")
                    self.release_lock(self.commands_file)
                else:
                    self.commands = default_commands
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to load commands from file: {str(e)}\nUsing default commands.")
                self.commands = default_commands
        else:
            self.commands = default_commands

    def load_links(self):
        default_links = [
            {"name": "Team Wiki", "url": "https://wiki.example.com", "tooltip": "Wiki resources", "icon": "icon-wiki.png"},
            {"name": "IT Policies", "url": "https://policies.example.com", "tooltip": "Company IT policies", "icon": "icon-policy.png"},
            {"name": "Helpdesk", "url": "https://helpdesk.example.com", "tooltip": "Submit or view tickets", "icon": "icon-helpdesk.png"},
            {"name": "IT Dashboard", "url": "https://dashboard.example.com", "tooltip": "IT metrics and stats", "icon": "icon-dashboard.png"},
            {"name": "Support Portal", "url": "https://support.example.com", "tooltip": "Access support resources", "icon": "icon-support.png"},
            {"name": "Outlook", "url": "outlook.exe", "tooltip": "Launch local Outlook client", "icon": "icon-outlook.png"},
            {"name": "OWA", "url": "https://outlook.office365.com", "tooltip": "Open Outlook Web App", "icon": "icon-owa.png"}
        ]
        if os.path.exists(self.links_file):
            try:
                if self.acquire_lock(self.links_file):
                    with open(self.links_file, 'r') as f:
                        self.links = json.load(f)
                    # Migrate existing color and favorite fields to user preferences
                    for link in self.links:
                        link_name = link["name"]
                        link_key = self.get_link_settings_key(link_name)
                        if "color" in link or "favorite" in link:
                            self.user_preferences["link_settings"][link_key] = {
                                "color": link.get("color", "#0078d4"),
                                "favorite": link.get("favorite", False)
                            }
                            # Remove color and favorite from the link
                            link.pop("color", None)
                            link.pop("favorite", None)
                    # Clean up orphaned preferences
                    current_link_names = {link["name"] for link in self.links}
                    current_mode = self.settings.get("mode", "Local")
                    preferences_to_remove = []
                    for key in self.user_preferences["link_settings"]:
                        mode, link_name = key.split(":", 1)
                        if mode == current_mode and link_name not in current_link_names:
                            preferences_to_remove.append(key)
                    for key in preferences_to_remove:
                        del self.user_preferences["link_settings"][key]
                    self.save_user_preferences()
                    self.release_lock(self.links_file)
                else:
                    self.links = default_links
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to load links from file: {str(e)}\nUsing default links.")
                self.links = default_links
        else:
            self.links = default_links

    def load_guides(self):
        default_guides = [
            {"title": "Set Up VPN", "description": '<p style="font-size: 16px;">Guide to configure VPN for remote access.</p>', "steps": '<ol style="font-size: 16px;"><li>Download VPN client from IT portal</li><li>Install the client</li><li>Enter company credentials</li><li>Connect to the VPN server</li></ol>'},
            {"title": "Access Shared Drive", "description": '<p style="font-size: 16px;">Steps to access the company shared drive.</p>', "steps": '<ol style="font-size: 16px;"><li>Connect to VPN</li><li>Open File Explorer</li><li>Navigate to \\\\server\\share</li><li>Enter credentials if prompted</li></ol>'}
        ]
        if os.path.exists(self.guides_file):
            try:
                if self.acquire_lock(self.guides_file):
                    with open(self.guides_file, 'r') as f:
                        self.guides = json.load(f)
                    for guide in self.guides:
                        if isinstance(guide["steps"], list):
                            guide["steps"] = '<ol style="font-size: 16px;">' + "".join(f"<li>{step}</li>" for step in guide["steps"]) + "</ol>"
                        if not guide["description"].startswith("<"):
                            guide["description"] = f'<p style="font-size: 16px;">{guide["description"]}</p>'
                    self.release_lock(self.guides_file)
                else:
                    self.guides = default_guides
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to load guides from file: {str(e)}\nUsing default guides.")
                self.guides = default_guides
        else:
            self.guides = default_guides
            self.save_guides()

    def save_links(self):
        if self.acquire_lock(self.links_file):
            try:
                with open(self.links_file, 'w') as f:
                    json.dump(self.links, f, indent=4)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save links to file: {str(e)}")
            finally:
                self.release_lock(self.links_file)
        else:
            QMessageBox.warning(self, "Error", "Could not save links: file is locked by another user.")

    def save_commands(self):
        if self.acquire_lock(self.commands_file):
            try:
                with open(self.commands_file, 'w') as f:
                    json.dump(self.commands, f, indent=4)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save commands to file: {str(e)}")
            finally:
                self.release_lock(self.commands_file)
        else:
            QMessageBox.warning(self, "Error", "Could not save commands: file is locked by another user.")

    def save_guides(self):
        if self.acquire_lock(self.guides_file):
            try:
                with open(self.guides_file, 'w') as f:
                    json.dump(self.guides, f, indent=4)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save guides to file: {str(e)}")
            finally:
                self.release_lock(self.guides_file)
        else:
            QMessageBox.warning(self, "Error", "Could not save guides: file is locked by another user.")

    def closeEvent(self, event):
        self.save_commands()
        self.save_links()
        self.save_guides()
        self.save_user_preferences()
        event.accept()

    def update_tab_styles(self, index):
        stylesheet = ""
        for i in range(self.tabs.count()):
            if i == index:
                self.tabs.tabBar().setTabTextColor(i, Qt.yellow)
                stylesheet += f"""
                    QTabBar::tab:nth-child({i+1}) {{
                        background: #222;
                        color: #ffeb3b;
                        font-size: 18px;
                        font-weight: bold;
                        border: 1px solid #444;
                        border-bottom: none;
                        margin-bottom: -1px;
                    }}
                """
            else:
                self.tabs.tabBar().setTabTextColor(i, Qt.gray)
                stylesheet += f"""
                    QTabBar::tab:nth-child({i+1}) {{
                        background: #222;
                        color: #aaa;
                        font-size: 16px;
                        font-weight: bold;
                        border: 1px solid #444;
                        border-bottom: 1px solid #444;
                    }}
                """
        self.tabs.tabBar().setStyleSheet(stylesheet)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.tabs.currentIndex() == 0:
            self.refresh_launchpad_tab()

    def filter_links(self, query):
        # Find the grid widget in the current launchpad tab
        if not hasattr(self, 'launchpad_widget') or not self.launchpad_widget:
            self.refresh_launchpad_tab()
            return
        grid_widget = self.launchpad_widget.findChild(QWidget, "grid_widget")
        if not grid_widget:
            self.refresh_launchpad_tab()
            return

        # Clear existing grid
        search_widget = self.launchpad_widget.findChild(QLineEdit)
        if search_widget:
            search_widget.blockSignals(True)
        grid_layout = grid_widget.layout()
        while grid_layout.count():
            item = grid_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # Calculate grid layout
        available_width = self.width() - 10
        button_width = 150
        spacing = 5
        buttons_per_row = max(1, available_width // (button_width + spacing))

        # Filter and sort links using the local query parameter
        filtered_links = [
            link for link in self.links
            if (query.lower() in link.get("name", "").lower() or
                query.lower() in link.get("url", "").lower() or
                query.lower() in link.get("tooltip", "").lower())
        ]
        # Split into favorites and non-favorites using user preferences
        favorites = []
        non_favorites = []
        for link in filtered_links:
            link_key = self.get_link_settings_key(link["name"])
            link_settings = self.user_preferences["link_settings"].get(link_key, {})
            if link_settings.get("favorite", False):
                favorites.append(link)
            else:
                non_favorites.append(link)
        # Sort each group according to the selected sort order
        if self.sort_order == "Alphabetical (A-Z)":
            favorites.sort(key=lambda x: x["name"].lower())
            non_favorites.sort(key=lambda x: x["name"].lower())
        elif self.sort_order == "Alphabetical (Z-A)":
            favorites.sort(key=lambda x: x["name"].lower(), reverse=True)
            non_favorites.sort(key=lambda x: x["name"].lower(), reverse=True)

        # Add action buttons (always in row 0)
        add_link_btn = ActionButtonWidget("Add Link", self.add_new_link)
        taka_helper_btn = ActionButtonWidget("Open Taka Helper", self.launch_taka_helper)
        page_oncall_btn = ActionButtonWidget("Page OnCall", self.send_oncall_email, is_oncall=True)
        buttons = [add_link_btn, taka_helper_btn, page_oncall_btn]

        row, col = 0, 0
        for btn in buttons:
            grid_layout.addWidget(btn, row, col)
            col += 1
            if col >= buttons_per_row:
                col = 0
                row += 1

        # Add favorited buttons (continue in row 0 until full)
        for link in favorites:
            local_extensions = ('.exe', '.txt', '.pdf', '.docx', '.xlsx', '.xlsm', '.bat', '.dotx', '.py')
            url = link["url"].strip()
            url_normalized = os.path.normpath(url).lower()
            is_url = url_normalized.startswith(("http://", "https://"))
            url_basename = os.path.basename(url_normalized)
            is_local_file = not is_url and (any(url_basename.endswith(ext) for ext in local_extensions) or any(url_normalized.endswith(ext) for ext in local_extensions))

            btn_widget = LinkButtonWidget(
                link,
                self.launch_local_app if is_local_file else self.open_website,
                self.edit_link,
                self
            )
            if col >= buttons_per_row:
                col = 0
                row += 1
            grid_layout.addWidget(btn_widget, row, col)
            col += 1

        # Add a labeled delimiter between row 1 (action buttons + favorites) and row 2 (non-favorited buttons)
        # Determine the row for the delimiter
        delimiter_row = row
        if col > 0:  # If the current row isn't empty, move to the next row for the delimiter
            col = 0
            delimiter_row += 1
        # Create a labeled delimiter
        delimiter = QLabel("Favorites Above | Regular Links Below")
        delimiter.setStyleSheet("color: #ffeb3b; font-size: 12px; padding: 2px 0px 2px 0px; border-top: 1px solid #666; border-bottom: 1px solid #666; background-color: #2b2b2b;")
        delimiter.setAlignment(Qt.AlignCenter)
        delimiter.setFixedHeight(16)
        grid_layout.addWidget(delimiter, delimiter_row, 0, 1, buttons_per_row)  # Span the entire row

        # Force non-favorited buttons to start on the row after the delimiter
        col = 0
        row = delimiter_row + 1

        # Add non-favorited buttons starting on the row after the delimiter
        for link in non_favorites:
            local_extensions = ('.exe', '.txt', '.pdf', '.docx', '.xlsx', '.xlsm', '.bat', '.dotx', '.py')
            url = link["url"].strip()
            url_normalized = os.path.normpath(url).lower()
            is_url = url_normalized.startswith(("http://", "https://"))
            url_basename = os.path.basename(url_normalized)
            is_local_file = not is_url and (any(url_basename.endswith(ext) for ext in local_extensions) or any(url_normalized.endswith(ext) for ext in local_extensions))

            btn_widget = LinkButtonWidget(
                link,
                self.launch_local_app if is_local_file else self.open_website,
                self.edit_link,
                self
            )
            if col >= buttons_per_row:
                col = 0
                row += 1
            grid_layout.addWidget(btn_widget, row, col)
            col += 1

        if search_widget:
            search_widget.blockSignals(False)
        grid_widget.update()  # Force UI refresh

    def sort_links(self, order):
        self.sort_order = order
        self.filter_links(self.search_query)

    def launchpad_tab(self):
        widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setSpacing(5)
        main_layout.setContentsMargins(0, 0, 0, 0)
        control_layout = QHBoxLayout()
        search = QLineEdit()
        search.setPlaceholderText("Search links...")
        search.setStyleSheet("""
            QLineEdit {
                background-color: #444;
                color: #fff;
                padding: 8px;
                border: 1px solid #555;
                border-radius: 5px;
                font-size: 16px;
            }
        """)
        try:
            search.textChanged.disconnect()
        except TypeError:
            pass
        search.textChanged.connect(self.filter_links)
        control_layout.addWidget(search)

        sort_combo = QComboBox()
        sort_combo.addItems(["Sort", "Alphabetical (A-Z)", "Alphabetical (Z-A)"])
        sort_combo.setStyleSheet("""
            QComboBox {
                background-color: #444;
                color: white;
                font-size: 16px;
                padding: 8px;
                border: 1px solid #555;
                border-radius: 5px;
                min-height: 20px;
            }
            QComboBox::drop-down {
                width: 20px;
                border: none;
                background: transparent;
            }
            QComboBox QAbstractItemView {
                background-color: #333;
                color: white;
                selection-background-color: #555;
            }
        """)
        sort_combo.setCurrentText("Alphabetical (A-Z)")  # Default to A-Z
        try:
            sort_combo.currentTextChanged.disconnect()
        except TypeError:
            pass
        sort_combo.currentTextChanged.connect(self.sort_links)
        control_layout.addWidget(sort_combo)
        control_layout.addStretch()
        main_layout.addLayout(control_layout)

        grid_widget = QWidget()
        grid_widget.setObjectName("grid_widget")
        grid_layout = QGridLayout()
        grid_layout.setHorizontalSpacing(5)
        grid_layout.setVerticalSpacing(5)
        grid_layout.setContentsMargins(2, 2, 2, 2)
        grid_widget.setLayout(grid_layout)

        container = QWidget()
        container_layout = QHBoxLayout()
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.addWidget(grid_widget)
        container.setLayout(container_layout)
        main_layout.addWidget(container)
        main_layout.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Minimum, QSizePolicy.Expanding))
        widget.setLayout(main_layout)

        self.launchpad_widget = widget
        # Explicitly call filter_links with an empty query to ensure favorites are positioned correctly
        self.filter_links("")
        return widget

    def add_new_link(self):
        dialog = EditLinkDialog(self, is_new=True)
        if dialog.exec_():
            new_link = dialog.get_data()
            if new_link:
                self.links.append(new_link)
                self.save_links()
                self.refresh_launchpad_tab()

    def edit_link(self, link):
        dialog = EditLinkDialog(self, link, is_new=False)
        if dialog.exec_():
            new_link = dialog.get_data()
            if new_link is None:
                self.links.remove(link)
            else:
                index = self.links.index(link)
                self.links[index] = new_link
            self.save_links()
            self.save_user_preferences()
            self.filter_links(self.search_query)

    def launch_taka_helper(self):
        try:
            subprocess.Popen([sys.executable, "TakaHelper.py"])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to launch Taka Helper: {str(e)}")

    def refresh_launchpad_tab(self):
        current_query = self.search_query
        current_sort = self.sort_order
        # Disconnect existing textChanged signal if it exists
        search_widget = self.launchpad_widget.findChild(QLineEdit)
        if search_widget:
            try:
                search_widget.textChanged.disconnect()
            except TypeError:
                pass  # No connection to disconnect
        self.tabs.removeTab(0)
        self.launchpad_widget = self.launchpad_tab()
        self.tabs.insertTab(0, self.launchpad_widget, "Launchpad")
        self.tabs.setCurrentIndex(0)
        # Restore search query without triggering signal
        search_widget = self.launchpad_widget.findChild(QLineEdit)
        if search_widget:
            search_widget.blockSignals(True)
            search_widget.setText(current_query)
            search_widget.blockSignals(False)
        self.search_query = current_query
        self.sort_order = current_sort

    def send_oncall_email(self):
        dialog = PageOnCallDialog(self)
        if dialog.exec_():
            email_data = dialog.get_email_data()
            try:
                import win32com.client
                outlook = win32com.client.Dispatch("Outlook.Application")
                mail = outlook.CreateItem(0)  # 0 = MailItem
                mail.To = email_data['to']
                mail.Subject = email_data['subject']
                mail.Body = email_data['body']
                mail.Send()
                QMessageBox.information(self, "Success", "OnCall email sent successfully via Outlook.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to send email via Outlook: {str(e)}")

    def show_user_guide(self):
            dialog = HelpDialog(self)
            dialog.exec_()

    def show_about(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("About LaunchPad")
        msg.setText("LaunchPad Version 2.0\n"
                    "An IT Hub for quick access to tools, commands, and guides.\n\n"
                    "Developed by: Enda Rensing\n")
        msg.setIcon(QMessageBox.NoIcon)  # No icon, no sound
        msg.exec_()

    def show_version_info(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("Version Info")
        msg.setText("LaunchPad Version: 2.0.0\n"
                    "Release Date: May 24, 2025\n"
                    "Built with: Python 3, PyQt5, pywin32")
        msg.setIcon(QMessageBox.NoIcon)  # No icon, no sound
        msg.exec_()

    def show_changelog(self):
        dialog = ChangelogDialog(self)
        dialog.exec_()

    def refresh_shared_files(self):
        if self.settings.get("mode") == "Shared":
            self.load_links()
            self.load_commands()
            self.load_guides()
            # Clear cmdlet_syntax_cache in NewCommandDialog instances
            for dialog in QApplication.topLevelWidgets():
                if isinstance(dialog, NewCommandDialog):
                    dialog.cmdlet_syntax_cache.clear()
                    dialog.help_repo = dialog.load_help_repository()
            self.refresh_launchpad_tab()
            self.tabs.removeTab(2)
            self.tabs.insertTab(2, self.howto_guides_tab(), "How-To Guides")
            self.tabs.removeTab(1)
            self.tabs.insertTab(1, self.commands_tab(), "Commands")
            QMessageBox.information(self, "Refresh", "Shared files reloaded.")

    def launch_local_app(self, app_name, browser=None):
        try:
            # Normalize the path to handle UNC paths and forward/backward slashes
            app_name = os.path.normpath(app_name)
            # Special handling for .py files to run with Python
            if app_name.lower().endswith('.py'):
                subprocess.Popen([sys.executable, app_name], shell=True)
            else:
                subprocess.Popen(app_name, shell=True)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to launch {app_name}: {str(e)}")

    def open_website(self, url, browser=None):
        try:
            if not url.startswith(("http://", "https://")):
                url = f"https://{url}"
            if browser and browser in self.browser_paths:
                subprocess.Popen([self.browser_paths[browser], url], shell=False)
            else:
                webbrowser.open(url)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open website: {str(e)}")

    def open_settings_dialog(self):
        dialog = SettingsDialog(self.settings, self)
        if dialog.exec_():
            new_settings = dialog.get_settings()
            if new_settings != self.settings:
                self.settings = new_settings
                self.save_settings()
                self.commands_file = self.get_file_path("commands.json")
                self.links_file = self.get_file_path("launchpad_links.json")
                self.guides_file = self.get_file_path("howto_guides.json")
                self.load_commands()
                self.load_links()
                self.load_guides()
                self.refresh_launchpad_tab()
                self.tabs.removeTab(2)
                self.tabs.insertTab(2, self.howto_guides_tab(), "How-To Guides")
                self.tabs.removeTab(1)
                self.tabs.insertTab(1, self.commands_tab(), "Commands")
                self.tabs.setCurrentIndex(0)

    def get_all_aliases(self):
        help_file = self.get_file_path("powershell_help.json")
        try:
            with open(help_file, 'r', encoding='utf-8-sig') as f:
                help_data = json.load(f)
            aliases = []
            for cmdlet, data in help_data["cmdlets"].items():
                aliases.extend(data.get("aliases", []))
            return aliases
        except:
            return []

    def open_add_example_dialog(self):
        dialog = AddExampleDialog(self, main_app=self)
        if dialog.exec_():
            data = dialog.get_data()
            help_file = self.get_file_path("powershell_help.json")
            if self.acquire_lock(help_file):
                try:
                    with open(help_file, 'r', encoding='utf-8-sig') as f:
                        help_data = json.load(f)

                    # Parse cmdlet and pipeline
                    cmdlet = data["cmdlet"]
                    parts = [part.strip() for part in cmdlet.split("|")]
                    main_cmdlet = parts[0]
                    pipeline_cmdlet = parts[1] if len(parts) > 1 else None

                    # Initialize structure if not present
                    if main_cmdlet not in help_data["cmdlets"]:
                        help_data["cmdlets"][main_cmdlet] = {
                            "aliases": [],
                            "commands": [],
                            "pipelines": {},
                            "fullPipelines": []
                        }

                    # Save standalone or pipeline examples
                    if pipeline_cmdlet:
                        if pipeline_cmdlet not in help_data["cmdlets"][main_cmdlet]["pipelines"]:
                            help_data["cmdlets"][main_cmdlet]["pipelines"][pipeline_cmdlet] = []
                        help_data["cmdlets"][main_cmdlet]["pipelines"][pipeline_cmdlet].extend(data["examples"])
                    else:
                        help_data["cmdlets"][main_cmdlet]["commands"].extend(data["examples"])
                        help_data["cmdlets"][main_cmdlet]["aliases"] = list(set(
                            help_data["cmdlets"][main_cmdlet]["aliases"] + data["aliases"]
                        ))

                    # Save full pipeline example
                    if data["full_pipeline"]:
                        full_pipeline = data["full_pipeline"]
                        if "fullPipelines" not in help_data["cmdlets"][main_cmdlet]:
                            help_data["cmdlets"][main_cmdlet]["fullPipelines"] = []
                        help_data["cmdlets"][main_cmdlet]["fullPipelines"].append(full_pipeline)

                    # Write to file
                    with open(help_file, 'w', encoding='utf-8') as f:
                        json.dump(help_data, f, indent=4)

                    QMessageBox.information(self, "Success", f"Added examples to {cmdlet}.")
                    self.refresh_help_cache()

                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to save example: {str(e)}")
                finally:
                    self.release_lock(help_file)
            else:
                QMessageBox.warning(self, "Error", "Could not save: powershell_help.json is locked.")

    def refresh_help_cache(self):
            for dialog in QApplication.topLevelWidgets():
                if isinstance(dialog, NewCommandDialog):
                    dialog.cmdlet_syntax_cache.clear()
                    dialog.help_repo = dialog.load_help_repository()

    def commands_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        top_layout = QHBoxLayout()
        search = QLineEdit()
        search.setPlaceholderText("Search commands...")
        search.setStyleSheet("""
            QLineEdit {
                background-color: #444;
                color: #fff;
                padding: 8px;
                border: 1px solid #555;
                border-radius: 5px;
                font-size: 16px;
            }
        """)
        top_layout.addWidget(search)
        import_btn = QPushButton("Import Commands")
        import_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                font-size: 14px;
                padding: 5px;
                border-radius: 3px;
                border: 1px solid #555;
            }
            QPushButton:hover { background-color: #005ba1; }
        """)
        top_layout.addWidget(import_btn)
        export_btn = QPushButton("Export Selected")
        export_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                font-size: 14px;
                padding: 5px;
                border-radius: 3px;
                border: 1px solid #555;
            }
            QPushButton:hover { background-color: #005ba1; }
        """)
        top_layout.addWidget(export_btn)
        new_command_btn = QPushButton("New Command")
        new_command_btn.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                font-size: 14px;
                padding: 5px;
                border-radius: 3px;
                border: 1px solid #555;
            }
            QPushButton:hover { background-color: #218838; }
        """)
        top_layout.addWidget(new_command_btn)
        layout.addLayout(top_layout)
        split_layout = QVBoxLayout()
        select_all_layout = QHBoxLayout()
        self.select_all_check = QCheckBox("Select All")
        self.select_all_check.setStyleSheet("""
            QCheckBox {
                color: #fff;
                font-size: 14px;
                padding: 5px;
            }
        """)
        select_all_layout.addWidget(self.select_all_check)
        select_all_layout.addStretch()
        split_layout.addLayout(select_all_layout)
        self.command_list = QListWidget()
        self.command_list.setSelectionMode(QListWidget.ExtendedSelection)  # Changed to allow Ctrl/Shift multi-selection
        self.command_list.setStyleSheet("""
            QListWidget {
                background-color: #333;
                color: #fff;
                border: 1px solid #444;
                font-size: 16px;
                padding: 5px;
            }
            QListWidget::item:hover {
                background-color: #444;
            }
            QListWidget::item:selected {
                background-color: #555;
            }
        """)
        split_layout.addWidget(self.command_list)
        self.details_panel = QTextEdit()
        self.details_panel.setReadOnly(True)
        self.details_panel.setStyleSheet("""
            QTextEdit {
                background-color: #333;
                color: #cccccc;
                border: 1px solid #444;
                font-size: 16px;
                padding: 5px;
            }
        """)
        split_layout.addWidget(self.details_panel)
        buttons_layout = QHBoxLayout()
        self.run_button = QPushButton("Run Commands")
        self.run_button.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                font-size: 14px;
                padding: 5px;
                border-radius: 3px;
                border: 1px solid #555;
            }
            QPushButton:hover { background-color: #005ba1; }
        """)
        self.run_button.setEnabled(False)
        buttons_layout.addWidget(self.run_button)
        self.edit_button = QPushButton("Edit Command")
        self.edit_button.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                font-size: 14px;
                padding: 5px;
                border-radius: 3px;
                border: 1px solid #555;
            }
            QPushButton:hover { background-color: #005ba1; }
        """)
        self.edit_button.setEnabled(False)
        buttons_layout.addWidget(self.edit_button)
        self.delete_button = QPushButton("Delete Selected")
        self.delete_button.setStyleSheet("""
            QPushButton {
                background-color: #d9534f;
                color: white;
                font-size: 14px;
                padding: 5px;
                border-radius: 3px;
                border: 1px solid #555;
            }
            QPushButton:hover { background-color: #c9302c; }
        """)
        self.delete_button.setEnabled(False)
        buttons_layout.addWidget(self.delete_button)
        split_layout.addLayout(buttons_layout)
        layout.addLayout(split_layout)
        self.current_run_connection = None

        def toggle_select_all():
            check_state = Qt.Checked if self.select_all_check.isChecked() else Qt.Unchecked
            for i in range(self.command_list.count()):
                item = self.command_list.item(i)
                item.setCheckState(check_state)

        def export_commands():
            selected_commands = []
            for i in range(self.command_list.count()):
                item = self.command_list.item(i)
                if item.checkState() == Qt.Checked:
                    command = item.data(Qt.UserRole)
                    selected_commands.append(command)
            if not selected_commands:
                QMessageBox.warning(widget, "No Selection", "Please check at least one command to export.")
                return
            file_name, _ = QFileDialog.getSaveFileName(widget, "Export Selected Commands", "", "JSON Files (*.json)")
            if file_name:
                try:
                    with open(file_name, 'w') as f:
                        json.dump(selected_commands, f, indent=4)
                    QMessageBox.information(widget, "Success", f"Exported {len(selected_commands)} commands successfully.")
                except Exception as e:
                    QMessageBox.critical(widget, "Error", f"Failed to export commands: {str(e)}")

        def import_commands():
            file_name, _ = QFileDialog.getOpenFileName(widget, "Import Commands", "", "JSON Files (*.json)")
            if file_name:
                try:
                    with open(file_name, 'r') as f:
                        imported_commands = json.load(f)
                    dialog = ImportCommandsDialog(imported_commands, widget)
                    if dialog.exec_():
                        selected_commands = dialog.get_selected_commands()
                        if not selected_commands:
                            QMessageBox.warning(widget, "No Selection", "No commands were selected to import.")
                            return
                        existing_titles = {command["title"] for command in self.commands}
                        new_commands = []
                        for command in selected_commands:
                            if command["title"] not in existing_titles:
                                if "pause_between_commands" not in command:
                                    command["pause_between_commands"] = False
                                new_commands.append(command)
                                existing_titles.add(command["title"])
                            else:
                                QMessageBox.warning(widget, "Duplicate Command", f"Command '{command['title']}' already exists and was skipped.")
                        self.commands.extend(new_commands)
                        populate_commands(self.commands)
                        self.save_commands()
                        QMessageBox.information(widget, "Success", f"Imported {len(new_commands)} new commands successfully.")
                except Exception as e:
                    QMessageBox.critical(widget, "Error", f"Failed to import commands: {str(e)}")

        def extract_commands(steps):
            recognized_commands = self.default_commands + self.custom_commands
            commands = []
            for step in steps:
                if step["type"] == "command":
                    cmd = step["content"].strip()
                    # Normalize cmd and check if it starts with any recognized command
                    cmd_lower = cmd.lower()
                    if cmd == "password.amazon.com":
                        commands.append(("open_website", cmd))
                    elif any(cmd_lower.startswith(recognized_cmd) for recognized_cmd in recognized_commands):
                        commands.append(("execute", cmd))
            return commands

        def populate_commands(commands_to_show):
            self.command_list.clear()
            for command in commands_to_show:
                item = QListWidgetItem(command["title"])
                item.setData(Qt.UserRole, command)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
                item.setTextAlignment(Qt.AlignLeft)
                self.command_list.addItem(item)

        def filter_commands():
            query = search.text().lower()
            filtered = [command for command in self.commands if query in command["title"].lower() or any(query in step["content"].lower() for step in command["steps"])]
            checked_titles = {self.command_list.item(i).text() for i in range(self.command_list.count()) if self.command_list.item(i).checkState() == Qt.Checked}
            populate_commands(filtered)
            for i in range(self.command_list.count()):
                item = self.command_list.item(i)
                if item.text() in checked_titles:
                    item.setCheckState(Qt.Checked)

        def show_details():
            if self.current_run_connection is not None:
                try:
                    self.run_button.clicked.disconnect(self.current_run_connection)
                except TypeError:
                    pass
                self.current_run_connection = None
            selected_items = self.command_list.selectedItems()
            if not selected_items:
                self.details_panel.setText("")
                self.run_button.setEnabled(False)
                self.edit_button.setEnabled(False)
                self.delete_button.setEnabled(False)
                return
            self.edit_button.setEnabled(len(selected_items) == 1)
            checked_items = [self.command_list.item(i) for i in range(self.command_list.count()) if self.command_list.item(i).checkState() == Qt.Checked]
            self.delete_button.setEnabled(bool(selected_items) or bool(checked_items))
            item = selected_items[0]
            command = item.data(Qt.UserRole)
            commands = extract_commands(command["steps"])
            desc_display = f"<b>{command['title']}</b><br><br>"
            for i, step in enumerate(command["steps"], 1):
                content = step["content"]  # Raw content
                step_type = "Command" if step["type"] == "command" else "Output"
                content = content.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")  # Escape HTML
                if step["type"] == "command":
                    content = f"<span style='color: #66b3ff;'>{content}</span>"
                desc_display += f"{i}. {step_type}: {content}<br>"
                desc_display += f"Delay: {step['delay']}ms<br><br>"
            self.details_panel.setHtml(desc_display)
            self.run_button.setEnabled(bool(commands))
            if commands:
                def run_commands():
                    run_all_commands(commands, command)
                    self.command_list.clearSelection()
                    self.command_list.setCurrentItem(item)
                self.current_run_connection = run_commands
                self.run_button.clicked.connect(run_commands)

        def run_all_commands(commands, command):
            shell = command.get("shell", "CMD")
            elevated = command.get("elevated", False)
            pause_between_commands = command.get("pause_between_commands", False)
            cmd_sequence = []
            placeholders = []
            output_file = None
            for step in command["steps"]:
                content = step["content"]
                if step["type"] == "command":
                    matches = re.findall(r'<([^|>]+)(?:\|validate_service)?>', content)
                    for match in matches:
                        full_placeholder = f"{match}|validate_service" if f"<{match}|validate_service>" in content else match
                        if match.lower() not in ["whid", "domain"] and full_placeholder not in placeholders:
                            placeholders.append(full_placeholder)
                        if match == "output_path" and "<output_path>" in content:
                            output_file = content.split("<output_path>")[1].strip()
            if placeholders:
                dialog = MultiInputDialog(placeholders, widget)
                if not dialog.exec_():
                    return
                values = dialog.get_values()
                for i, step in enumerate(command["steps"]):
                    content = step["content"]
                    for orig_ph in placeholders:
                        ph = orig_ph.split("|")[0]
                        if ph in values:
                            placeholder = f"<{orig_ph}>"
                            command["steps"][i]["content"] = content.replace(placeholder, values[ph])
                            content = command["steps"][i]["content"]
                            if orig_ph == "output_path" and values[ph]:
                                output_file = values[ph]
            for i, step in enumerate(command["steps"]):
                cmd_type = step["type"]
                content = step["content"]
                delay = step["delay"] / 1000.0
                if cmd_type == "command":
                    if "<domain>" in content.lower():
                        domain, ok = QInputDialog.getText(widget, "Domain Input", f"Enter domain or IP for {content}:")
                        if not ok or not domain.strip():
                            QMessageBox.warning(widget, "Input Error", "A valid domain or IP is required.")
                            return
                        content = content.replace("<domain>", domain.strip()).replace("<DOMAIN>", domain.strip())
                    if "<whid>" in content.lower():
                        whid, ok = QInputDialog.getText(widget, "WHID Input", f"Enter WHID for {content}:")
                        if not ok or not whid.strip():
                            QMessageBox.warning(widget, "Input Error", "A valid WHID is required.")
                            return
                        content = content.replace("<whid>", whid.strip()).replace("<WHID>", whid.strip())
                    content = os.path.expandvars(content)
                    if content == "password.amazon.com":
                        self.open_website(content)
                        continue
                    cmd_sequence.append(content)
                    if delay > 0:
                        if shell == "CMD":
                            cmd_sequence.append(f"timeout /t {int(delay)} /nobreak >nul")
                        elif shell == "PowerShell":
                            cmd_sequence.append(f"Start-Sleep -Milliseconds {int(delay * 1000)}")
                        else:
                            time.sleep(delay)
                elif cmd_type == "output":
                    content = os.path.expandvars(content)
                    if shell == "CMD":
                        cmd_sequence.append(f"echo {content}")
                    elif shell == "PowerShell":
                        escaped_content = content.replace('"', '`"')
                        cmd_sequence.append(f"Write-Output \"{escaped_content}\"")
                    else:
                        continue
                    if delay > 0:
                        if shell == "CMD":
                            cmd_sequence.append(f"timeout /t {int(delay)} /nobreak >nul")
                        elif shell == "PowerShell":
                            cmd_sequence.append(f"Start-Sleep -Milliseconds {int(delay * 1000)}")
                        else:
                            time.sleep(delay)
                if pause_between_commands and i < len(command["steps"]) - 1:
                    if shell == "CMD":
                        cmd_sequence.append("pause")
                    elif shell == "PowerShell":
                        cmd_sequence.append("Read-Host -Prompt \"Press Enter to continue\"")
            if not cmd_sequence:
                return
            if shell == "CMD":
                full_command = " & ".join(cmd_sequence)
                shell_cmd = ["cmd", "/c", full_command]
                shell_flag = False
                creationflags = subprocess.CREATE_NEW_CONSOLE
            elif shell == "PowerShell":
                cmd_sequence_adjusted = []
                for cmd in cmd_sequence:
                    if cmd.lower().startswith("invoke-command"):
                        cmd_sequence_adjusted.append("Write-Output 'Invoke-Command requires a -ScriptBlock, e.g., Invoke-Command -ScriptBlock { Write-Output \"Hello\" }'")
                    elif cmd.lower().startswith("write-output"):
                        cmd_sequence_adjusted.append(f"{cmd} | Out-String")
                    else:
                        cmd_sequence_adjusted.append(f"{cmd} | Out-String")
                full_command = "; ".join(cmd_sequence_adjusted)
                resize_cmd = "[Console]::WindowWidth=120; [Console]::WindowHeight=30; [Console]::BufferWidth=120"
                error_handling = "trap { Write-Error $_.Exception.Message; continue }"
                shell_cmd = ["powershell", "-Command", f"{error_handling}; {resize_cmd}; {full_command}"]
                shell_flag = True
                creationflags = subprocess.CREATE_NEW_CONSOLE
            else:
                shell_cmd = ["wt"]
                shell_flag = False
                creationflags = subprocess.CREATE_NEW_CONSOLE
            if elevated:
                if shell == "CMD":
                    shell_cmd = ["powershell", "-Command", f"Start-Process cmd -ArgumentList '/c {full_command}' -Verb RunAs"]
                    shell_flag = True
                    creationflags = 0
                elif shell == "PowerShell":
                    shell_cmd = ["powershell", "-Command", f"Start-Process -FilePath 'powershell.exe' -ArgumentList '-Command', '{full_command}' -Verb RunAs"]
                    shell_flag = True
                    creationflags = 0
                elif shell == "Terminal":
                    shell_cmd = ["powershell", "-Command", "Start-Process -FilePath wt -Verb RunAs"]
                    shell_flag = True
                    creationflags = 0
            try:
                process = subprocess.run(shell_cmd, shell=shell_flag, check=True)
                if output_file and os.path.exists(output_file):
                    os.startfile(output_file)
                self.command_list.clearSelection()  # Reset selection after run
            except Exception as e:
                QMessageBox.critical(widget, "Error", f"Failed to execute command sequence: {str(e)}")
                return
            
        def add_new_command():
            dialog = NewCommandDialog(widget, main_app=self, default_commands=self.default_commands, custom_commands=self.custom_commands)
            if dialog.exec_():
                new_command = dialog.command
                self.commands.append(new_command)
                populate_commands(self.commands)
                for i in range(self.command_list.count()):
                    item = self.command_list.item(i)
                    if item.text() == new_command["title"]:
                        self.command_list.setCurrentItem(item)
                        break

        def edit_command():
            selected_items = self.command_list.selectedItems()
            if len(selected_items) != 1:
                QMessageBox.warning(widget, "Selection Error", "Please select exactly one command to edit.")
                return
            item = selected_items[0]
            command = item.data(Qt.UserRole)
            dialog = NewCommandDialog(widget, main_app=self, command=command, default_commands=self.default_commands, custom_commands=self.custom_commands)
            if dialog.exec_():
                updated_command = dialog.command
                command_index = self.commands.index(command)
                self.commands[command_index] = updated_command
                populate_commands(self.commands)
                for i in range(self.command_list.count()):
                    item = self.command_list.item(i)
                    if item.text() == updated_command["title"]:
                        self.command_list.setCurrentItem(item)
                        break

        def delete_command():
            # Collect all checked commands
            checked_items = [self.command_list.item(i) for i in range(self.command_list.count()) if self.command_list.item(i).checkState() == Qt.Checked]
            if not checked_items:
                QMessageBox.warning(widget, "No Selection", "Please check at least one command to delete.")
                return
            if QMessageBox.question(widget, "Confirm Delete", f"Are you sure you want to delete {len(checked_items)} command(s)?") == QMessageBox.Yes:
                for item in checked_items:
                    command = item.data(Qt.UserRole)
                    self.commands.remove(command)
                populate_commands(self.commands)
                self.details_panel.setText("")
                self.run_button.setEnabled(False)
                self.edit_button.setEnabled(False)
                self.delete_button.setEnabled(False)
                self.select_all_check.setChecked(False)

        self.select_all_check.stateChanged.connect(toggle_select_all)
        search.textChanged.connect(filter_commands)
        self.command_list.itemSelectionChanged.connect(show_details)
        new_command_btn.clicked.connect(add_new_command)
        self.edit_button.clicked.connect(edit_command)
        self.delete_button.clicked.connect(delete_command)
        export_btn.clicked.connect(export_commands)
        import_btn.clicked.connect(import_commands)
        populate_commands(self.commands)
        widget.setLayout(layout)
        return widget

    def howto_guides_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        top_layout = QHBoxLayout()
        search = QLineEdit()
        search.setPlaceholderText("Search guides...")
        search.setStyleSheet("""
            QLineEdit {
                background-color: #444;
                color: #fff;
                padding: 8px;
                border: 1px solid #555;
                border-radius: 5px;
                font-size: 16px;
            }
        """)
        top_layout.addWidget(search)
        import_btn = QPushButton("Import Guides")
        import_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                font-size: 14px;
                padding: 5px;
                border-radius: 3px;
                border: 1px solid #555;
            }
            QPushButton:hover { background-color: #005ba1; }
        """)
        top_layout.addWidget(import_btn)
        export_btn = QPushButton("Export Selected")
        export_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                font-size: 14px;
                padding: 5px;
                border-radius: 3px;
                border: 1px solid #555;
            }
            QPushButton:hover { background-color: #005ba1; }
        """)
        top_layout.addWidget(export_btn)
        new_guide_btn = QPushButton("New Guide")
        new_guide_btn.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                font-size: 14px;
                padding: 5px;
                border-radius: 3px;
                border: 1px solid #555;
            }
            QPushButton:hover { background-color: #218838; }
        """)
        top_layout.addWidget(new_guide_btn)
        layout.addLayout(top_layout)
        split_layout = QVBoxLayout()
        select_all_layout = QHBoxLayout()
        self.guide_select_all_check = QCheckBox("Select All")
        self.guide_select_all_check.setStyleSheet("""
            QCheckBox {
                color: #fff;
                font-size: 14px;
                padding: 5px;
            }
        """)
        select_all_layout.addWidget(self.guide_select_all_check)
        select_all_layout.addStretch()
        split_layout.addLayout(select_all_layout)
        self.guide_list = QListWidget()
        self.guide_list.setSelectionMode(QListWidget.SingleSelection)
        self.guide_list.setStyleSheet("""
            QListWidget {
                background-color: #333;
                color: #fff;
                border: 1px solid #444;
                font-size: 16px;
                padding: 5px;
            }
            QListWidget::item:hover {
                background-color: #444;
            }
            QListWidget::item:selected {
                background-color: #555;
            }
        """)
        split_layout.addWidget(self.guide_list)
        self.guide_details_panel = QTextBrowser()
        self.guide_details_panel.setReadOnly(True)
        self.guide_details_panel.setStyleSheet("""
            QTextBrowser {
                background-color: #333;
                color: #cccccc;
                border: 1px solid #444;
                font-size: 16px;
                padding: 5px;
            }
        """)
        split_layout.addWidget(self.guide_details_panel)
        buttons_layout = QHBoxLayout()
        self.guide_edit_button = QPushButton("Edit Guide")
        self.guide_edit_button.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                font-size: 14px;
                padding: 5px;
                border-radius: 3px;
                border: 1px solid #555;
            }
            QPushButton:hover { background-color: #005ba1; }
        """)
        self.guide_edit_button.setEnabled(False)
        buttons_layout.addWidget(self.guide_edit_button)
        self.guide_delete_button = QPushButton("Delete Guide")
        self.guide_delete_button.setStyleSheet("""
            QPushButton {
                background-color: #d9534f;
                color: white;
                font-size: 14px;
                padding: 5px;
                border-radius: 3px;
                border: 1px solid #555;
            }
            QPushButton:hover { background-color: #c9302c; }
        """)
        self.guide_delete_button.setEnabled(False)
        buttons_layout.addWidget(self.guide_delete_button)
        split_layout.addLayout(buttons_layout)
        layout.addLayout(split_layout)

        def toggle_select_all():
            check_state = Qt.Checked if self.guide_select_all_check.isChecked() else Qt.Unchecked
            for i in range(self.guide_list.count()):
                item = self.guide_list.item(i)
                item.setCheckState(check_state)

        def export_guides():
            selected_guides = []
            for i in range(self.guide_list.count()):
                item = self.guide_list.item(i)
                if item.checkState() == Qt.Checked:
                    guide = item.data(Qt.UserRole)
                    selected_guides.append(guide)
            if not selected_guides:
                QMessageBox.warning(widget, "No Selection", "Please check at least one guide to export.")
                return
            file_name, _ = QFileDialog.getSaveFileName(widget, "Export Selected Guides", "", "JSON Files (*.json)")
            if file_name:
                try:
                    with open(file_name, 'w') as f:
                        json.dump(selected_guides, f, indent=4)
                    QMessageBox.information(widget, "Success", f"Exported {len(selected_guides)} guides successfully.")
                except Exception as e:
                    QMessageBox.critical(widget, "Error", f"Failed to export guides: {str(e)}")

        def import_guides():
            file_name, _ = QFileDialog.getOpenFileName(widget, "Import Guides", "", "JSON Files (*.json)")
            if file_name:
                try:
                    with open(file_name, 'r') as f:
                        imported_guides = json.load(f)
                    for guide in imported_guides:
                        if isinstance(guide["steps"], list):
                            guide["steps"] = '<ol style="font-size: 16px;">' + "".join(f"<li>{step}</li>" for step in guide["steps"]) + "</ol>"
                        if not guide["description"].startswith("<"):
                            guide["description"] = f'<p style="font-size: 16px;">{guide["description"]}</p>'
                    dialog = ImportGuidesDialog(imported_guides, widget)
                    if dialog.exec_():
                        selected_guides = dialog.get_selected_guides()
                        if not selected_guides:
                            QMessageBox.warning(widget, "No Selection", "No guides were selected to import.")
                            return
                        existing_titles = {guide["title"] for guide in self.guides}
                        new_guides = []
                        for guide in selected_guides:
                            if guide["title"] not in existing_titles:
                                new_guides.append(guide)
                                existing_titles.add(guide["title"])
                            else:
                                QMessageBox.warning(widget, "Duplicate Guide", f"Guide '{guide['title']}' already exists and was skipped.")
                        self.guides.extend(new_guides)
                        populate_guides(self.guides)
                        self.save_guides()
                        QMessageBox.information(widget, "Success", f"Imported {len(new_guides)} new guides successfully.")
                except Exception as e:
                    QMessageBox.critical(widget, "Error", f"Failed to import guides: {str(e)}")

        def populate_guides(guides_to_show):
            self.guide_list.clear()
            for guide in guides_to_show:
                item = QListWidgetItem(guide["title"])
                item.setData(Qt.UserRole, guide)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
                item.setTextAlignment(Qt.AlignLeft)
                self.guide_list.addItem(item)

        def filter_guides():
            query = search.text().lower()
            filtered = []
            for guide in self.guides:
                if query in guide["title"].lower() or query in guide["description"].lower():
                    filtered.append(guide)
                    continue
                steps_text = re.sub(r'<[^>]+>', '', guide["steps"]).lower()
                if query in steps_text:
                    filtered.append(guide)
            checked_titles = {self.guide_list.item(i).text() for i in range(self.guide_list.count()) if self.guide_list.item(i).checkState() == Qt.Checked}
            populate_guides(filtered)
            for i in range(self.guide_list.count()):
                item = self.guide_list.item(i)
                if item.text() in checked_titles:
                    item.setCheckState(Qt.Checked)

        def make_urls_clickable(text):
            text = re.sub(r'<!DOCTYPE[^>]*>', '', text, flags=re.IGNORECASE)
            url_pattern = r'(?<!href=")(https?://[^\s<"]+|www\.[^\s<"]+)'
            def replace_url(match):
                url = match.group(0)
                if not url.startswith('http'):
                    url = 'https://' + url
                return f'<a href="{url}" style="color: #66b3ff;">{match.group(0)}</a>'
            result = re.sub(url_pattern, replace_url, text)
            return result

        def show_guide_details():
            selected_items = self.guide_list.selectedItems()
            if not selected_items:
                self.guide_details_panel.setHtml("")
                self.guide_edit_button.setEnabled(False)
                self.guide_delete_button.setEnabled(False)
                return
            item = selected_items[0]
            guide = item.data(Qt.UserRole)
            description_with_links = make_urls_clickable(guide['description'])
            steps_with_links = make_urls_clickable(guide['steps'])
            desc_display = f"<b>{guide['title']}</b><br><br>"
            desc_display += f"<b>Description:</b><br>{description_with_links}<br><br>"
            desc_display += f"<b>Steps:</b><br>{steps_with_links}"
            self.guide_details_panel.setHtml(desc_display)
            self.guide_details_panel.setOpenExternalLinks(True)
            self.guide_edit_button.setEnabled(len(selected_items) == 1)
            self.guide_delete_button.setEnabled(True)

        def add_new_guide():
            dialog = NewGuideDialog(widget)
            if dialog.exec_():
                new_guide = dialog.guide_data
                self.guides.append(new_guide)
                populate_guides(self.guides)
                self.save_guides()
                for i in range(self.guide_list.count()):
                    item = self.guide_list.item(i)
                    if item.text() == new_guide["title"]:
                        self.guide_list.setCurrentItem(item)
                        break

        def edit_guide():
            selected_items = self.guide_list.selectedItems()
            if len(selected_items) != 1:
                QMessageBox.warning(widget, "Selection Error", "Please select one guide to edit.")
                return
            item = selected_items[0]
            guide = item.data(Qt.UserRole)
            dialog = NewGuideDialog(widget, guide)
            if dialog.exec_():
                updated_guide = dialog.guide_data
                guide_index = self.guides.index(guide)
                self.guides[guide_index] = updated_guide
                populate_guides(self.guides)
                self.save_guides()
                for i in range(self.guide_list.count()):
                    item = self.guide_list.item(i)
                    if item.text() == updated_guide["title"]:
                        self.guide_list.setCurrentItem(item)
                        break

        def delete_guide():
            selected_items = self.guide_list.selectedItems()
            if not selected_items:
                QMessageBox.warning(widget, "No Selection", "Please select at least one guide to delete.")
                return
            if QMessageBox.question(widget, "Confirm Delete", f"Are you sure you want to delete {len(selected_items)} guide(s)?") == QMessageBox.Yes:
                for item in selected_items:
                    guide = item.data(Qt.UserRole)
                    self.guides.remove(guide)
                populate_guides(self.guides)
                self.save_guides()
                self.guide_details_panel.setHtml("")
                self.guide_edit_button.setEnabled(False)
                self.guide_delete_button.setEnabled(False)

        self.guide_select_all_check.stateChanged.connect(toggle_select_all)
        search.textChanged.connect(filter_guides)
        self.guide_list.itemSelectionChanged.connect(show_guide_details)
        new_guide_btn.clicked.connect(add_new_guide)
        self.guide_edit_button.clicked.connect(edit_guide)
        self.guide_delete_button.clicked.connect(delete_guide)
        populate_guides(self.guides)
        widget.setLayout(layout)
        return widget

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec_())
