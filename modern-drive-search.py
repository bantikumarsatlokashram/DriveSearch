import sys
import os
import time
import shutil
import random
from datetime import datetime
import json

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QLineEdit, QPushButton, QToolButton, QTableWidget, QTableWidgetItem,
    QAbstractItemView, QLabel, QScrollArea, QFrame, QMessageBox,
    QFileDialog, QInputDialog, QProgressBar, QMenu
)
from PyQt5.QtCore import (
    Qt, QThread, pyqtSignal
)
from PyQt5.QtGui import (
    QIcon, QCursor, QColor, QBrush
)

import win32api
import subprocess

# For advanced shell features (Share, Properties, etc.)
try:
    import pythoncom
    import win32com.client
    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False

# For Recycle Bin deletes
try:
    import send2trash
    HAS_SEND2TRASH = True
except ImportError:
    HAS_SEND2TRASH = False

################################################################################
# Constants
################################################################################
THEME_FILE = "user_theme.json"
SYSTEM_FOLDERS = ["\\Windows", "\\Program Files", "\\ProgramData", "\\AppData"]

################################################################################
# Theme Persistence
################################################################################
def load_theme_preference():
    """Load theme ('light' or 'dark') from JSON file, default to 'light'."""
    if os.path.exists(THEME_FILE):
        try:
            with open(THEME_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data.get("theme", "light")
        except:
            return "light"
    return "light"

def save_theme_preference(theme_name):
    """Save theme preference to JSON file."""
    with open(THEME_FILE, "w", encoding="utf-8") as f:
        json.dump({"theme": theme_name}, f, indent=2)

################################################################################
# File Extension Colors
################################################################################
class FileTypeColors:
    def __init__(self):
        self.colors = {}
        self.type_colors = {
            'pdf': '#FF6B6B',
            'doc': '#4ECDC4',
            'docx': '#4ECDC4',
            'txt': '#95A5A6',
            'rtf': '#95A5A6',

            'jpg': '#45B7D1',
            'jpeg': '#45B7D1',
            'png': '#45B7D1',
            'gif': '#45B7D1',
            'bmp': '#45B7D1',

            'mp3': '#9B59B6',
            'wav': '#9B59B6',
            'mp4': '#8E44AD',
            'avi': '#8E44AD',
            'mkv': '#8E44AD',

            'zip': '#F39C12',
            'rar': '#F39C12',
            '7z': '#F39C12',

            'py': '#2ECC71',
            'js': '#2ECC71',
            'html': '#2ECC71',
            'css': '#2ECC71',
            'cpp': '#2ECC71',

            'exe': '#E74C3C',
            'msi': '#E74C3C',
        }

    def get_color(self, ext):
        lower_ext = ext.lower()
        if lower_ext not in self.colors:
            if lower_ext in self.type_colors:
                self.colors[lower_ext] = self.type_colors[lower_ext]
            else:
                # random pastel
                hue = random.random()
                saturation = 0.3 + random.random() * 0.2
                value = 0.9 + random.random() * 0.1
                color = QColor.fromHsvF(hue, saturation, value)
                self.colors[lower_ext] = color.name()
        return self.colors[lower_ext]


################################################################################
# SearchWorker: Single-Pass, "Live" Emission of Found Files
################################################################################
class SearchWorker(QThread):
    foundMatch = pyqtSignal(str, str, str, float, float) 
    # name, path, ext, size_bytes, modified_timestamp
    statusMsg = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, drives, search_term, skip_folders, parent=None):
        super().__init__(parent)
        self.drives = drives
        self.search_term = search_term.lower()
        self.skip_folders = skip_folders
        self._stopped = False

    def stop(self):
        self._stopped = True

    def run(self):
        self.statusMsg.emit("Starting live search...")
        # For each drive, skip system folders, check each file name
        for d in self.drives:
            if self._stopped:
                break
            self.statusMsg.emit(f"Scanning: {d}")
            for root, dirs, files in os.walk(d):
                if self._stopped:
                    break
                if self.should_skip(root):
                    continue
                for fname in files:
                    if self._stopped:
                        break
                    # Quick name match
                    if self.search_term in fname.lower():
                        full_path = os.path.join(root, fname)
                        try:
                            st = os.stat(full_path)
                            size_bytes = st.st_size
                            mod_time = st.st_mtime
                            ext = os.path.splitext(fname)[1][1:] or "FILE"
                            self.foundMatch.emit(fname, full_path, ext.upper(), size_bytes, mod_time)
                        except:
                            continue
        self.statusMsg.emit("Search complete.")
        self.finished.emit()

    def should_skip(self, path):
        for sf in self.skip_folders:
            if sf.lower() in path.lower():
                return True
        return False


################################################################################
# Main Window
################################################################################
class ModernDriveSearch(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Modern Drive Search")
        self.resize(1280, 720)

        self.current_theme = load_theme_preference()
        self.file_colors = FileTypeColors()

        self.selected_drive = None
        self.search_thread = None
        self.searching = False

        icon_path = os.path.join(os.path.dirname(__file__), "app_icon.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.init_ui()
        self.apply_theme(self.current_theme)

    def init_ui(self):
        # QSS
        self.light_qss = """
            QMainWindow {
                background-color: #f8f9fa;
            }
            QLineEdit {
                padding: 8px 12px;
                border: 1px solid #ced4da;
                border-radius: 20px;
                font-size: 14px;
                color: #495057;
            }
            QLineEdit:focus {
                border: 1px solid #7456f1;
            }
            QPushButton {
                background-color: #7456f1;
                color: white;
                border: none;
                border-radius: 16px;
                font-size: 14px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #5e46c2;
            }
            QToolButton {
                background: transparent;
                border: none;
                font-size: 14px;
                color: #495057;
                padding: 6px;
            }
            QToolButton:hover {
                color: #7456f1;
                background: rgba(116, 86, 241, 0.1);
                border-radius: 8px;
            }
            #SidebarFrame {
                background-color: #ffffff;
                border-right: 1px solid #dee2e6;
            }
            #DriveButton {
                background: transparent;
                border: none;
                text-align: left;
                padding: 10px 20px;
                font-size: 14px;
                color: #495057;
            }
            #DriveButton:hover {
                background: rgba(116, 86, 241, 0.1);
                color: #7456f1;
                border-radius: 8px;
            }
            QTableWidget {
                background-color: white;
                border: 1px solid #dee2e6;
                gridline-color: #dee2e6;
            }
            QHeaderView::section {
                background-color: #e9ecef;
                border: none;
                font-weight: bold;
            }
            QLabel {
                color: #495057;
            }
        """

        self.dark_qss = """
            QMainWindow {
                background-color: #1e1e1e;
            }
            QLineEdit {
                padding: 8px 12px;
                border: 1px solid #444;
                border-radius: 20px;
                font-size: 14px;
                color: #eee;
                background-color: #2c2c2c;
            }
            QLineEdit:focus {
                border: 1px solid #7456f1;
            }
            QPushButton {
                background-color: #7456f1;
                color: white;
                border: none;
                border-radius: 16px;
                font-size: 14px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #5e46c2;
            }
            QToolButton {
                background: transparent;
                border: none;
                font-size: 14px;
                color: #ccc;
                padding: 6px;
            }
            QToolButton:hover {
                color: #7456f1;
                background: rgba(116, 86, 241, 0.1);
                border-radius: 8px;
            }
            #SidebarFrame {
                background-color: #2c2c2c;
                border-right: 1px solid #444;
            }
            #DriveButton {
                background: transparent;
                border: none;
                text-align: left;
                padding: 10px 20px;
                font-size: 14px;
                color: #ccc;
            }
            #DriveButton:hover {
                background: rgba(116, 86, 241, 0.2);
                color: #7456f1;
                border-radius: 8px;
            }
            QTableWidget {
                background-color: #2c2c2c;
                border: 1px solid #444;
                color: #ccc;
                gridline-color: #444;
            }
            QHeaderView::section {
                background-color: #333;
                color: #ccc;
                border: none;
                font-weight: bold;
            }
            QLabel {
                color: #ccc;
            }
        """

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(0,0,0,0)
        main_layout.setSpacing(0)

        # Left sidebar
        self.sidebar_frame = QFrame()
        self.sidebar_frame.setObjectName("SidebarFrame")
        self.sidebar_frame.setFixedWidth(240)
        sb_layout = QVBoxLayout(self.sidebar_frame)
        sb_layout.setContentsMargins(10,10,10,10)
        sb_layout.setSpacing(10)

        label_drives = QLabel("Drives")
        label_drives.setStyleSheet("font-weight: bold; font-size:16px;")
        sb_layout.addWidget(label_drives)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_container = QWidget()
        self.drives_vlayout= QVBoxLayout(scroll_container)
        self.drives_vlayout.setSpacing(5)
        self.drives_vlayout.setContentsMargins(0,0,0,0)

        self.populate_drive_list()

        scroll_area.setWidget(scroll_container)
        sb_layout.addWidget(scroll_area)

        sb_layout.addStretch()
        main_layout.addWidget(self.sidebar_frame)

        # Right
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(20,20,20,20)
        right_layout.setSpacing(10)

        # top bar
        top_bar = QHBoxLayout()
        self.search_line = QLineEdit()
        self.search_line.setPlaceholderText("Type a term, press Enter (live results)...")
        self.search_line.returnPressed.connect(self.on_search_clicked)

        self.search_btn=QPushButton("Search")
        self.search_btn.clicked.connect(self.on_search_clicked)

        self.dark_mode_btn=QToolButton()
        self.dark_mode_btn.setCheckable(True)
        if self.current_theme=="dark":
            self.dark_mode_btn.setText("Light Mode")
            self.dark_mode_btn.setChecked(True)
        else:
            self.dark_mode_btn.setText("Dark Mode")
        self.dark_mode_btn.clicked.connect(self.on_toggle_dark)

        top_bar.addWidget(self.search_line)
        top_bar.addWidget(self.search_btn)
        top_bar.addStretch()
        top_bar.addWidget(self.dark_mode_btn)

        right_layout.addLayout(top_bar)

        self.status_label=QLabel("Idle...")
        right_layout.addWidget(self.status_label)

        # Table
        self.file_table=QTableWidget()
        self.file_table.setColumnCount(5)
        self.file_table.setHorizontalHeaderLabels(["Type","Name","Path","Size","Modified"])
        self.file_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.file_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.file_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.file_table.horizontalHeader().setStretchLastSection(True)
        self.file_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.file_table.customContextMenuRequested.connect(self.show_context_menu)

        right_layout.addWidget(self.file_table,1)
        main_layout.addWidget(right_widget,1)

    def apply_theme(self,t):
        if t=="dark":
            self.setStyleSheet(self.dark_qss)
        else:
            self.setStyleSheet(self.light_qss)

    def on_toggle_dark(self):
        isChecked=self.dark_mode_btn.isChecked()
        if isChecked:
            self.dark_mode_btn.setText("Light Mode")
            self.current_theme="dark"
        else:
            self.dark_mode_btn.setText("Dark Mode")
            self.current_theme="light"
        save_theme_preference(self.current_theme)
        self.apply_theme(self.current_theme)

    ############################################
    # Drive list
    ############################################
    def populate_drive_list(self):
        drive_str=win32api.GetLogicalDriveStrings()
        letters=drive_str.split('\000')[:-1]
        for d in letters:
            btn=QToolButton()
            btn.setObjectName("DriveButton")
            try:
                vol_name=win32api.GetVolumeInformation(d)[0]
                if not vol_name:
                    vol_name="Local Disk"
            except:
                vol_name="Unknown"
            short = d.replace("\\","")
            btn.setText(f"{vol_name} ({short}:)")
            def drive_clicked(checked=False,dr=d):
                self.selected_drive=dr
                self.status_label.setText(f"Selected drive: {dr} (search restricted).")
                self.file_table.setRowCount(0)
            btn.clicked.connect(drive_clicked)
            self.drives_vlayout.addWidget(btn)
        self.drives_vlayout.addStretch()

    ############################################
    # Searching
    ############################################
    def on_search_clicked(self):
        # if we're currently searching, let's stop
        if self.searching:
            self.stop_search()
            return

        term = self.search_line.text().strip()
        if not term:
            QMessageBox.warning(self,"No term","Please type something.")
            return

        self.file_table.setRowCount(0)
        # decide which drives
        if self.selected_drive:
            drives=[self.selected_drive]
            self.status_label.setText(f"Searching drive: {self.selected_drive}")
        else:
            drive_str=win32api.GetLogicalDriveStrings()
            letters=drive_str.split('\000')[:-1]
            drives=letters
            self.status_label.setText("Searching all drives...")

        self.search_thread=SearchWorker(drives, term, SYSTEM_FOLDERS)
        self.search_thread.foundMatch.connect(self.add_live_match)
        self.search_thread.statusMsg.connect(self.update_status)
        self.search_thread.finished.connect(self.on_search_finished)
        self.search_thread.start()

        self.searching=True
        self.search_btn.setText("Stop")
        self.update_status("Starting live search...")

    def stop_search(self):
        if self.search_thread and self.search_thread.isRunning():
            self.search_thread.stop()
        self.search_thread=None
        self.searching=False
        self.search_btn.setText("Search")
        self.update_status("Search stopped.")

    def on_search_finished(self):
        self.searching=False
        self.search_thread=None
        self.search_btn.setText("Search")
        self.update_status("Done scanning.")

    def add_live_match(self, fname, full_path, ext, size_bytes, mod_time):
        # Insert a row for each match
        row_idx = self.file_table.rowCount()
        self.file_table.insertRow(row_idx)

        color = self.file_colors.get_color(ext.lower())
        row_data=[
            ext,
            fname,
            full_path,
            self.format_size(size_bytes),
            datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M:%S')
        ]
        for c, val in enumerate(row_data):
            item=QTableWidgetItem(val)
            if c==0:
                item.setForeground(QBrush(QColor(color)))
            self.file_table.setItem(row_idx,c,item)

    def update_status(self,msg):
        self.status_label.setText(msg)

    def format_size(self, size_bytes):
        size = float(size_bytes)
        for unit in ['B','KB','MB','GB']:
            if size<1024:
                return f"{size:.2f} {unit}"
            size/=1024
        return f"{size:.2f} TB"

    ############################################
    # Context Menu
    ############################################
    def show_context_menu(self,pos):
        row=self.file_table.currentRow()
        if row<0:
            return
        file_path = self.file_table.item(row,2).text() # path is column=2

        menu=QMenu(self)
        copy_act=menu.addAction("Copy")
        cut_act=menu.addAction("Cut")
        rename_act=menu.addAction("Rename")
        short_act=menu.addAction("Create Shortcut")
        fav_act=menu.addAction("Add to Favorites")
        cpath_act=menu.addAction("Copy Path")
        cname_act=menu.addAction("Copy Full Name")
        openloc_act=menu.addAction("Open File Location")
        share_act=menu.addAction("Share")
        openwith_act=menu.addAction("Open With")
        del_act=menu.addAction("Delete")
        prop_act=menu.addAction("Properties")

        chosen=menu.exec_(self.file_table.mapToGlobal(pos))
        if chosen is None:
            return

        if chosen==copy_act:
            self.copy_file(file_path)
        elif chosen==cut_act:
            self.cut_file(file_path)
        elif chosen==rename_act:
            self.rename_file(file_path)
        elif chosen==short_act:
            self.create_shortcut(file_path)
        elif chosen==fav_act:
            self.add_to_favorites(file_path)
        elif chosen==cpath_act:
            self.copy_path_to_clipboard(file_path)
        elif chosen==cname_act:
            self.copy_full_name(file_path)
        elif chosen==openloc_act:
            self.open_file_location(file_path)
        elif chosen==share_act:
            self.share_file(file_path)
        elif chosen==openwith_act:
            self.open_with(file_path)
        elif chosen==del_act:
            self.delete_file(file_path)
        elif chosen==prop_act:
            self.show_properties(file_path)

    ############################################
    # 12 Actions (same logic)
    ############################################
    def copy_file(self,path):
        if not os.path.exists(path):
            QMessageBox.warning(self,"File Not Found",f"Cannot copy: {path}")
            return
        target=QFileDialog.getExistingDirectory(self,"Select Destination")
        if target:
            try:
                shutil.copy2(path,target)
                self.status_label.setText(f"Copied '{os.path.basename(path)}' to '{target}'")
            except Exception as e:
                QMessageBox.warning(self,"Copy Error",str(e))

    def cut_file(self,path):
        if not os.path.exists(path):
            QMessageBox.warning(self,"File Not Found",f"Cannot move: {path}")
            return
        target=QFileDialog.getExistingDirectory(self,"Select Destination")
        if target:
            try:
                shutil.move(path,target)
                self.status_label.setText(f"Moved '{os.path.basename(path)}' to '{target}'")
            except Exception as e:
                QMessageBox.warning(self,"Move Error",str(e))

    def rename_file(self,path):
        if not os.path.exists(path):
            QMessageBox.warning(self,"File Not Found",f"Cannot rename: {path}")
            return
        new_name,ok=QInputDialog.getText(self,"Rename File","New name:", text=os.path.basename(path))
        if ok and new_name.strip():
            newp=os.path.join(os.path.dirname(path), new_name.strip())
            if os.path.exists(newp):
                QMessageBox.warning(self,"Rename Error","A file with that name already exists.")
                return
            try:
                os.rename(path,newp)
                self.status_label.setText(f"Renamed to: {os.path.basename(newp)}")
            except Exception as e:
                QMessageBox.warning(self,"Rename Error",str(e))

    def create_shortcut(self,path):
        if not os.path.exists(path):
            QMessageBox.warning(self,"File Not Found",f"Cannot create shortcut: {path}")
            return
        if not HAS_PYWIN32:
            QMessageBox.information(self,"Shortcut Error","PyWin32 not installed.")
            return
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            desk=os.path.join(os.environ["USERPROFILE"],"Desktop")
            name,_=os.path.splitext(os.path.basename(path))
            lnk_path=os.path.join(desk,f"{name}.lnk")
            shell=win32com.client.Dispatch("WScript.Shell")
            scut=shell.CreateShortCut(lnk_path)
            scut.Targetpath=path
            scut.WorkingDirectory=os.path.dirname(path)
            scut.IconLocation=path
            scut.save()
            self.status_label.setText(f"Shortcut created: {lnk_path}")
        except Exception as e:
            QMessageBox.warning(self,"Shortcut Error",str(e))

    def add_to_favorites(self,path):
        # stub or your own logic
        pass

    def copy_path_to_clipboard(self,path):
        QApplication.clipboard().setText(path)
        self.status_label.setText(f"Copied path: {path}")

    def copy_full_name(self,path):
        base=os.path.basename(path)
        QApplication.clipboard().setText(base)
        self.status_label.setText(f"Copied name: {base}")

    def open_file_location(self,path):
        if not os.path.isfile(path):
            QMessageBox.warning(self,"Not Found",f"Cannot locate: {path}")
            return
        subprocess.run(["explorer","/select,", path])

    def share_file(self,path):
        if not HAS_PYWIN32:
            QMessageBox.information(self,"Share Error","PyWin32 not installed.")
            return
        if not os.path.exists(path):
            QMessageBox.warning(self,"File Not Found",f"Cannot share: {path}")
            return
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            shell=win32com.client.Dispatch("Shell.Application")
            folder=shell.Namespace(os.path.dirname(path))
            item=folder.ParseName(os.path.basename(path))
            item.InvokeVerb("share")
            self.status_label.setText(f"Shared: {os.path.basename(path)}")
        except Exception as e:
            QMessageBox.warning(self,"Share Error",str(e))

    def open_with(self,path):
        if not HAS_PYWIN32:
            QMessageBox.information(self,"Open With Error","PyWin32 not installed.")
            return
        if not os.path.exists(path):
            QMessageBox.warning(self,"File Not Found",f"Cannot open with: {path}")
            return
        try:
            subprocess.run(["rundll32.exe","shell32.dll,OpenAs_RunDLL",path])
        except Exception as e:
            QMessageBox.warning(self,"Open With Error",str(e))

    def delete_file(self,path):
        if not os.path.exists(path):
            QMessageBox.warning(self,"File Not Found",f"Cannot delete: {path}")
            return
        confirm=QMessageBox.question(self,"Delete File",
                                     f"Are you sure you want to delete:\n{path}?",
                                     QMessageBox.Yes|QMessageBox.No)
        if confirm==QMessageBox.Yes:
            try:
                if HAS_SEND2TRASH:
                    send2trash.send2trash(path)
                    self.status_label.setText(f"Deleted (Recycle Bin): {path}")
                else:
                    os.remove(path)
                    self.status_label.setText(f"Deleted permanently: {path}")
            except Exception as e:
                QMessageBox.warning(self,"Delete Error",str(e))

    def show_properties(self,path):
        if not HAS_PYWIN32:
            QMessageBox.information(self,"Properties Error","PyWin32 not installed.")
            return
        if not os.path.exists(path):
            QMessageBox.warning(self,"File Not Found",f"Cannot show properties: {path}")
            return
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            shell=win32com.client.Dispatch("Shell.Application")
            folder=shell.Namespace(os.path.dirname(path))
            item=folder.ParseName(os.path.basename(path))
            item.InvokeVerb("properties")
            self.status_label.setText(f"Properties for: {os.path.basename(path)}")
        except Exception as e:
            QMessageBox.warning(self,"Properties Error",str(e))

def main():
    app=QApplication(sys.argv)
    w=ModernDriveSearch()
    w.show()
    sys.exit(app.exec_())

if __name__=="__main__":
    main()
