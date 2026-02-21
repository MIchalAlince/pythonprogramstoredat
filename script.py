import sys, os, platform, shutil, requests
from packaging.version import Version
from PySide6.QtWidgets import (
    QApplication, QWidget, QListWidget, QListWidgetItem,
    QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QMessageBox, QFileDialog, QComboBox, QLineEdit
)
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtCore import Qt, QPropertyAnimation, QSize

SYSTEM = platform.system()
INSTALL_ROOT = os.path.join(os.getenv("LOCALAPPDATA"), "PythonProgramStore")
os.makedirs(INSTALL_ROOT, exist_ok=True)
APPS_URL = "https://raw.githubusercontent.com/MIchalAlince/pythonprogramstoredat/json/apps.json"

# ================= Shortcut =================
def create_shortcut(target, shortcut_path, icon=None):
    import win32com.client
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    shortcut.TargetPath = target
    shortcut.WorkingDirectory = os.path.dirname(target)
    if icon:
        shortcut.IconLocation = icon
    shortcut.save()

def start_menu_path(name):
    return os.path.join(
        os.getenv("APPDATA"),
        r"Microsoft\Windows\Start Menu\Programs",
        name + ".lnk"
    )

# ================= STORE =================
class Store(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Python Program Store")
        self.resize(1000, 650)

        main = QVBoxLayout(self)

        # ===== TOP BAR =====
        top = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Vyhledat aplikaci...")
        self.search.textChanged.connect(self.filter_apps)

        self.sort_combo = QComboBox()
        self.sort_combo.addItems(["Název A–Z", "Název Z–A", "Nejnovější verze"])
        self.sort_combo.currentIndexChanged.connect(self.sort_apps)

        self.view_combo = QComboBox()
        self.view_combo.addItems(["Řádky", "Dlaždice"])
        self.view_combo.currentIndexChanged.connect(self.change_view)

        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Světlý režim", "Tmavý režim"])
        self.theme_combo.currentIndexChanged.connect(self.change_theme)

        top.addWidget(self.search)
        top.addWidget(self.sort_combo)
        top.addWidget(self.view_combo)
        top.addWidget(self.theme_combo)
        main.addLayout(top)

        # ===== CONTENT =====
        content = QHBoxLayout()
        self.list = QListWidget()
        self.list.itemClicked.connect(self.show_details)

        self.details_icon = QLabel()
        self.details_icon.setFixedSize(128, 128)

        self.details = QLabel("Vyber aplikaci")
        self.details.setWordWrap(True)

        right_panel = QVBoxLayout()
        right_panel.addWidget(self.details_icon)
        right_panel.addWidget(self.details)

        content.addWidget(self.list, 3)
        content.addLayout(right_panel, 2)
        main.addLayout(content)

        # ===== BUTTONS =====
        btns = QHBoxLayout()
        self.install_btn = QPushButton("Install")
        self.uninstall_btn = QPushButton("Uninstall")
        btns.addWidget(self.install_btn)
        btns.addWidget(self.uninstall_btn)
        main.addLayout(btns)

        self.install_btn.clicked.connect(self.install)
        self.uninstall_btn.clicked.connect(self.uninstall)

        self.set_light_theme()
        self.load_apps()

    # ================= THEME =================
    def set_light_theme(self):
        self.setStyleSheet("""
            QWidget { background-color: white; color: black; }
            QListWidget { background-color: #F2F2F2; }
        """)

    def set_dark_theme(self):
        self.setStyleSheet("""
            QWidget { background-color: #1E1E1E; color: white; }
            QListWidget { background-color: #252526; }
        """)

    def change_theme(self):
        if self.theme_combo.currentIndex() == 0:
            self.set_light_theme()
        else:
            self.set_dark_theme()

    # ================= LOAD =================
    def load_apps(self):
        self.list.clear()
        self.apps = requests.get(APPS_URL).json()
        self.displayed_apps = self.apps.copy()
        self.populate_list()

    def populate_list(self):
        self.list.clear()
        for app in self.displayed_apps:
            item = QListWidgetItem(app["name"])
            item.app = app
            if "icon_url" in app:
                try:
                    r = requests.get(app["icon_url"])
                    pix = QPixmap()
                    pix.loadFromData(r.content)
                    item.setIcon(QIcon(pix))
                except:
                    pass
            self.list.addItem(item)

    # ================= FILTER =================
    def filter_apps(self):
        text = self.search.text().lower()
        self.displayed_apps = [a for a in self.apps if text in a["name"].lower()]
        self.sort_apps()

    # ================= SORT =================
    def sort_apps(self):
        mode = self.sort_combo.currentIndex()
        if mode == 0:
            self.displayed_apps.sort(key=lambda x: x["name"])
        elif mode == 1:
            self.displayed_apps.sort(key=lambda x: x["name"], reverse=True)
        elif mode == 2:
            self.displayed_apps.sort(
                key=lambda x: Version(x["version"]),
                reverse=True
            )
        self.populate_list()

    # ================= VIEW =================
    def change_view(self):
        if self.view_combo.currentIndex() == 0:
            self.list.setViewMode(QListWidget.ListMode)
        else:
            self.list.setViewMode(QListWidget.IconMode)
            self.list.setIconSize(QSize(96,96))
        self.populate_list()

    # ================= DETAILS =================
    def show_details(self, item):
        app = item.app
        if "icon_url" in app:
            try:
                r = requests.get(app["icon_url"])
                pix = QPixmap()
                pix.loadFromData(r.content)
                self.details_icon.setPixmap(pix.scaled(128,128))
            except:
                self.details_icon.clear()

        sha_value = app["sha256"][SYSTEM]
        if sha_value.startswith("sha256:"):
            sha_value = sha_value.split(":",1)[1]

        self.details.setText(
            f"<h2>{app['name']}</h2>"
            f"<b>Verze:</b> {app['version']}<br>"
            f"<b>Datum vydání:</b> {app.get('release_date','?')}<br>"
            f"<b>SHA256:</b> {sha_value}<br>"
            f"{app['description']}"
        )

        anim = QPropertyAnimation(self.details, b"windowOpacity")
        anim.setDuration(200)
        anim.setStartValue(0)
        anim.setEndValue(1)
        anim.start()

    # ================= INSTALL =================
    def install(self):
        item = self.list.currentItem()
        if not item: return
        app = item.app

        folder = QFileDialog.getExistingDirectory(
            self, "Vyber složku",
            os.path.join(INSTALL_ROOT, app["name"])
        )
        if not folder: return

        os.makedirs(folder, exist_ok=True)
        exe_name = os.path.basename(app["files"][SYSTEM])
        exe_path = os.path.join(folder, exe_name)

        r = requests.get(app["files"][SYSTEM])
        with open(exe_path, "wb") as f:
            f.write(r.content)

        create_shortcut(
            exe_path,
            os.path.join(os.path.expanduser("~"), "Desktop", app["name"]+".lnk"),
            exe_path
        )

        create_shortcut(
            exe_path,
            start_menu_path(app["name"]),
            exe_path
        )

        QMessageBox.information(self,"Hotovo","Instalace dokončena.")

    # ================= UNINSTALL =================
    def uninstall(self):
        item = self.list.currentItem()
        if not item: return
        app = item.app
        folder = os.path.join(INSTALL_ROOT, app["name"])

        if os.path.exists(folder):
            shutil.rmtree(folder)

        for path in [
            os.path.join(os.path.expanduser("~"), "Desktop", app["name"]+".lnk"),
            start_menu_path(app["name"])
        ]:
            if os.path.exists(path):
                os.remove(path)

        QMessageBox.information(self,"Hotovo","Aplikace odstraněna.")


if __name__=="__main__":
    app = QApplication(sys.argv)
    s = Store()
    s.show()
    sys.exit(app.exec())
