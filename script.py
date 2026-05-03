import sys, os, time, requests, hashlib
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *
from win32com.shell import shell, shellcon
from win32com.client import Dispatch

GITHUB_JSON_URL = "https://raw.githubusercontent.com/MIchalAlince/pythonprogramstoredat/json/apps.json"
INSTALLED_DIR = os.path.join(os.getcwd(), "installed_apps")
os.makedirs(INSTALLED_DIR, exist_ok=True)

window = None

# ---------------- INTERNET ----------------
def has_internet():
    try:
        requests.get("https://www.google.com", timeout=3)
        return True
    except:
        return False

# ---------------- SHA256 ----------------
def check_sha256(file_path, expected):
    if not expected:
        return True
    if expected.startswith("sha256:"):
        expected = expected.split("sha256:")[1]

    h = hashlib.sha256()
    with open(file_path, "rb") as f:
        while chunk := f.read(8192):
            h.update(chunk)

    return h.hexdigest().lower() == expected.lower()

# ---------------- LOAD ----------------
class LoadThread(QThread):
    progress = Signal(int)
    finished = Signal(list)
    offline = Signal(bool)

    def run(self):
        try:
            for i in range(0, 40, 5):
                self.progress.emit(i)
                self.msleep(40)

            while not has_internet():
                self.offline.emit(True)
                self.msleep(1000)

            self.offline.emit(False)

            r = requests.get(GITHUB_JSON_URL, timeout=10)
            data = r.json()

            if isinstance(data, dict):
                data = data.get("apps", [])

            for i in range(40, 100, 5):
                self.progress.emit(i)
                self.msleep(30)

            self.finished.emit(data)

        except:
            self.finished.emit([])

# ---------------- DOWNLOAD ----------------
class DownloadThread(QThread):
    progress = Signal(int, float)
    done = Signal(str)
    fail = Signal(str)

    def __init__(self, url, target):
        super().__init__()
        self.url = url
        self.target = target
        self.running = True

    def run(self):
        try:
            r = requests.get(self.url, stream=True, timeout=15)

            total = int(r.headers.get("content-length", 1))
            downloaded = 0
            start = time.time()

            with open(self.target, "wb") as f:
                for chunk in r.iter_content(8192):
                    if not self.running:
                        return

                    if not chunk:
                        continue

                    f.write(chunk)
                    downloaded += len(chunk)

                    percent = int(downloaded * 100 / total)
                    speed = downloaded / 1024 / 1024 / (time.time() - start + 0.01)

                    self.progress.emit(percent, speed)

            if self.running:
                self.done.emit(self.target)

        except Exception as e:
            self.fail.emit(str(e))

    def stop(self):
        self.running = False

# ---------------- INSTALL DIALOG ----------------
class InstallDialog(QDialog):
    def __init__(self, app, mode="install"):
        super().__init__()
        self.app = app
        self.mode = mode

        self.setWindowTitle(app["name"])
        self.setFixedSize(450, 200)

        self.thread = None
        self.target = None

        layout = QVBoxLayout(self)

        title = "Aktualizovat" if mode == "update" else "Instalovat"

        self.label = QLabel(f"{title} {app['name']}?")
        self.bar = QProgressBar()
        self.info = QLabel("")

        self.btn_start = QPushButton(title)
        self.btn_cancel = QPushButton("Zrušit")

        self.btn_start.clicked.connect(self.start)
        self.btn_cancel.clicked.connect(self.cancel)

        layout.addWidget(self.label)
        layout.addWidget(self.bar)
        layout.addWidget(self.info)

        btns = QHBoxLayout()
        btns.addWidget(self.btn_start)
        btns.addWidget(self.btn_cancel)
        layout.addLayout(btns)

    # START
    def start(self):
        url = self.app.get("files", {}).get("Windows")

        if not url:
            QMessageBox.critical(self, "Chyba", "Chybí URL")
            self.reject()
            return

        self.target = os.path.join(INSTALLED_DIR, self.app["name"] + ".exe")

        self.thread = DownloadThread(url, self.target)
        self.thread.progress.connect(self.update)
        self.thread.done.connect(self.finish)
        self.thread.fail.connect(self.error)
        self.thread.start()

        self.btn_start.setEnabled(False)

    # UPDATE UI
    def update(self, p, s):
        self.bar.setValue(p)
        self.info.setText(f"{p}% | {s:.2f} MB/s")

    # CANCEL
    def cancel(self):
        if self.thread:
            self.thread.stop()
            self.thread.wait()

        if self.target and os.path.exists(self.target):
            os.remove(self.target)

        self.reject()

    # FINISH
    def finish(self, path):
        sha = self.app.get("sha256", {}).get("Windows")

        if sha and not check_sha256(path, sha):
            QMessageBox.critical(self, "Chyba", "SHA256 nesouhlasí")
            os.remove(path)
            self.reject()
            return

        # uložit verzi
        with open(path + ".ver", "w") as f:
            f.write(self.app["version"])

        try:
            desktop = shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0)
            link = os.path.join(desktop, self.app["name"] + ".lnk")

            sh = Dispatch("WScript.Shell").CreateShortcut(link)
            sh.TargetPath = path
            sh.save()
        except:
            pass

        self.accept()

    def error(self, msg):
        QMessageBox.critical(self, "Chyba", msg)
        self.reject()

# ---------------- STORE ----------------
class Store(QMainWindow):
    def __init__(self, apps):
        super().__init__()
        self.apps = apps
        self.selected = None
        self.dark = False

        self.setWindowTitle("Python Program Store")
        self.resize(900, 600)

        self.ui()

    def ui(self):
        w = QWidget()
        self.setCentralWidget(w)
        l = QVBoxLayout(w)

        top = QHBoxLayout()

        self.install_btn = QPushButton("Instalovat")
        self.update_btn = QPushButton("Aktualizovat")
        self.uninstall_btn = QPushButton("Odinstalovat")
        self.theme = QPushButton("🌙")

        self.install_btn.clicked.connect(self.install)
        self.update_btn.clicked.connect(self.update)
        self.uninstall_btn.clicked.connect(self.uninstall)
        self.theme.clicked.connect(self.toggle)

        top.addWidget(self.install_btn)
        top.addWidget(self.update_btn)
        top.addWidget(self.uninstall_btn)
        top.addWidget(self.theme)
        top.addStretch()

        l.addLayout(top)

        self.grid = QListWidget()
        self.grid.setViewMode(QListWidget.IconMode)
        self.grid.setIconSize(QSize(96, 96))
        self.grid.setSpacing(10)
        self.grid.itemClicked.connect(self.select)

        l.addWidget(self.grid)

        self.refresh()

    # ---------------- VERSION ----------------
    def local_version(self, app):
        f = os.path.join(INSTALLED_DIR, app["name"] + ".exe.ver")
        if not os.path.exists(f):
            return None
        return open(f).read().strip()

    def is_installed(self, app):
        return os.path.exists(os.path.join(INSTALLED_DIR, app["name"] + ".exe"))

    def has_update(self, app):
        lv = self.local_version(app)
        return lv is not None and lv != app["version"]

    # ---------------- BUTTON STATE ----------------
    def update_buttons(self):
        if not self.selected:
            self.install_btn.setEnabled(False)
            self.update_btn.setEnabled(False)
            self.uninstall_btn.setEnabled(False)
            return

        installed = self.is_installed(self.selected)
        update = self.has_update(self.selected)

        self.install_btn.setEnabled(not installed)
        self.update_btn.setEnabled(update)
        self.uninstall_btn.setEnabled(installed)

    # ---------------- SELECT ----------------
    def select(self, item):
        self.selected = item.data(Qt.UserRole)
        self.update_buttons()

    # ---------------- ACTIONS ----------------
    def install(self):
        if self.selected:
            InstallDialog(self.selected, "install").exec()
            self.refresh()

    def update(self):
        if self.selected:
            InstallDialog(self.selected, "update").exec()
            self.refresh()

    def uninstall(self):
        if not self.selected:
            return

        exe = os.path.join(INSTALLED_DIR, self.selected["name"] + ".exe")

        if os.path.exists(exe):
            os.remove(exe)

        ver = exe + ".ver"
        if os.path.exists(ver):
            os.remove(ver)

        self.refresh()

    # ---------------- UI ----------------
    def refresh(self):
        self.grid.clear()

        for a in self.apps:
            item = QListWidgetItem(a["name"])

            try:
                r = requests.get(a["icon_url"], timeout=5)
                pix = QPixmap()
                pix.loadFromData(r.content)
                item.setIcon(QIcon(pix))
            except:
                pass

            item.setData(Qt.UserRole, a)
            self.grid.addItem(item)

        self.update_buttons()

    def toggle(self):
        self.dark = not self.dark
        self.setStyleSheet("background:#1e1e1e; color:white;" if self.dark else "")

# ---------------- RUN ----------------
if __name__ == "__main__":
    app = QApplication(sys.argv)

    splash = QWidget()
    splash.setWindowTitle("Loading")
    splash.resize(300, 200)

    layout = QVBoxLayout(splash)
    bar = QProgressBar()
    label = QLabel("Načítám...")

    layout.addWidget(label)
    layout.addWidget(bar)

    splash.show()

    def offline(state):
        label.setText("čekám na internet..." if state else "Načítám...")

    def loaded(data):
        global window
        splash.close()
        window = Store(data)
        window.show()

    t = LoadThread()
    t.progress.connect(bar.setValue)
    t.offline.connect(offline)
    t.finished.connect(loaded)
    t.start()

    sys.exit(app.exec())
