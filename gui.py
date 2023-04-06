import sys

if sys.platform != "win32":
    print("This program only works on Windows OS.")
    sys.exit(1)

import psutil

from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QProgressBar,
    QPushButton,
    QButtonGroup,
)
from PyQt6.uic.load_ui import loadUi
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtGui import QIcon, QPixmap


from win32con import MB_ICONWARNING, MB_OKCANCEL
from win32api import MessageBoxEx
from win32api import GetSystemMetrics


import time
from pathlib import Path

from subprocess import Popen, PIPE
from os import system

class Worker(QThread):
    updateRamUsageSignal = pyqtSignal(float)
    updateRamTextSignal = pyqtSignal(str)

    updateCpuUsageSignal = pyqtSignal(float)

    def __init__(self, parent=None):
        super().__init__(parent)

    def run(self):
        while True:
            memory = psutil.virtual_memory()
            usedGb, totalGb = memory.used / 1024**3, memory.total / 1024**3
            percent = memory.percent

            self.updateRamUsageSignal.emit(percent)
            self.updateRamTextSignal.emit(
                f"{percent}% | {round(usedGb, 2)} GB / {round(totalGb)} GB"
            )

            cpu = psutil.cpu_percent()
            self.updateCpuUsageSignal.emit(cpu)
            self.sleep(1)


class Cleaner(QThread):
    finished = pyqtSignal()
    statusbar = pyqtSignal(str)
    elapsed = pyqtSignal(str)

    def __init__(self, parent=None, clearModes: dict = None):
        super().__init__(parent)
        self.clearModes = clearModes
        self.rammap = str(Path("./rammap.exe").absolute())

    def run(self):
        kw = {
            "shell": True,
            "stdout": PIPE,
            "stderr": PIPE,
        }
        exc, err = Popen(
            'reg.exe query "HKCU\Software\Sysinternals\RamMap" /v EulaAccepted /t REG_DWORD',
            **kw,
        ).communicate()
        if err.decode():
            _ = Popen(
                'reg.exe ADD "HKCU\Software\Sysinternals\RamMap" /v EulaAccepted /t REG_DWORD /d 1 /f',
                **kw,
            ).wait()

        def calculate_runtime(commandstr: str):
            start = time.perf_counter()
            system(commandstr)
            end = time.perf_counter()
            time.sleep(2)
            return round(end - start, 2)

        fs = time.perf_counter()
        if self.clearModes["optionWorkingSets"]:
            self.statusbar.emit("Cleaning Working Sets...")
            elapsed = calculate_runtime(f"{self.rammap} -Ew")
            self.statusbar.emit(f"Cleaning Working Sets took {elapsed} seconds")

        if self.clearModes["optionMPages"]:
            self.statusbar.emit("Cleaning Modified Pages...")
            elapsed = calculate_runtime(f"{self.rammap} -Em")
            self.statusbar.emit(f"Cleaning Modified Pages took {elapsed} seconds")

        if self.clearModes["optionStandby"]:
            self.statusbar.emit("Cleaning Standby List...")
            elapsed = calculate_runtime(f"{self.rammap} -Es")
            self.statusbar.emit(f"Cleaning Standby List took {elapsed} seconds")

        if self.clearModes["optionPriority0"]:
            self.statusbar.emit("Cleaning Priority 0...")
            elapsed = calculate_runtime(f"{self.rammap} -E0")
            self.statusbar.emit(f"Cleaning Priority 0 took {elapsed} seconds")

        if self.clearModes["optionSysMPages"]:
            self.statusbar.emit("Cleaning System Modified Pages...")
            elapsed = calculate_runtime(f"{self.rammap} -Es")
            self.statusbar.emit(
                f"Cleaning System Modified Pages took {elapsed} seconds"
            )

        efs = time.perf_counter() - fs
        self.finished.emit()
        self.elapsed.emit(f"Cleaning took {round(efs, 2)} seconds.")


class Widget(QMainWindow):
    def __init__(self):
        super().__init__()
        loadUi("widget.ui", self)
        self.loadUi()
        self.show()

    def loadUi(self):
        self.ramUsage: QProgressBar
        self.cpuUsage: QProgressBar

        self.optionButtons: QButtonGroup

        screensize = GetSystemMetrics(0), GetSystemMetrics(1)
        self.setFixedSize(310 * screensize[0] // 1920, 300 * screensize[1] // 1080)

        self.clearModes = {}
        for button in self.optionButtons.buttons():
            button.clicked.connect(self.clearMethods)
            self.clearModes[button.objectName()] = button.isChecked()

        self.warnForSysMPages = True
        
        self.setWindowIcon(QIcon(str(Path("./icon.ico").absolute())))

        self.progressBarUpdaterTask = Worker()
        self.progressBarUpdaterTask.updateRamUsageSignal.connect(self.updateRamUsage)
        self.progressBarUpdaterTask.updateRamTextSignal.connect(self.ramUsage.setFormat)
        self.progressBarUpdaterTask.updateCpuUsageSignal.connect(self.updateCpuUsage)
        self.progressBarUpdaterTask.start()

        self.cleanerThread = Cleaner()
        self.cleanerThread.statusbar.connect(self.statusbar.showMessage)
        self.cleanerThread.finished.connect(self.finishedCleaning)
        self.cleanerThread.elapsed.connect(self.statusbar.showMessage)

        self.cleanRamButton: QPushButton
        self.cleanRamButton.clicked.connect(self.runCleaner)

        self.statusBar().showMessage("Ready!")

    def updateRamUsage(self, percent):
        self.ramUsage.setValue(int(percent))

    def updateCpuUsage(self, percent):
        self.cpuUsage.setValue(int(percent))

    def clearMethods(self):
        sender: QPushButton = self.sender()
        if sender.objectName() == "optionSysMPages" and self.warnForSysMPages:
            msgbox = MessageBoxEx(
                0,
                "Uyarı! Bu opsiyon sistem kararsızlığına sebep olabilir.",
                "RAM Cleaner",
                MB_OKCANCEL | MB_ICONWARNING,
            )
            if msgbox == 2:
                self.optionSysMPages.setChecked(False)
                return
            self.warnForSysMPages = False

        self.clearModes[sender.objectName()] = sender.isChecked()

        if not any(self.clearModes.values()):
            self.statusBar().showMessage("You must select at least one option!")
            self.cleanRamButton.setEnabled(False)
        else:
            self.statusBar().showMessage("Ready!")
            self.cleanRamButton.setEnabled(True)

    def runCleaner(self):
        self.cleanerThread.clearModes = self.clearModes
        self.cleanRamButton.setEnabled(False)
        self.cleanerThread.start()

    def finishedCleaning(self):
        self.cleanerThread.quit()
        self.cleanRamButton.setEnabled(True)


def main():
    app = QApplication(sys.argv)
    
    widget = Widget()
    widget.setContentsMargins(5, 5, 5, 5)
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
