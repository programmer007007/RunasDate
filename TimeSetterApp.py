import sys
import os
import configparser
import time
import ntplib
import subprocess
import ctypes
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox
from PyQt5.QtCore import QDateTime, Qt
import win32com.client 


class SYSTEMTIME(ctypes.Structure):
    _fields_ = [
        ("wYear", ctypes.c_ushort),
        ("wMonth", ctypes.c_ushort),
        ("wDayOfWeek", ctypes.c_ushort),
        ("wDay", ctypes.c_ushort),
        ("wHour", ctypes.c_ushort),
        ("wMinute", ctypes.c_ushort),
        ("wSecond", ctypes.c_ushort),
        ("wMilliseconds", ctypes.c_ushort)
    ]

class TimeSetterApp(QWidget):
    def __init__(self):
        super().__init__()       
        self.initUI()
        self.config = configparser.ConfigParser()
        self.ini_path = "runasdate.ini"
        self.load_config()
        self.kernel32 = ctypes.windll.kernel32
        self.SetLocalTime = self.kernel32.SetLocalTime
        self.SetLocalTime.argtypes = [ctypes.POINTER(SYSTEMTIME)]
        self.SetLocalTime.restype = ctypes.c_bool

    def initUI(self):
        self.setWindowTitle('Time Setter App')
        layout = QVBoxLayout()

        # Date and time input
        date_time_layout = QHBoxLayout()
        self.date_time_edit = QLineEdit(self)
        self.date_time_edit.setPlaceholderText("YYYY-MM-DD HH:MM:SS")
        date_time_layout.addWidget(QLabel("Set Date and Time:"))
        date_time_layout.addWidget(self.date_time_edit)
        layout.addLayout(date_time_layout)

        # File chooser
        file_layout = QHBoxLayout()
        self.file_path_edit = QLineEdit(self)
        self.file_path_edit.setReadOnly(True)
        file_choose_btn = QPushButton("Choose File", self)
        file_choose_btn.clicked.connect(self.choose_file)
        file_layout.addWidget(QLabel("Executable:"))
        file_layout.addWidget(self.file_path_edit)
        file_layout.addWidget(file_choose_btn)
        layout.addLayout(file_layout)

        # Buttons
        button_layout = QHBoxLayout()
        save_btn = QPushButton("Save Settings", self)
        save_btn.clicked.connect(self.save_settings)
        sync_time_btn = QPushButton("Sync Time from Internet", self)
        sync_time_btn.clicked.connect(self.sync_time)
        run_btn = QPushButton("Run Program", self)
        run_btn.clicked.connect(self.run_program)
        button_layout.addWidget(save_btn)
        button_layout.addWidget(sync_time_btn)
        button_layout.addWidget(run_btn)
        layout.addLayout(button_layout)


        # Add new button for creating shortcut
        create_shortcut_btn = QPushButton("Create Shortcut", self)
        create_shortcut_btn.clicked.connect(self.create_shortcut)
        button_layout.addWidget(create_shortcut_btn)

        
        self.setLayout(layout)

    def load_config(self):
        self.config.read(self.ini_path)
        if 'main' in self.config:
            date_time_str = self.config['main'].get('date_time', '')
            self.date_time_edit.setText(date_time_str)
            self.file_path_edit.setText(self.config['main'].get('executable', ''))

    def choose_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Executable")
        if file_name:
            self.file_path_edit.setText(file_name)

    def save_settings(self):
        try:
            date_time = QDateTime.fromString(self.date_time_edit.text(), "yyyy-MM-dd HH:mm:ss")
            if not date_time.isValid():
                raise ValueError("Invalid date and time format")

            if not self.config.has_section('main'):
                self.config.add_section('main')

            self.config['main']['date_time'] = self.date_time_edit.text()
            self.config['main']['executable'] = self.file_path_edit.text()

            with open(self.ini_path, 'w') as configfile:
                self.config.write(configfile)

            QMessageBox.information(self, "Success", "Settings saved successfully!")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to save settings: {str(e)}")

    def sync_time(self):
        try:
            client = ntplib.NTPClient()
            response = client.request('pool.ntp.org')
            date_time = datetime.fromtimestamp(response.tx_time)
            self.date_time_edit.setText(date_time.strftime("%Y-%m-%d %H:%M:%S"))
            QMessageBox.information(self, "Success", "Time synced from internet!")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to sync time: {str(e)}")

    def run_program(self):
        executable = self.file_path_edit.text()
        if not os.path.isfile(executable):
            QMessageBox.warning(self, "Error", "Invalid executable path")
            return

        try:
            date_time = QDateTime.fromString(self.date_time_edit.text(), "yyyy-MM-dd HH:mm:ss")
            if not date_time.isValid():
                raise ValueError("Invalid date and time format")

            # Set system time
            new_time = SYSTEMTIME(
                wYear=date_time.date().year(),
                wMonth=date_time.date().month(),
                wDay=date_time.date().day(),
                wHour=date_time.time().hour(),
                wMinute=date_time.time().minute(),
                wSecond=date_time.time().second(),
                wMilliseconds=0
            )
            if not self.SetLocalTime(ctypes.byref(new_time)):
                raise OSError("Failed to set system time")

            # Run the program
            subprocess.Popen(executable)
            time.sleep(15)
            #QMessageBox.information(self, "Success", f"Launched {executable}")

            # Restore original time
            self.restore_time()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to run program: {str(e)}")

    def restore_time(self):
        try:
            client = ntplib.NTPClient()
            response = client.request('pool.ntp.org')
            date_time = datetime.fromtimestamp(response.tx_time)

            current_time = SYSTEMTIME(
                wYear=date_time.year,
                wMonth=date_time.month,
                wDay=date_time.day,
                wHour=date_time.hour,
                wMinute=date_time.minute,
                wSecond=date_time.second,
                wMilliseconds=date_time.microsecond // 1000
            )

            if not self.SetLocalTime(ctypes.byref(current_time)):
                raise OSError("Failed to restore system time")

            #QMessageBox.information(self, "Success", "Time restored successfully!")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to restore time: {str(e)}")

    def create_shortcut(self):
        try:
            # Determine if we're running as a script or executable
            if getattr(sys, 'frozen', False):
                # Running as executable
                executable = sys.executable
            else:
                # Running as script
                executable = sys.executable
                script_path = os.path.abspath(__file__)
                executable = f'"{executable}" "{script_path}"'

            date_time = self.date_time_edit.text()
            target_executable = self.file_path_edit.text()

            # Create shortcut name based on the chosen file
            shortcut_name = os.path.splitext(os.path.basename(target_executable))[0] + "Date.lnk"
            
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            shortcut_path = os.path.join(desktop, shortcut_name)

            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            
            if getattr(sys, 'frozen', False):
                shortcut.Targetpath = executable
                shortcut.Arguments = f'--run-program "{date_time}" "{target_executable}"'
            else:
                shortcut.Targetpath = sys.executable
                shortcut.Arguments = f'"{script_path}" --run-program "{date_time}" "{target_executable}"'

            shortcut.WorkingDirectory = os.path.dirname(executable)
            shortcut.save()

            QMessageBox.information(self, "Success", f"Shortcut created on desktop: {shortcut_path}")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to create shortcut: {str(e)}")
class SilentTimeSetterApp(TimeSetterApp):
    def __init__(self):
        super().__init__()

    def run_program(self):
        executable = self.file_path_edit.text()
        if not os.path.isfile(executable):
            print("Error: Invalid executable path")
            return

        try:
            date_time = QDateTime.fromString(self.date_time_edit.text(), "yyyy-MM-dd HH:mm:ss")
            if not date_time.isValid():
                raise ValueError("Invalid date and time format")

            # Set system time
            new_time = SYSTEMTIME(
                wYear=date_time.date().year(),
                wMonth=date_time.date().month(),
                wDay=date_time.date().day(),
                wHour=date_time.time().hour(),
                wMinute=date_time.time().minute(),
                wSecond=date_time.time().second(),
                wMilliseconds=0
            )
            if not self.SetLocalTime(ctypes.byref(new_time)):
                raise OSError("Failed to set system time")

            # Run the program
            subprocess.Popen(executable)

            # Wait for 15 seconds
            time.sleep(15)
            #print(f"Launched {executable}")

            # Restore original time
            self.restore_time_silent()
        except Exception as e:
            print(f"Error: Failed to run program: {str(e)}")

    def restore_time_silent(self):
        try:
            client = ntplib.NTPClient()
            response = client.request('pool.ntp.org')
            date_time = datetime.fromtimestamp(response.tx_time)

            current_time = SYSTEMTIME(
                wYear=date_time.year,
                wMonth=date_time.month,
                wDay=date_time.day,
                wHour=date_time.hour,
                wMinute=date_time.minute,
                wSecond=date_time.second,
                wMilliseconds=date_time.microsecond // 1000
            )

            if not self.SetLocalTime(ctypes.byref(current_time)):
                raise OSError("Failed to restore system time")

            print("Time restored successfully")
        except Exception as e:
            print(f"Error: Failed to restore time: {str(e)}")

def run_program_silently(date_time_str, executable):
    app = QApplication(sys.argv)
    setter = SilentTimeSetterApp()
    setter.date_time_edit.setText(date_time_str)
    setter.file_path_edit.setText(executable)
    setter.run_program()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    if len(sys.argv) > 1 and sys.argv[1] == "--run-program":
        if len(sys.argv) == 4:
            run_program_silently(sys.argv[2], sys.argv[3])
        else:
            print("Invalid arguments for silent run")
    else:
        ex = TimeSetterApp()
        ex.show()
        sys.exit(app.exec_())
