import ctypes
ctypes.windll.ole32.CoInitializeEx(0, 2)  # COM-safe initialization

import sys
import math
import numpy as np
import soundcard as sc
import warnings
import colorsys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QSystemTrayIcon, QMenu, QAction
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QColor, QPainter, QFont, QIcon

warnings.filterwarnings("ignore", category=sc.SoundcardRuntimeWarning)

def hsv_to_rgb(h, s, v):
    r, g, b = colorsys.hsv_to_rgb(h, s, v)
    return int(r * 255), int(g * 255), int(b * 255)

class VisualizerWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Visualizer")
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(300, 120)

        screen_geometry = QApplication.primaryScreen().geometry()
        x = max(0, screen_geometry.width() - self.width() - 10)
        y = max(0, screen_geometry.height() - self.height() - 50)
        self.move(x, y)

        self.num_bars = 30
        self.bar_values = [0] * self.num_bars
        self.hue = 0.0
        self.glow_phase = 0.0
        self.glow_alpha = 80

        self.timer = QTimer()
        self.timer.timeout.connect(self.update)
        self.timer.start(50)

        self.samplerate = 44100
        self.blocksize = 1024
        output_device = sc.default_speaker().name
        self.loopback_mic = sc.get_microphone(output_device, include_loopback=True)
        self.recorder_context = self.loopback_mic.recorder(samplerate=self.samplerate)
        self.recorder_context.__enter__()

        self.audio_timer = QTimer()
        self.audio_timer.timeout.connect(self.poll_audio)
        self.audio_timer.start(100)

        # Close button
        self.close_btn = QPushButton("×", self)
        self.close_btn.setGeometry(self.width() - 25, 5, 20, 20)
        self.close_btn.setStyleSheet("color: white; background: transparent; border: none; font-size: 16px;")
        self.close_btn.clicked.connect(self.hide)

        # Tray icon
        self.tray_icon = QSystemTrayIcon(QIcon("icon.png"), self)
        tray_menu = QMenu()
        restore_action = QAction("Show Visualizer", self)
        restore_action.triggered.connect(self.show)
        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(QApplication.quit)
        tray_menu.addAction(restore_action)
        tray_menu.addAction(quit_action)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

    def poll_audio(self):
        try:
            data = self.recorder_context.record(numframes=self.blocksize)
            volume = np.mean(np.abs(data)) * 500
            self.hue = (self.hue + 0.01) % 1.0
            self.glow_phase = (self.glow_phase + 0.1) % (2 * math.pi)
            self.glow_alpha = int(60 + 40 * math.sin(self.glow_phase))
            for i in range(self.num_bars):
                randomness = np.random.uniform(0.5, 1.5)
                target = volume * randomness
                current = self.bar_values[i]
                if target > current:
                    self.bar_values[i] = 0.1 * current + 0.4 * target
                else:
                    self.bar_values[i] = 0.1 * current + 0.3 * target
        except Exception:
            pass

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        bar_spacing = self.width() / (self.num_bars + 1)
        center_y = self.height() // 2

        r, g, b = hsv_to_rgb(self.hue, 1, 1)
        color = QColor(r, g, b, 200)
        glow = QColor(r, g, b, self.glow_alpha)

        for i, val in enumerate(self.bar_values):
            bar_height = int(min(val, self.height() - 20))
            bar_width = 6
            x = int((i + 1) * bar_spacing - bar_width / 2)
            y = center_y - bar_height

            painter.setPen(Qt.NoPen)
            painter.setBrush(glow)
            painter.drawRoundedRect(x, y, bar_width, bar_height, 3, 3)

            painter.setBrush(color)
            painter.drawRoundedRect(x, y, bar_width, bar_height, 3, 3)

    
        painter.setPen(QColor(255, 255, 255, 180))
        painter.setFont(QFont("Segoe UI", 10))

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            self.move(event.globalPos() - self.drag_position)
            event.accept()

    def closeEvent(self, event):
        event.ignore()
        self.hide()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VisualizerWindow()
    window.show()
    sys.exit(app.exec_())