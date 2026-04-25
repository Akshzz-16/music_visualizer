import ctypes
ctypes.windll.ole32.CoInitializeEx(0, 2)  # COM-safe initialization

import asyncio
import sys
import math
import numpy as np
import soundcard as sc
import warnings
import colorsys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QSystemTrayIcon, QMenu, QAction
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QColor, QPainter, QFont, QIcon, QFontMetrics

try:
    from winsdk.windows.media.control import GlobalSystemMediaTransportControlsSessionManager
except ImportError:
    GlobalSystemMediaTransportControlsSessionManager = None

warnings.filterwarnings("ignore", category=sc.SoundcardRuntimeWarning)

def hsv_to_rgb(h, s, v):
    r, g, b = colorsys.hsv_to_rgb(h, s, v)
    return int(r * 255), int(g * 255), int(b * 255)

class VisualizerWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Visualizer")
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(300, 120)
        self.drag_position = None
        self.drag_active = False
        self.hovered = False
        self.setCursor(Qt.PointingHandCursor)
        self.setMouseTracking(True)

        screen_geometry = QApplication.primaryScreen().geometry()
        x = max(0, screen_geometry.width() - self.width() - 10)
        y = max(0, screen_geometry.height() - self.height() - 5)
        self.move(x, y)

        self.num_bars = 30
        self.bar_values = [0] * self.num_bars
        self.hue = 0.0
        self.glow_phase = 0.0
        self.glow_alpha = 80
        self.idle_phase = 0.0
        self.waveform = []
        self.current_track = ""
        self.media_available = GlobalSystemMediaTransportControlsSessionManager is not None

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

        self.media_timer = QTimer()
        self.media_timer.timeout.connect(self.refresh_media_info)
        self.media_timer.start(1500)

        # Close button
        self.close_btn = QPushButton("×", self)
        self.close_btn.setGeometry(self.width() - 42, 8, 34, 34)
        self.close_btn.setCursor(Qt.PointingHandCursor)
        self.close_btn.setStyleSheet(
            "QPushButton { color: white; background: rgba(255, 255, 255, 18); border: none; "
            "border-radius: 17px; font-size: 20px; font-weight: bold; }"
            "QPushButton:hover { background: rgba(255, 255, 255, 36); }"
        )
        self.close_btn.clicked.connect(self.hide)

        # Tray icon
        self.tray_icon = QSystemTrayIcon(QIcon("icon.ico"), self)
        self.tray_icon.setToolTip("Music Visualizer")

        menu = QMenu()
        restore_action = QAction("Show", self)
        restore_action.triggered.connect(self.show)
        menu.addAction(restore_action)

        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(QApplication.quit)
        menu.addAction(quit_action)

        self.tray_icon.setContextMenu(menu)
        self.tray_icon.show()

        # Start in bottom-origin bar mode
        self.visual_mode = 2
        self.show_bars = True
        self.show_waveform = False
        self.bar_origin = "bottom"
        self.mode_label = ""
        self.label_opacity = 0.0
        self.label_timer = QTimer(self)
        self.label_timer.timeout.connect(self.fade_label)

    def refresh_media_info(self):
        if not self.media_available:
            self.current_track = ""
            return

        try:
            self.current_track = asyncio.run(self.fetch_media_info())
        except Exception:
            self.current_track = ""

    async def fetch_media_info(self):
        session_manager = await GlobalSystemMediaTransportControlsSessionManager.request_async()
        current_session = session_manager.get_current_session()
        if current_session is None:
            return ""

        media_properties = await current_session.try_get_media_properties_async()
        title = (media_properties.title or "").strip()
        artist = (media_properties.artist or "").strip()

        if title and artist:
            return f"{title} - {artist}"
        return title or artist

    def poll_audio(self):
        try:
            data = self.recorder_context.record(numframes=self.blocksize)
            if data is None or len(data) == 0:
                raise ValueError("Empty audio frame")

            samples = data[:, 0]
            self.waveform = samples.copy()

            volume = np.mean(np.abs(samples)) * 500
            self.hue = (self.hue + 0.01) % 1.0
            self.glow_phase = (self.glow_phase + 0.1) % (2 * math.pi)
            self.glow_alpha = int(60 + 40 * math.sin(self.glow_phase))

            if volume < 0.01:
                self.recorder_context.__exit__(None, None, None)
                self.recorder_context = self.loopback_mic.recorder(samplerate=self.samplerate)
                self.recorder_context.__enter__()

                self.idle_phase = (self.idle_phase + 0.1) % (2 * math.pi)
                idle_value = 10 + 5 * math.sin(self.idle_phase)
                self.bar_values = [idle_value for _ in range(self.num_bars)]
                return

            for i in range(self.num_bars):
                randomness = np.random.uniform(0.5, 1.5)
                target = volume * randomness
                current = self.bar_values[i]
                if target > current:
                    self.bar_values[i] = 0.1 * current + 0.2 * target
                else:
                    self.bar_values[i] = 0.1 * current + 0.2 * target

        except Exception:
            try:
                self.recorder_context.__exit__(None, None, None)
                self.recorder_context = self.loopback_mic.recorder(samplerate=self.samplerate)
                self.recorder_context.__enter__()
            except:
                pass

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        center_y = self.height() // 2
        top_padding = 38 if self.hovered and self.current_track else 14
        bottom_padding = 16
        max_bar_height = max(24, self.height() - top_padding - bottom_padding)
        min_bar_height = 4
        bar_spacing = self.width() / (self.num_bars + 1)

        r, g, b = hsv_to_rgb(self.hue, 1, 1)
        color = QColor(r, g, b, 200)
        glow = QColor(r, g, b, self.glow_alpha)

        if self.hovered and self.current_track:
            font = QFont("Segoe UI", 10)
            painter.setFont(font)
            metrics = QFontMetrics(font)
            track_text = metrics.elidedText(self.current_track, Qt.ElideRight, self.width() - 72)
            text_rect = metrics.boundingRect(track_text)
            pill_width = text_rect.width() + 22
            pill_height = 24
            pill_x = max(12, (self.width() - pill_width) // 2)
            pill_y = 8

            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(0, 0, 0, 120))
            painter.drawRoundedRect(pill_x, pill_y, pill_width, pill_height, 13, 13)

            painter.setPen(QColor(255, 255, 255, 230))
            painter.drawText(pill_x + 11, pill_y + 16, track_text)

        # Waveform trail
        if self.show_waveform and len(self.waveform) > 0:
            painter.setPen(QColor(255, 255, 255, 60))
            scale = max_bar_height // 2
            step = max(1, len(self.waveform) // self.width())

            points = []
            for i in range(0, len(self.waveform), step):
                x = int(i / len(self.waveform) * self.width())
                y = int(center_y - self.waveform[i] * scale)
                points.append((x, y))

            for i in range(1, len(points)):
                painter.drawLine(points[i - 1][0], points[i - 1][1], points[i][0], points[i][1])

        # Bars
        if self.show_bars:
            for i, val in enumerate(self.bar_values):
                val = 0 if math.isnan(val) or math.isinf(val) else val
                bar_height = max(min_bar_height, int(min(val, max_bar_height)))
                bar_width = 6
                x = int((i + 1) * bar_spacing - bar_width / 2)

                if self.bar_origin == "center":
                    y = center_y - bar_height // 2
                else:  # bottom
                    y = self.height() - bar_height - bottom_padding

                painter.setPen(Qt.NoPen)
                painter.setBrush(glow)
                painter.drawRoundedRect(x, y, bar_width, bar_height, 3, 3)

                painter.setBrush(color)
                painter.drawRoundedRect(x, y, bar_width, bar_height, 3, 3)
        

        if self.mode_label and self.label_opacity > 0:
            painter.setOpacity(self.label_opacity)
            painter.setPen(QColor(255, 255, 255, int(255 * self.label_opacity)))
            font = QFont("Arial", 14)
            painter.setFont(font)
            painter.drawText(60, 30, self.mode_label)
            painter.setOpacity(1.0)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            if self.childAt(event.pos()) == self.close_btn:
                return
            self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
            self.drag_active = True
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()

    def mouseMoveEvent(self, event):
        if self.drag_active and event.buttons() == Qt.LeftButton and self.drag_position is not None:
            self.move(event.globalPos() - self.drag_position)
            event.accept()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_active = False
            self.setCursor(Qt.PointingHandCursor)
            event.accept()

    def enterEvent(self, event):
        self.hovered = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.hovered = False
        self.drag_active = False
        self.setCursor(Qt.PointingHandCursor)
        self.update()
        super().leaveEvent(event)

    def closeEvent(self, event):
        event.ignore()
        self.hide()
        self.tray_icon.showMessage(
            "Music Visualizer",
            "App is still running in the tray.",
            QSystemTrayIcon.Information,
            2000
        )

    def fade_label(self):
        self.label_opacity -= 0.05
        if self.label_opacity <= 0:
            self.label_opacity = 0
            self.mode_label = ""
            self.label_timer.stop()
        self.update()   

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Tab:
            self.visual_mode = (self.visual_mode + 1) % 3
            if self.visual_mode == 0:
                self.show_bars = True
                self.show_waveform = False
                self.bar_origin = "center"
                self.mode_label = "Center Bars Mode"
            elif self.visual_mode == 1:
                self.show_bars = False
                self.show_waveform = True
                self.mode_label = "Waveform Mode"
            elif self.visual_mode == 2:
                self.show_bars = True
                self.show_waveform = False
                self.bar_origin = "bottom"
                self.mode_label = "Bottom Bars Mode"

            self.label_opacity = 1.0
            self.label_timer.start(50)
            self.update()

    def restore_window(self):
        self.showNormal()
        self.activateWindow()

    def on_tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.restore_window()
    

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VisualizerWindow()
    window.show()
    sys.exit(app.exec_())
