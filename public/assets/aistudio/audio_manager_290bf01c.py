# -*- coding: utf-8 -*-
"""
Module: audio_manager.py
Chức năng: Quản lý Micro, thu âm liên tục rảnh tay, lọc nhiễu nền ao tôm,
nhận diện giọng nói tiếng Việt và phát âm thanh phản hồi (Beep).
"""

import time
import queue
import threading
from typing import Callable, Optional
import speech_recognition as sr

def play_audio_feedback(sound_type: str = 'beep'):
    """
    Phát âm thanh phản hồi nhẹ nhàng bằng luồng riêng (không khóa giao diện):
    - 'beep': Âm thanh Tít nhẹ khi nhận diện thành công (1200 Hz, 140 ms)
    - 'undo': Âm thanh 2 nốt lùi khi hoàn tác / hủy (900 Hz -> 650 Hz)
    - 'warning': Âm cảnh báo khi lỗi hoặc không nhận diện được (500 Hz, 200 ms)
    - 'start': Âm bật micro (800 Hz -> 1200 Hz)
    - 'stop': Âm tắt micro (1200 Hz -> 800 Hz)
    """
    def _run():
        try:
            import winsound
            if sound_type == 'beep':
                winsound.Beep(1200, 140)
            elif sound_type == 'undo':
                winsound.Beep(900, 100)
                time.sleep(0.04)
                winsound.Beep(650, 160)
            elif sound_type == 'warning':
                winsound.Beep(500, 200)
            elif sound_type == 'start':
                winsound.Beep(850, 100)
                winsound.Beep(1250, 140)
            elif sound_type == 'stop':
                winsound.Beep(1200, 100)
                winsound.Beep(750, 150)
        except Exception:
            # Fallback nếu thiết bị không hỗ trợ winsound
            pass

    threading.Thread(target=_run, daemon=True).start()

class VoiceListener:
    def __init__(
        self,
        on_text_recognized: Callable[[str], None],
        on_status_changed: Callable[[str, str], None],
        on_noise_adjusted: Optional[Callable[[float], None]] = None
    ):
        """
        Khởi tạo bộ lắng nghe giọng nói rảnh tay:
        :param on_text_recognized: Hàm callback khi AI nhận diện xong văn bản (str)
        :param on_status_changed: Hàm callback thông báo trạng thái ('LISTENING', 'PROCESSING', 'STOPPED', 'ERROR', message)
        :param on_noise_adjusted: Hàm callback khi cập nhật ngưỡng năng lượng lọc ồn (float)
        """
        self.on_text_recognized = on_text_recognized
        self.on_status_changed = on_status_changed
        self.on_noise_adjusted = on_noise_adjusted

        self.recognizer = sr.Recognizer()
        
        # MỤC 5: TỐI ƯU MÔI TRƯỜNG THỰC TẾ (Lọc nhiễu sục khí oxy, gió, nước)
        # Bật tự động điều chỉnh ngưỡng năng lượng động
        self.recognizer.dynamic_energy_threshold = True
        self.recognizer.dynamic_energy_ratio = 1.6
        self.recognizer.dynamic_energy_adjustment_damping = 0.15

        # Tối ưu nhận diện khoảng lặng để cắt câu nhanh (0.6s)
        self.recognizer.pause_threshold = 0.65
        self.recognizer.phrase_threshold = 0.2
        self.recognizer.non_speaking_duration = 0.4
        self.recognizer.energy_threshold = 800  # Khởi đầu phù hợp môi trường thực tế

        self.is_listening = False
        self._thread: Optional[threading.Thread] = None
        self.device_index: Optional[int] = None

    def set_device_index(self, index: Optional[int]):
        """Chọn micro cụ thể theo device index."""
        self.device_index = index

    def set_energy_threshold(self, threshold: float):
        """Người dùng điều chỉnh thủ công độ nhạy/ngưỡng lọc ồn."""
        self.recognizer.energy_threshold = max(50.0, float(threshold))

    def calibrate_ambient_noise(self, duration: float = 1.5):
        """
        Chạy hàm adjust_for_ambient_noise để đo tiếng ồn nền thực tế tại ao
        (tiếng sục khí oxy, máy quạt nước, tiếng gió).
        """
        def _calibrate():
            try:
                self.on_status_changed("CALIBRATING", "Đang đo mức ồn môi trường ao nuôi (Vui lòng giữ im lặng)...")
                with sr.Microphone(device_index=self.device_index) as source:
                    # Lọc nhiễu nền
                    self.recognizer.adjust_for_ambient_noise(source, duration=duration)
                
                curr_threshold = self.recognizer.energy_threshold
                msg = f"Đã hiệu chuẩn lọc ồn ao nuôi! Ngưỡng năng lượng: {curr_threshold:.0f}"
                self.on_status_changed("READY", msg)
                if self.on_noise_adjusted:
                    self.on_noise_adjusted(curr_threshold)
                play_audio_feedback('beep')
            except Exception as e:
                self.on_status_changed("ERROR", f"Lỗi hiệu chuẩn lọc ồn: {str(e)}")

        threading.Thread(target=_calibrate, daemon=True).start()

    def start_listening(self):
        """Bật chế độ rảnh tay - lắng nghe liên tục."""
        if self.is_listening:
            return

        self.is_listening = True
        play_audio_feedback('start')
        self._thread = threading.Thread(target=self._listen_loop, daemon=True)
        self._thread.start()

    def stop_listening(self):
        """Tắt micro."""
        if not self.is_listening:
            return

        self.is_listening = False
        play_audio_feedback('stop')
        self.on_status_changed("STOPPED", "Đã tắt Micro.")

    def _listen_loop(self):
        """Vòng lặp lắng nghe liên tục trong background thread."""
        try:
            with sr.Microphone(device_index=self.device_index) as source:
                # Tự động cân chỉnh nhanh 0.6s khi bắt đầu bật mic nếu chưa được đo
                self.on_status_changed("CALIBRATING", "Đang thích ứng lọc âm thanh nền ao...")
                self.recognizer.adjust_for_ambient_noise(source, duration=0.6)
                if self.on_noise_adjusted:
                    self.on_noise_adjusted(self.recognizer.energy_threshold)

                self.on_status_changed("LISTENING", "Micro ĐANG BẬT - Hãy đọc số...")

                while self.is_listening:
                    try:
                        # Lắng nghe âm thanh với timeout 1 giây để kiểm tra cờ is_listening
                        audio = self.recognizer.listen(source, timeout=1.2, phrase_time_limit=5.0)
                    except sr.WaitTimeoutError:
                        continue
                    except Exception as e:
                        if self.is_listening:
                            self.on_status_changed("LISTENING", f"Đang chờ tiếng: {str(e)[:30]}")
                        continue

                    if not self.is_listening:
                        break

                    # Khi nhận được âm thanh, chuyển sang luồng phụ xử lý AI để không làm gián đoạn lắng nghe câu tiếp theo
                    threading.Thread(target=self._recognize_audio, args=(audio,), daemon=True).start()

        except Exception as e:
            self.is_listening = False
            self.on_status_changed("ERROR", f"Lỗi khởi động Micro: {str(e)}")

    def _recognize_audio(self, audio: sr.AudioData):
        """Xử lý nhận diện âm thanh tiếng Việt bằng Google Speech Recognition."""
        try:
            self.on_status_changed("PROCESSING", "Đang nhận diện giọng nói...")
            text = self.recognizer.recognize_google(audio, language="vi-VN")
            if text and text.strip():
                # Gọi callback trả về chuỗi văn bản đã nhận diện
                self.on_text_recognized(text.strip())
                self.on_status_changed("LISTENING", f"Đã nghe: \"{text.strip()}\" - Sẵn sàng câu tiếp...")
            else:
                self.on_status_changed("LISTENING", "Micro đang bật - Sẵn sàng...")
        except sr.UnknownValueError:
            # Âm thanh không rõ hoặc là tiếng ồn máy bơm/sục khí
            self.on_status_changed("LISTENING", "Âm thanh không rõ hoặc tiếng sục khí - Hãy đọc lại...")
        except sr.RequestError as e:
            play_audio_feedback('warning')
            self.on_status_changed("ERROR", f"Lỗi kết nối máy chủ nhận diện (Cần Internet): {str(e)[:40]}")
        except Exception as e:
            self.on_status_changed("ERROR", f"Lỗi nhận diện: {str(e)[:40]}")

    @staticmethod
    def get_microphone_list():
        """Lấy danh sách các micro có trong máy tính."""
        try:
            return sr.Microphone.list_microphone_names()
        except Exception:
            return []
