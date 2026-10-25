import threading
import queue
import time
import os
import tempfile
import base64
import requests
import speech_recognition as sr
import pyttsx3
from gtts import gTTS
import pygame

# Ngưỡng kiểm tra chất lượng file âm thanh tối thiểu (1KB = 1024 bytes)
MIN_AUDIO_BYTES = 1024

# Cấu hình khoảng lặng mặc định (1.0s - 1.5s) để tự động cắt câu khi ngừng nói
DEFAULT_SILENCE_TIMEOUT_SEC = 1.2


def validate_wav_bytes(wav_bytes: bytes) -> dict:
    """
    Kiểm tra chất lượng dữ liệu âm thanh đầu vào:
    Nếu file âm thanh rỗng hoặc dung lượng nhỏ hơn 1KB (1024 bytes),
    chặn lại không gửi lên API để tránh quá tải hoặc lỗi Server 500.
    """
    if not wav_bytes or len(wav_bytes) < MIN_AUDIO_BYTES:
        size = len(wav_bytes) if wav_bytes else 0
        msg = f"File âm thanh quá nhỏ ({size} bytes < 1KB). Đã chặn gửi để tránh lỗi máy chủ."
        print(f"Error details: {msg}")
        return {"status": "error", "message": msg, "valid": False, "bytes": size}
    return {"status": "ok", "message": "Hợp lệ", "valid": True, "bytes": len(wav_bytes)}


def cleanup_temp_audio_file(temp_file_path: str = "temp_voice.wav"):
    """
    Xóa ngay file âm thanh tạm (ví dụ `temp_voice.wav`) trên máy tính bằng `os.remove()`
    ngay sau khi xử lý xong để tránh đọc trùng ở chu kỳ tiếp theo.
    """
    if temp_file_path and os.path.exists(temp_file_path):
        try:
            os.remove(temp_file_path)
        except OSError as e:
            print(f"Error details (Không thể xóa file tạm {temp_file_path}): {e}")


def post_voice_to_backend(
    wav_bytes: bytes,
    recognized_text: str = "",
    active_tab_idx: int = 0,
    endpoint_url: str = "http://localhost:3000/api/analyze-voice",
) -> dict:
    """
    Truyền dữ liệu âm thanh/văn bản qua giao thức HTTP POST tiêu chuẩn thay vì WebSocket.
    Bọc toàn bộ trong try ... except Exception as e để không bao giờ làm sập luồng thu âm.
    """
    try:
        check = validate_wav_bytes(wav_bytes)
        if not check["valid"]:
            return {"status": "error", "message": check["message"]}

        audio_b64 = base64.b64encode(wav_bytes).decode("utf-8")
        payload = {
            "audioBase64": audio_b64,
            "audioByteLength": len(wav_bytes),
            "mimeType": "audio/wav",
            "transcriptText": recognized_text,
            "activeTabIdx": active_tab_idx,
        }
        response = requests.post(endpoint_url, json=payload, timeout=5)
        return response.json()
    except Exception as e:
        print(f"Error details: {e}")
        return {"status": "error", "message": str(e)}


class TTSReader:
    """
    Class quản lý việc đọc văn bản thành giọng nói (Text-to-Speech) chuẩn Tiếng Việt (Chị Google)
    Chạy trên một luồng (Thread) riêng biệt để không làm đơ giao diện (GUI).
    """
    def __init__(self):
        self.speech_queue = queue.Queue()
        self.is_running = True

        try:
            pygame.mixer.init()
        except Exception as e:
            print(f"Error details (Lỗi khởi tạo Pygame Mixer): {e}")

        self.worker_thread = threading.Thread(target=self._process_queue, daemon=True)
        self.worker_thread.start()

    def _process_queue(self):
        fallback_engine = None
        try:
            fallback_engine = pyttsx3.init()
            fallback_engine.setProperty("rate", 170)
        except Exception as e:
            print(f"Error details (pyttsx3 init): {e}")

        while self.is_running:
            try:
                text = self.speech_queue.get(timeout=0.5)
                if text:
                    self._speak_vietnamese(text, fallback_engine)
                self.speech_queue.task_done()
            except queue.Empty:
                continue
            except Exception as e:
                print(f"Error details (Lỗi TTS): {e}")

    def _speak_vietnamese(self, text: str, fallback_engine):
        temp_path = None
        try:
            tts = gTTS(text=text, lang="vi", slow=False)
            fd, temp_path = tempfile.mkstemp(suffix=".mp3")
            os.close(fd)
            tts.save(temp_path)

            pygame.mixer.music.load(temp_path)
            pygame.mixer.music.play()

            while pygame.mixer.music.get_busy() and self.is_running:
                time.sleep(0.05)

            pygame.mixer.music.unload()
        except Exception as e:
            print(f"Error details (Không thể dùng giọng Google TTS, chuyển sang giọng máy tính): {e}")
            if fallback_engine:
                try:
                    fallback_engine.say(text)
                    fallback_engine.runAndWait()
                except Exception as ex:
                    print(f"Error details (Fallback TTS): {ex}")
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass

    def speak(self, text: str):
        while not self.speech_queue.empty():
            try:
                self.speech_queue.get_nowait()
                self.speech_queue.task_done()
            except queue.Empty:
                break
        self.speech_queue.put(text)

    def stop(self):
        self.is_running = False
        try:
            pygame.mixer.quit()
        except Exception as e:
            print(f"Error details (pygame quit): {e}")


class AudioListener:
    """
    Class quản lý việc lắng nghe từ Microphone LIÊN TỤC & CHỐNG NHẬP TRÙNG DỮ LIỆU 100%:
    1. Silence Timeout (1.0s - 1.5s): Tự động cắt câu ngay khi người dùng ngừng nói 1.2 giây.
    2. Processing Flag (`self.is_processed`):
       - Mặc định `self.is_processed = False`.
       - Khi có file âm thanh mới gửi lên, kiểm tra `if not self.is_processed:` mới tiến hành phân tích
         và chuyển ngay `self.is_processed = True`.
       - Chỉ khi Micro tự động bật lại để chờ câu nói mới ở chu kỳ tiếp theo mới đặt `self.is_processed = False`.
    3. Reset Audio State (`self.audio_bytes = None` & `os.remove("temp_voice.wav")`):
       - Ngay sau khi `/api/analyze-voice` xử lý xong và chèn dòng vào bảng, xóa sạch bộ nhớ đệm
         `self.audio_bytes = None` và xóa file tạm `temp_voice.wav` bằng `os.remove()`.
    """
    def __init__(
        self,
        on_speech_recognized_callback,
        on_status_change_callback=None,
        silence_timeout: float = DEFAULT_SILENCE_TIMEOUT_SEC,
        endpoint_url: str = "http://localhost:3000/api/analyze-voice",
    ):
        self.recognizer = sr.Recognizer()
        self.is_listening = False
        self.callback = on_speech_recognized_callback
        self.status_callback = on_status_change_callback
        self.endpoint_url = endpoint_url

        # Biến trạng thái chống nhập trùng dữ liệu (Processing Flag & Audio Buffer)
        self.is_processed = False
        self.audio_bytes = None
        self.temp_voice_path = "temp_voice.wav"
        self.auto_cycle_count = 0

        # Cấu hình Silence Timeout (1.0s - 1.5s) & Lọc ồn ao tôm
        self.recognizer.energy_threshold = 700
        self.recognizer.dynamic_energy_threshold = False
        self.set_silence_timeout(silence_timeout)

        self.listen_thread = None

    def set_silence_timeout(self, silence_sec: float):
        clamped = max(1.0, min(1.5, float(silence_sec)))
        self.recognizer.pause_threshold = clamped
        self.recognizer.non_speaking_duration = min(0.5, clamped * 0.4)
        self.recognizer.phrase_threshold = 0.2

    def set_energy_threshold(self, value: float):
        self.recognizer.energy_threshold = max(50, float(value))

    def adjust_ambient_noise(self, duration: float = 1.5) -> float:
        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=duration)
                self.recognizer.energy_threshold = max(350, self.recognizer.energy_threshold * 1.25)
                return self.recognizer.energy_threshold
        except Exception as e:
            print(f"Error details (Lỗi hiệu chuẩn tiếng ồn): {e}")
            return self.recognizer.energy_threshold

    def start_listening(self):
        if not self.is_listening:
            self.is_listening = True
            self.is_processed = False
            self.audio_bytes = None
            cleanup_temp_audio_file(self.temp_voice_path)
            self.listen_thread = threading.Thread(target=self._listen_loop, daemon=True)
            self.listen_thread.start()

    def stop_listening(self):
        self.is_listening = False
        self.audio_bytes = None
        cleanup_temp_audio_file(self.temp_voice_path)
        if self.status_callback:
            self.status_callback("Đã tắt Micro.")

    def _listen_loop(self):
        """
        Vòng lặp lắng nghe liên tục kèm khóa `is_processed` và xóa bộ nhớ đệm `audio_bytes = None`
        sau mỗi lần xử lý để đảm bảo đọc 1 lần chỉ nhảy đúng 1 dòng duy nhất.
        """
        while self.is_listening:
            try:
                with sr.Microphone() as source:
                    if self.status_callback:
                        self.status_callback(
                            f"🟢 Đang lắng nghe liên tục (Tự cắt sau {self.recognizer.pause_threshold:.1f}s lặng)..."
                        )

                    while self.is_listening:
                        try:
                            # Khi Micro tự động bật lại để chờ câu nói tiếp theo -> Đặt is_processed = False
                            self.is_processed = False
                            self.audio_bytes = None

                            audio = self.recognizer.listen(
                                source,
                                timeout=1.5,
                                phrase_time_limit=8.0,
                            )

                            if not self.is_listening:
                                break

                            # Nếu gói âm thanh của chu kỳ này chưa được xử lý (is_processed == False)
                            if not self.is_processed:
                                self.auto_cycle_count += 1
                                self._process_single_audio_once(audio, self.auto_cycle_count)

                            # Xóa sạch tham chiếu đối tượng audio trước khi bước sang chu kỳ kế tiếp
                            audio = None

                        except sr.WaitTimeoutError:
                            continue
                        except Exception as loop_err:
                            print(f"Error details (Chu kỳ thu âm tự động): {loop_err}")
                            time.sleep(0.2)
                            continue

            except Exception as e:
                print(f"Error details (Lỗi kết nối Micro, đang tự động thử lại sau 1.5s): {e}")
                if self.status_callback and self.is_listening:
                    self.status_callback("Đang tự động kết nối lại Micro...")
                time.sleep(1.5)

    def _process_single_audio_once(self, audio, cycle_id: int):
        """
        Xử lý đúng 1 lần duy nhất cho mỗi file âm thanh:
        - Kiểm tra `if self.is_processed: return`
        - Đánh dấu `self.is_processed = True` ngay lập tức
        - Xóa `self.audio_bytes = None` và `os.remove(self.temp_voice_path)` trong khối `finally`
        """
        if self.is_processed:
            return

        # Khóa cờ trạng thái ngay lập tức để chặn mọi lệnh gọi lặp lần 2
        self.is_processed = True

        try:
            self.audio_bytes = audio.get_wav_data()
            quality_check = validate_wav_bytes(self.audio_bytes)
            if not quality_check["valid"]:
                return

            # Nếu có ghi ra file tạm temp_voice.wav để xử lý
            with open(self.temp_voice_path, "wb") as f:
                f.write(self.audio_bytes)

            if self.status_callback:
                self.status_callback(f"⏳ Đang phân tích câu #{cycle_id} qua HTTP POST...")

            # Nhận diện bằng nguồn Google Speech Recognition gốc (vi-VN)
            text = self.recognizer.recognize_google(audio, language="vi-VN")

            if text and text.strip():
                text_lower = text.strip().lower()
                print(f"[Chu kỳ #{cycle_id} | Đọc 1 lần duy nhất]: \"{text_lower}\"")
                # Chèn dòng số liệu mới vào bảng đúng 1 lần
                self.callback(text_lower)

        except sr.UnknownValueError:
            pass
        except sr.RequestError as e:
            print(f"Error details (RequestError): {e}")
        except Exception as e:
            print(f"Error details (Xử lý âm thanh): {e}")
        finally:
            # 1. XÓA BỘ NHỚ ĐỆM ÂM THANH NGAY SAU KHI SỬ DỤNG (Reset Audio State)
            self.audio_bytes = None
            # 2. XÓA FILE TẠM TRÊN MÁY TÍNH NGAY LẬP TỨC TRƯỚC KHI KÍCH HOẠT LẠI MICRO
            cleanup_temp_audio_file(self.temp_voice_path)

            if self.status_callback and self.is_listening:
                self.status_callback(
                    f"🟢 Đã xóa bộ nhớ đệm & mở lại Micro (Chu kỳ #{cycle_id + 1}) - Hãy đọc tiếp..."
                )
