# -*- coding: utf-8 -*-
"""
HỆ THỐNG QUẢN LÝ VÀ ĐẾM SỐ LƯỢNG TÔM THÔNG MINH BẰNG GIỌNG NÓI TIẾNG VIỆT
(GOOGLE SPEECH RECOGNITION - GOOGLE STT VI-VN & NÓI TIẾNG VIỆT TTS)
Các tính năng nâng cấp:
1. Tích hợp đọc giọng nói tiếng Việt (gTTS & pyttsx3/SAPI) tự động đọc lại đúng số liệu vừa nhập qua loa/tai nghe:
   - Chế độ 1: "Đã lưu 45 gam đực", "Đã lưu 35 gam cái"
   - Chế độ 2: "Đã thêm 5 con, tổng 55", "Đã hủy 5 con, tổng còn 50"
2. Lọc câu phiếm ngoài ao: Nếu câu nhận diện không chứa số hoặc không chứa từ khóa "đực/cái/xóa/hủy", tự động bỏ qua.
3. Thống kê tự động: Ô tính Size tôm trung bình (= Tổng số con / Tổng trọng lượng).
4. Bảng ánh xạ dịch từ chữ sang số tự động cho Chế độ 2 (Đếm cộng dồn).
5. Giữ nguyên toàn bộ giao diện 2 Tab, đếm mẫu, cộng dồn (Undo) và xuất báo cáo Excel 2 Sheet.
"""

import os
import sys
import re
import time
import queue
import hashlib
import datetime
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Đảm bảo in console tiếng Việt không lỗi trên Windows
if sys.stdout and sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import speech_recognition as sr
from excel_manager import save_shrimp_report, DEFAULT_EXCEL_FILENAME
from vietnamese_parser import parse_mode1_weighing_and_gender

# =============================================================================
# BẢNG ÁNH XẠ DỊCH TỪ CHỮ SANG SỐ TIẾNG VIỆT TỰ ĐỘNG CHO CHẾ ĐỘ 2
# =============================================================================
VIETNAMESE_WORD_TO_NUMBER = {
    # 0 - 9
    'không': 0, 'khong': 0, 'zero': 0,
    'một': 1, 'mot': 1, 'mốt': 1,
    'hai': 2,
    'ba': 3,
    'bốn': 4, 'bon': 4, 'tư': 4, 'tu': 4,
    'năm': 5, 'nam': 5, 'lăm': 5, 'lam': 5, 'nhăm': 5,
    'sáu': 6, 'sau': 6,
    'bảy': 7, 'bay': 7, 'bẩy': 7,
    'tám': 8, 'tam': 8,
    'chín': 9, 'chin': 9,

    # 10 - 19
    'mười': 10, 'muoi': 10, 'mươi': 10,
    'mười một': 11, 'mười hai': 12, 'mười ba': 13,
    'mười bốn': 14, 'mười tư': 14,
    'mười lăm': 15, 'mười năm': 15, 'mười sáu': 16,
    'mười bảy': 17, 'mười bẩy': 17, 'mười tám': 18, 'mười chín': 19,

    # 20 - 29
    'hai mươi': 20, 'hai chục': 20,
    'hai mốt': 21, 'hai một': 21, 'hai hai': 22, 'hai ba': 23,
    'hai tư': 24, 'hai bốn': 24,
    'hai lăm': 25, 'hai năm': 25, 'hai sáu': 26,
    'hai bảy': 27, 'hai bẩy': 27, 'hai tám': 28, 'hai chín': 29,

    # 30 - 39
    'ba mươi': 30, 'ba chục': 30,
    'ba mốt': 31, 'ba hai': 32, 'ba ba': 33,
    'ba tư': 34, 'ba bốn': 34,
    'ba lăm': 35, 'ba năm': 35, 'ba sáu': 36,
    'ba bảy': 37, 'ba bẩy': 37, 'ba tám': 38, 'ba chín': 39,

    # 40 - 49
    'bốn mươi': 40, 'bốn chục': 40,
    'bốn mốt': 41, 'bốn hai': 42, 'bốn ba': 43,
    'bốn tư': 44, 'bốn bốn': 44,
    'bốn lăm': 45, 'bốn năm': 45, 'bốn sáu': 46,
    'bốn bảy': 47, 'bốn bẩy': 47, 'bốn tám': 48, 'bốn chín': 49,

    # 50 - 59
    'năm mươi': 50, 'năm chục': 50,
    'năm mốt': 51, 'năm hai': 52, 'năm ba': 53,
    'năm tư': 54, 'năm bốn': 54,
    'năm lăm': 55, 'năm năm': 55, 'năm sáu': 56,
    'năm bảy': 57, 'năm bẩy': 57, 'năm tám': 58, 'năm chín': 59,

    # 60 - 100
    'sáu mươi': 60, 'sáu chục': 60, 'sáu lăm': 65,
    'bảy mươi': 70, 'bảy chục': 70, 'bảy lăm': 75,
    'tám mươi': 80, 'tám chục': 80, 'tám lăm': 85,
    'chín mươi': 90, 'chín chục': 90, 'chín lăm': 95,
    'một trăm': 100, 'trăm': 100
}

# =============================================================================
# CƠ CHẾ KHÓA AN TOÀN TOÀN CỤC (DEBOUNCE / COOLDOWN LOCK) CHỐNG DOUBLE TRIGGER
# =============================================================================
is_processing = False

# =============================================================================
# CẤU HÌNH TỐC ĐỘ ĐỌC GIỌNG NÓI TIẾNG VIỆT (TTS SPEECH RATE)
# =============================================================================
# Bạn có thể dễ dàng thay đổi tốc độ đọc mặc định tại đây (hoặc chọn trên thanh điều khiển GUI):
# - 1.0 : Tốc độ chuẩn gốc của Google TTS (~120 từ/phút)
# - 1.2 : Nhanh vừa phải (nhanh hơn 20%, nghe rõ ràng, tự nhiên)
# - 1.3 : Tối ưu nhất cho môi trường ao nuôi (khuyên dùng, nhanh hơn 30%, lanh lợi, dứt khoát)
# - 1.5 : Rất nhanh (nhanh hơn 50%, phù hợp kỹ thuật viên bốc tôm tốc độ cao)
# - 1.8 : Siêu tốc
DEFAULT_TTS_SPEED_RATE = 1.3

# =============================================================================
# BỘ LỌC TỪ KHÓA NHẬN DIỆN GIỚI TÍNH TÔM TIẾNG ANH & TIẾNG VIỆT (CHẾ ĐỘ 1)
# =============================================================================
def detect_shrimp_gender_mode1(text: str) -> str:
    """
    Nhận diện giới tính tôm tại Chế độ 1 (Cân mẫu):
    1. Nếu văn bản nhận diện được chứa từ "cái", "gái", hoặc "ái" -> Xác định là giống Cái -> Trả về "Female".
    2. Nếu văn bản nhận diện được chứa từ "trai", "chai", "đực", hoặc "tai" -> Xác định là giống Đực -> Trả về "Male".
    3. Hỗ trợ dự phòng các từ khóa thông dụng khác nếu có (female, male, meo, mê, mel, mail, phi meo, pheme, femail, phi, trống, mái).
    4. Nếu không nghe thấy từ khóa giới tính -> Giữ nguyên "Chưa rõ".
    """
    if not text:
        return "Chưa rõ"

    # Chèn khoảng trắng giữa chữ và số để tách từ chuẩn xác (VD: "35trai" -> "35 trai", "35chai" -> "35 chai", "45gái" -> "45 gái")
    t = re.sub(r'(\d+)([a-zA-Zà-ỹÀ-Ỹ]+)', r'\1 \2', text.lower().strip())
    t = re.sub(r'([a-zA-Zà-ỹÀ-Ỹ]+)(\d+)', r'\1 \2', t)

    # 1. KIỂM TRA GIỐNG CÁI (FEMALE): "cái", "gái", "ái" (kiểm tra trước)
    if re.search(r'\b(cái|cai|gái|gai|ái|female|phi\s*meo|phimeo|pheme|femail|phi|mái|mai)\b', t):
        return "Female"

    # 2. KIỂM TRA GIỐNG ĐỰC (MALE): "trai", "chai", "đực", "tai"
    if re.search(r'\b(trai|chai|đực|duc|tai|male|mail|meo|mê|mel|trống|trong)\b', t):
        return "Male"

    return "Chưa rõ"

def is_relevant_shrimp_command(text: str) -> bool:
    """
    MỤC 2: LOGIC LỌC CÂU NÓI CHUYỆN PHIẾM
    Nếu văn bản nhận diện được không chứa số hoặc không chứa từ khóa 'đực/cái/xóa/hủy/reset',
    hãy bỏ qua và không thực hiện hành động gì để tránh lỗi khi người dùng nói chuyện phiếm ngoài ao.
    """
    if not text:
        return False
    t = text.lower().strip()

    # Bắt ngay các từ khóa điều khiển: Xóa, Hủy, Reset, Xóa dòng, Xóa hết, Xóa tất cả...
    if any(k in t for k in ("xóa", "xoa", "hủy", "huy", "undo", "reset")):
        return True

    # 1. Kiểm tra có chứa chữ số (0 - 9)
    if re.search(r'\d', t):
        return True

    # 2. Kiểm tra từ khóa số tiếng Việt
    num_pattern = r'\b(không|khong|một|mot|mốt|hai|ba|bốn|bon|tư|tu|năm|nam|lăm|lam|nhăm|nham|sáu|sau|bảy|bay|bẩy|bey|tám|tam|chín|chin|mười|muoi|mươi|chục|trăm|tram)\b'
    if re.search(num_pattern, t):
        return True

    # 3. Kiểm tra từ khóa bắt buộc: 'trai', 'chai', 'đực', 'tai', 'cái', 'gái', 'ái', 'xóa', 'hủy', 'reset', ...
    kw_pattern = r'\b(trai|chai|đực|duc|tai|cái|cai|gái|gai|ái|trống|mái|xóa|xoa|hủy|huy|undo|reset|bỏ|bo|trừ|tru|male|female|mail|meo|mê|mel|pheme|femail|phi)\b'
    if re.search(kw_pattern, t) or "phi meo" in t:
        return True

    # Không chứa số và không chứa từ khóa -> Câu nói chuyện phiếm ngoài ao
    return False

def map_spoken_text_to_batch(speech_text: str):
    """
    Áp dụng bảng ánh xạ dịch từ chữ sang số tự động cho Chế độ 2:
    - Nếu là lệnh Undo ("xóa", "hủy", "undo", "bỏ", "trừ") -> ('UNDO', None)
    - Nếu là số ("6", "sáu", "10", "mười", "15", "mười lăm", "ba lăm") -> ('ADD', số)
    """
    if not speech_text:
        return None, None

    text = speech_text.lower().strip()
    text = re.sub(r'[,.](?!\d)', ' ', text)
    text = re.sub(r'[\-_/\\:;!?()\[\]{}"\'~`@#$%^&*+=|]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()

    # 1. Kiểm tra từ khóa Hoàn tác (Undo / Hủy)
    if re.search(r'\b(xóa|xoa|hủy|huy|undo|bỏ|bo|trừ|tru|quay lại|lùi|lui|nhầm|nham)\b', text):
        return 'UNDO', None

    # 2. Loại bỏ các từ đệm phổ biến khi bốc tôm
    cleaned = re.sub(r'\b(mẻ|me|thêm|them|cộng|cong|con|tôm|tom|bốc|boc|đợt|dot)\b', ' ', text)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()

    # 3. Tra bảng ánh xạ trực tiếp
    if cleaned in VIETNAMESE_WORD_TO_NUMBER:
        val = VIETNAMESE_WORD_TO_NUMBER[cleaned]
        if val > 0:
            return 'ADD', val

    # 4. Kiểm tra chuỗi số trực tiếp ("6", "10", "15", "20", "25", "50")
    num_match = re.search(r'\b(\d+)\b', cleaned)
    if num_match:
        try:
            val = int(num_match.group(1))
            if val > 0:
                return 'ADD', val
        except ValueError:
            pass

    # 5. Phân tích số ghép nâng cao
    words = cleaned.split()
    total = 0
    curr = 0
    for w in words:
        if w in VIETNAMESE_WORD_TO_NUMBER:
            n = VIETNAMESE_WORD_TO_NUMBER[w]
            if n == 100:
                total += (curr if curr > 0 else 1) * 100
                curr = 0
            elif n == 10 and curr > 0:
                curr = curr * 10
            elif curr > 0 and curr % 10 == 0:
                curr += n
            elif curr == 0:
                curr = n
            else:
                curr = curr * 10 + n
        elif w.isdigit():
            curr = int(w)
    total += curr

    if total > 0:
        return 'ADD', total

    return None, None

# =============================================================================
# MỤC 1: BỘ PHÁT GIỌNG NÓI TIẾNG VIỆT (gTTS & PYTTSX3 / SAPI FALLBACK - HỖ TRỢ TĂNG TỐC ĐỘ)
# =============================================================================
class VietnameseTTS:
    """
    Bộ đọc giọng nói tiếng Việt tự động đọc lại số liệu vừa nhập qua loa/tai nghe:
    - Hỗ trợ tùy chỉnh tốc độ đọc (speech rate: x1.0, x1.2, x1.3, x1.5...) linh hoạt.
    - Tích hợp PyAV (FFmpeg atempo filter) để tăng tốc độ phát âm giữ nguyên cao độ (pitch),
      giọng đọc nữ tiếng Việt tự nhiên, lanh lợi, dứt khoát, không bị méo tiếng.
    - Tự động lưu cache file âm thanh tăng tốc (.wav) để phát tức thì (< 1ms).
    - Dự phòng pyttsx3 / Windows SAPI khi mất kết nối mạng.
    """
    def __init__(self, cache_dir=None, speed_rate: float = DEFAULT_TTS_SPEED_RATE):
        self.cache_dir = cache_dir or os.path.join(os.path.dirname(__file__), "tts_cache")
        try:
            os.makedirs(self.cache_dir, exist_ok=True)
        except Exception:
            pass
        self._lock = threading.Lock()
        self.enabled = True
        self.speed_rate = max(0.5, min(3.0, float(speed_rate)))

    def set_speed_rate(self, speed_rate: float):
        """
        Thay đổi tham số tốc độ đọc (speech rate):
        - 1.0 : Tốc độ bình thường
        - 1.2 : Nhanh vừa phải (khuyên dùng nghe nhẹ nhàng)
        - 1.3 : Nhanh hơn 30% (mặc định tối ưu cho trại tôm)
        - 1.5 : Rất nhanh (phù hợp kỹ thuật viên thao tác nhanh)
        """
        try:
            val = float(speed_rate)
            self.speed_rate = max(0.5, min(3.0, val))
        except (ValueError, TypeError):
            self.speed_rate = 1.3

    def speak_async(self, text: str):
        """Phát âm thanh trong background thread không làm gián đoạn giao diện."""
        if not self.enabled or not text:
            return
        threading.Thread(target=self._speak_worker, args=(text,), daemon=True).start()

    def _speedup_audio(self, src_mp3: str, dst_wav: str, speed_val: float) -> bool:
        """
        Tăng tốc độ file âm thanh MP3 sang WAV với speech rate chỉ định (x1.2, x1.5...):
        - Sử dụng PyAV (FFmpeg libavfilter atempo) để time-stretch giữ nguyên cao độ giọng nói (pitch).
        - Xuất file WAV chuẩn PCM 16-bit để phát tức thì qua Windows winmm / winsound.
        """
        try:
            import av
            from av.filter import Graph
            import numpy as np
            import wave

            container = av.open(src_mp3)
            audio_stream = next(s for s in container.streams if s.type == 'audio')
            sample_rate = audio_stream.rate or 24000

            graph = Graph()
            src = graph.add_abuffer(template=audio_stream)
            atempo_filter = graph.add("atempo", str(speed_val))
            sink = graph.add("abuffersink")

            src.link_to(atempo_filter)
            atempo_filter.link_to(sink)
            graph.configure()

            out_frames = []
            for frame in container.decode(audio_stream):
                graph.push(frame)
                while True:
                    try:
                        out_f = graph.pull()
                        out_frames.append(out_f.to_ndarray())
                    except Exception:
                        break

            graph.push(None)  # Flush
            while True:
                try:
                    out_f = graph.pull()
                    out_frames.append(out_f.to_ndarray())
                except Exception:
                    break

            container.close()

            if not out_frames:
                return False

            data_tempo = np.concatenate(out_frames, axis=1)
            if data_tempo.dtype == np.int16:
                int16_data = data_tempo[0]
            else:
                int16_data = np.clip(data_tempo[0] * 32767, -32768, 32767).astype(np.int16)

            with wave.open(dst_wav, "wb") as wf:
                wf.setnchannels(1)
                wf.setsampwidth(2)
                wf.setframerate(sample_rate)
                wf.writeframes(int16_data.tobytes())

            return True
        except Exception as e:
            return False

    def _speak_worker(self, text: str):
        with self._lock:
            # 1. Thử dùng gTTS (Giọng đọc tiếng Việt tự nhiên) kết hợp tăng tốc độ
            try:
                txt_hash = hashlib.md5(text.encode('utf-8')).hexdigest()
                cached_orig = os.path.join(self.cache_dir, f"{txt_hash}.mp3")

                speed_val = round(self.speed_rate, 2)
                speed_tag = f"{speed_val:.2f}".rstrip('0').rstrip('.')
                cached_fast = os.path.join(self.cache_dir, f"{txt_hash}_x{speed_tag}.wav")

                # Trường hợp tốc độ chuẩn 1.0x
                if speed_val == 1.0:
                    if not os.path.exists(cached_orig):
                        from gtts import gTTS
                        tts = gTTS(text=text, lang='vi', slow=False)
                        tts.save(cached_orig)
                    self._play_sound(cached_orig)
                    return

                # Trường hợp tốc độ tăng tốc (!= 1.0x): Nếu đã có cache WAV -> Phát tức thì!
                if os.path.exists(cached_fast):
                    self._play_sound(cached_fast)
                    return

                # Nếu chưa có file gốc MP3 -> Tải qua gTTS
                if not os.path.exists(cached_orig):
                    from gtts import gTTS
                    tts = gTTS(text=text, lang='vi', slow=False)
                    tts.save(cached_orig)

                # Tăng tốc độ bằng PyAV atempo
                success = self._speedup_audio(cached_orig, cached_fast, speed_val)
                if success and os.path.exists(cached_fast):
                    self._play_sound(cached_fast)
                    return
                else:
                    # Dự phòng: Phát file gốc nếu không tăng tốc được
                    self._play_sound(cached_orig)
                    return
            except Exception:
                pass

            # 2. Dự phòng bằng pyttsx3 (Nếu máy có giọng đọc tiếng Việt SAPI)
            try:
                import pyttsx3
                engine = pyttsx3.init()
                default_rate = engine.getProperty('rate') or 200
                engine.setProperty('rate', int(default_rate * self.speed_rate))
                engine.say(text)
                engine.runAndWait()
                return
            except Exception:
                pass

            # 3. Dự phòng bằng Windows PowerShell SpeechSynthesizer
            try:
                clean_t = text.replace("'", " ").replace('"', " ")
                ps_rate = max(-10, min(10, int((self.speed_rate - 1.0) * 8)))
                cmd = f'powershell -Command "Add-Type -AssemblyName System.Speech; $s = New-Object System.Speech.Synthesis.SpeechSynthesizer; $s.Rate = {ps_rate}; $s.Speak(\'{clean_t}\')"'
                subprocess.run(cmd, shell=True, timeout=5)
            except Exception:
                pass

    def _play_sound(self, sound_path: str):
        """Phát âm thanh (.wav hoặc .mp3) qua thư viện chuẩn winmm của Windows."""
        try:
            import ctypes
            winmm = ctypes.windll.winmm
            alias = f"tts_{int(time.time() * 1000)}"
            ext = os.path.splitext(sound_path)[1].lower()
            device_type = "waveaudio" if ext == ".wav" else "mpegvideo"

            winmm.mciSendStringW(f'open "{sound_path}" type {device_type} alias {alias}', None, 0, 0)
            winmm.mciSendStringW(f'play {alias} wait', None, 0, 0)
            winmm.mciSendStringW(f'close {alias}', None, 0, 0)
        except Exception:
            try:
                if sound_path.endswith('.wav'):
                    import winsound
                    winsound.PlaySound(sound_path, winsound.SND_FILENAME)
            except Exception:
                pass

def play_audio_feedback(sound_type: str = 'beep'):
    """Phát âm thanh Tít nhẹ phản hồi phụ."""
    def _run():
        try:
            import winsound
            if sound_type == 'beep':
                winsound.Beep(1200, 130)
            elif sound_type == 'undo':
                winsound.Beep(900, 100)
                time.sleep(0.04)
                winsound.Beep(650, 160)
            elif sound_type == 'warning':
                winsound.Beep(500, 200)
            elif sound_type == 'start':
                winsound.Beep(850, 90)
                winsound.Beep(1250, 130)
            elif sound_type == 'stop':
                winsound.Beep(1200, 90)
                winsound.Beep(750, 140)
        except Exception:
            pass

    threading.Thread(target=_run, daemon=True).start()

# =============================================================================
# LỚP LẮNG NGHE GIỌNG NÓI GOOGLE SPEECH RECOGNITION (GOOGLE STT)
# =============================================================================
class GoogleSpeechVoiceListener:
    def __init__(
        self,
        on_text_recognized,
        on_status_changed,
        on_noise_adjusted=None
    ):
        self.on_text_recognized = on_text_recognized
        self.on_status_changed = on_status_changed
        self.on_noise_adjusted = on_noise_adjusted

        self.recognizer = sr.Recognizer()

        # Cấu hình tối ưu lọc ồn và cắt câu tức thì
        self.recognizer.dynamic_energy_threshold = True
        self.recognizer.dynamic_energy_ratio = 1.6
        self.recognizer.dynamic_energy_adjustment_damping = 0.15
        self.recognizer.pause_threshold = 0.65       # Ngắt câu sau 0.65s khoảng lặng
        self.recognizer.phrase_threshold = 0.2
        self.recognizer.non_speaking_duration = 0.4
        self.recognizer.energy_threshold = 700      # Ngưỡng năng lượng khởi điểm

        self.is_listening = False
        self.device_index = None
        self._thread = None

    def set_device_index(self, index):
        self.device_index = index

    def set_energy_threshold(self, threshold):
        self.recognizer.energy_threshold = max(50.0, float(threshold))

    def calibrate_ambient_noise(self, duration=1.5):
        """Đo mức ồn nền thực tế của quạt nước và sục khí tại ao nuôi."""
        def _calib():
            try:
                self.on_status_changed("CALIBRATING", "Đang đo mức ồn quạt nước và sục khí ao nuôi...")
                with sr.Microphone(device_index=self.device_index) as source:
                    self.recognizer.adjust_for_ambient_noise(source, duration=duration)
                curr = self.recognizer.energy_threshold
                msg = f"Đã lọc ồn ao nuôi! Ngưỡng năng lượng: {curr:.0f}"
                self.on_status_changed("READY", msg)
                if self.on_noise_adjusted:
                    self.on_noise_adjusted(curr)
                play_audio_feedback('beep')
            except Exception as e:
                self.on_status_changed("ERROR", f"Lỗi hiệu chuẩn lọc ồn: {str(e)[:40]}")

        threading.Thread(target=_calib, daemon=True).start()

    def start_listening(self):
        """Bật chế độ rảnh tay liên tục."""
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
        """Vòng lặp thu âm rảnh tay trên background thread."""
        try:
            with sr.Microphone(device_index=self.device_index) as source:
                self.on_status_changed("CALIBRATING", "Đang thích ứng với âm thanh ao nuôi...")
                self.recognizer.adjust_for_ambient_noise(source, duration=0.6)
                if self.on_noise_adjusted:
                    self.on_noise_adjusted(self.recognizer.energy_threshold)

                self.on_status_changed("LISTENING", "Micro ĐANG BẬT (Google STT) - Hãy đọc số...")

                while self.is_listening:
                    try:
                        audio = self.recognizer.listen(source, timeout=1.2, phrase_time_limit=5.0)
                    except sr.WaitTimeoutError:
                        continue
                    except Exception:
                        continue

                    if not self.is_listening:
                        break

                    threading.Thread(target=self._recognize_worker, args=(audio,), daemon=True).start()

        except Exception as e:
            self.is_listening = False
            self.on_status_changed("ERROR", f"Lỗi khởi động Micro: {str(e)[:40]}")

    def _recognize_worker(self, audio):
        """Gửi âm thanh lên Google Speech Recognition (vi-VN)."""
        global is_processing
        try:
            self.on_status_changed("PROCESSING", "Đang nhận diện giọng nói...")
            text = self.recognizer.recognize_google(audio, language="vi-VN")
            if text and text.strip():
                # LỖI 3: Khóa Debounce - Ngay khi nhận được kết quả từ Google,
                # nếu is_processing là True thì lập tức bỏ qua câu lệnh đó
                if is_processing:
                    print(f"[DEBOUNCE] Đang xử lý câu trước, bỏ qua double trigger từ Google STT: \"{text.strip()}\"")
                    return

                self.on_text_recognized(text.strip())
                self.on_status_changed("LISTENING", f"Đã nghe: \"{text.strip()}\" - Sẵn sàng...")
            else:
                self.on_status_changed("LISTENING", "Micro đang bật - Sẵn sàng...")
        except sr.UnknownValueError:
            self.on_status_changed("LISTENING", "Tiếng sục khí hoặc không rõ chữ - Hãy đọc lại...")
        except sr.RequestError as e:
            play_audio_feedback('warning')
            self.on_status_changed("ERROR", f"Lỗi kết nối Google STT (Cần Internet): {str(e)[:40]}")
        except Exception as e:
            self.on_status_changed("ERROR", f"Lỗi nhận diện: {str(e)[:40]}")

    @staticmethod
    def get_microphone_list():
        try:
            return sr.Microphone.list_microphone_names()
        except Exception:
            return ["Mặc định hệ thống"]

# =============================================================================
# GIAO DIỆN CHÍNH ỨNG DỤNG (TKINTER GUI)
# =============================================================================
class ShrimpSmartCounterApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("HỆ THỐNG QUẢN LÝ & ĐẾM TÔM THÔNG MINH BẰNG GIỌNG NÓI TIẾNG VIỆT - AI SHRIMP COUNTER")
        self.geometry("1160x820")
        self.minsize(1020, 720)

        # Định vị thư mục ứng dụng và file Excel (Tương thích 100% khi đóng gói PyInstaller .EXE)
        if getattr(sys, 'frozen', False):
            # Khi chạy từ file .exe độc lập, lưu dữ liệu tại thư mục chứa file .exe của người dùng
            self.app_dir = os.path.dirname(sys.executable)
            self.bundle_dir = getattr(sys, '_MEIPASS', self.app_dir)
        else:
            # Khi chạy từ file script .py thông thường
            self.app_dir = os.path.dirname(os.path.abspath(__file__))
            self.bundle_dir = self.app_dir

        self.excel_path = os.path.join(self.app_dir, DEFAULT_EXCEL_FILENAME)

        # Dữ liệu nội tại
        self.weight_records = []
        self.batch_records = []
        self.total_batch_count = 0
        self.batch_counter = 0

        # LỖI 3: Khóa an toàn Debounce / Cooldown Lock chống double trigger
        self.is_processing = False

        # Hàng đợi giao tiếp thread-safe
        self.msg_queue = queue.Queue()

        # Khởi tạo bộ đọc giọng nói tiếng Việt TTS (hỗ trợ tùy chỉnh tốc độ đọc)
        self.tts = VietnameseTTS(
            cache_dir=os.path.join(self.app_dir, "tts_cache"),
            speed_rate=DEFAULT_TTS_SPEED_RATE
        )

        # Khởi tạo bộ nhận diện Google Speech Recognition
        self.voice_listener = GoogleSpeechVoiceListener(
            on_text_recognized=self._on_speech_text_received,
            on_status_changed=self._on_audio_status_changed,
            on_noise_adjusted=self._on_noise_adjusted
        )

        # Thiết lập giao diện
        self._setup_styles()
        self._build_ui()

        # Bắt đầu vòng lặp đọc hàng đợi UI
        self.after(80, self._process_msg_queue)

        # Giao thức khi tắt ứng dụng
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _setup_styles(self):
        self.style = ttk.Style(self)
        try:
            self.style.theme_use('clam')
        except Exception:
            pass

        self.COLOR_BG = "#F4F7F9"
        self.COLOR_NAVY = "#0F2C59"
        self.COLOR_BLUE_HEADER = "#1E3E62"
        self.COLOR_GREEN = "#28A745"
        self.COLOR_GREEN_HOVER = "#218838"
        self.COLOR_RED = "#DC3545"
        self.COLOR_RED_HOVER = "#C82333"
        self.COLOR_BORDER = "#D6E0EA"

        self.configure(bg=self.COLOR_BG)

        self.style.configure(
            "TNotebook",
            background=self.COLOR_BG,
            borderwidth=0
        )
        self.style.configure(
            "TNotebook.Tab",
            font=("Segoe UI", 11, "bold"),
            padding=[20, 10],
            background="#DDE4ED",
            foreground="#333333"
        )
        self.style.map(
            "TNotebook.Tab",
            background=[("selected", "#0F2C59")],
            foreground=[("selected", "#FFFFFF")]
        )

        self.style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background="#1E3E62",
            foreground="#FFFFFF",
            relief="flat",
            padding=[5, 6]
        )
        self.style.map("Treeview.Heading", background=[("active", "#0F2C59")])
        self.style.configure(
            "Treeview",
            font=("Segoe UI", 10),
            rowheight=28,
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            borderwidth=1
        )
        self.style.map(
            "Treeview",
            background=[("selected", "#CDE5FF")],
            foreground=[("selected", "#000000")]
        )

    def _build_ui(self):
        # 1. HEADER (MỤC 1)
        self._build_header_panel()

        # 2. LỌC ỒN AO NUÔI & TÙY CHỌN TTS (MỤC 1 & 5)
        self._build_noise_control_panel()

        # 3. NỘI DUNG 2 TAB CHUYÊN BIỆT (MỤC 2 & MỤC 3)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=15, pady=(5, 10))

        # Tab 1: Cân trọng lượng & Phân loại giới tính (Kèm ô tính Size tôm)
        self.tab1_frame = tk.Frame(self.notebook, bg=self.COLOR_BG)
        self.notebook.add(self.tab1_frame, text="  ⚖️ CHẾ ĐỘ 1: CÂN TRỌNG LƯỢNG & GIỚI TÍNH  ")
        self._build_tab1_content()

        # Tab 2: Đếm số lượng cộng dồn tốc độ cao
        self.tab2_frame = tk.Frame(self.notebook, bg=self.COLOR_BG)
        self.notebook.add(self.tab2_frame, text="  ⚡ CHẾ ĐỘ 2: ĐẾM CỘNG DỒN TỐC ĐỘ CAO (UNDO)  ")
        self._build_tab2_content()

        # 4. THANH HÀNH ĐỘNG DƯỚI CÙNG & XUẤT EXCEL
        self._build_bottom_panel()

    # =========================================================================
    # MỤC 1: CẤU TRÚC GIAO DIỆN CHUNG
    # =========================================================================
    def _build_header_panel(self):
        header_card = tk.Frame(self, bg="#FFFFFF", highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        header_card.pack(fill="x", padx=15, pady=(10, 5))

        info_frame = tk.Frame(header_card, bg="#FFFFFF")
        info_frame.pack(side="left", padx=15, pady=8)

        lbl_pond = tk.Label(info_frame, text="TÊN AO NUÔI / MÃ LÔ:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", fg="#333333")
        lbl_pond.grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.entry_pond = tk.Entry(info_frame, font=("Segoe UI", 11, "bold"), width=20, fg="#0F2C59", relief="solid", bd=1)
        self.entry_pond.grid(row=0, column=1, sticky="w", padx=5, pady=2)
        self.entry_pond.insert(0, "Ao Số 1 - Khu A")

        lbl_staff = tk.Label(info_frame, text="TÊN NHÂN VIÊN:", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", fg="#333333")
        lbl_staff.grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.entry_staff = tk.Entry(info_frame, font=("Segoe UI", 11), width=20, fg="#0F2C59", relief="solid", bd=1)
        self.entry_staff.grid(row=1, column=1, sticky="w", padx=5, pady=2)
        self.entry_staff.insert(0, "Kỹ thuật viên 1")

        mid_frame = tk.Frame(header_card, bg="#FFFFFF")
        mid_frame.pack(side="left", expand=True, fill="both", padx=15, pady=8)

        self.lbl_mic_status_badge = tk.Label(
            mid_frame,
            text="🔴 MICRO ĐANG TẮT",
            font=("Segoe UI", 11, "bold"),
            bg="#F8D7DA",
            fg="#721C24",
            padx=12,
            pady=3,
            relief="groove"
        )
        self.lbl_mic_status_badge.pack(anchor="w")

        self.lbl_mic_status_detail = tk.Label(
            mid_frame,
            text="Bấm 'BẬT MICRO' để bắt đầu đếm rảnh tay (Tự động đọc lại số liệu qua tai nghe).",
            font=("Segoe UI", 9, "italic"),
            bg="#FFFFFF",
            fg="#666666"
        )
        self.lbl_mic_status_detail.pack(anchor="w", pady=(3, 0))

        btn_frame = tk.Frame(header_card, bg="#FFFFFF")
        btn_frame.pack(side="right", padx=15, pady=8)

        self.btn_toggle_mic = tk.Button(
            btn_frame,
            text="🎤 BẬT MICRO (RẢNH TAY)",
            font=("Segoe UI", 13, "bold"),
            bg=self.COLOR_GREEN,
            fg="#FFFFFF",
            activebackground=self.COLOR_GREEN_HOVER,
            activeforeground="#FFFFFF",
            padx=18,
            pady=10,
            cursor="hand2",
            relief="raised",
            bd=3,
            command=self._toggle_microphone
        )
        self.btn_toggle_mic.pack()

    # =========================================================================
    # LỌC ỒN AO NUÔI & TÙY CHỌN PHÁT GIỌNG NÓI TTS
    # =========================================================================
    def _build_noise_control_panel(self):
        panel = tk.Frame(self, bg="#E8EEF5", highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        panel.pack(fill="x", padx=15, pady=(0, 6))

        # Checkbox bật/tắt đọc giọng nói tiếng Việt (MỤC 1)
        self.var_tts_enabled = tk.BooleanVar(value=True)
        chk_tts = tk.Checkbutton(
            panel,
            text="🔊 Đọc Lại Số Liệu",
            variable=self.var_tts_enabled,
            font=("Segoe UI", 9, "bold"),
            bg="#E8EEF5",
            fg="#0F2C59",
            activebackground="#E8EEF5",
            command=self._on_tts_toggle
        )
        chk_tts.pack(side="left", padx=(10, 2), pady=5)

        # Tùy chọn tốc độ đọc giọng nói TTS (speech rate: x1.0, x1.2, x1.3, x1.5, x1.8)
        lbl_speed = tk.Label(panel, text="Tốc độ:", font=("Segoe UI", 9, "bold"), bg="#E8EEF5", fg="#0F2C59")
        lbl_speed.pack(side="left", padx=(4, 2), pady=5)

        self.combo_tts_speed = ttk.Combobox(
            panel,
            values=["x1.0 (Chuẩn)", "x1.2 (Nhanh vừa)", "x1.3 (Mặc định)", "x1.5 (Rất nhanh)", "x1.8 (Siêu tốc)"],
            state="readonly",
            width=14
        )
        speed_map_init = {
            1.0: "x1.0 (Chuẩn)",
            1.2: "x1.2 (Nhanh vừa)",
            1.3: "x1.3 (Mặc định)",
            1.5: "x1.5 (Rất nhanh)",
            1.8: "x1.8 (Siêu tốc)"
        }
        self.combo_tts_speed.set(speed_map_init.get(DEFAULT_TTS_SPEED_RATE, f"x{DEFAULT_TTS_SPEED_RATE}"))
        self.combo_tts_speed.pack(side="left", padx=(0, 6), pady=5)
        self.combo_tts_speed.bind("<<ComboboxSelected>>", self._on_tts_speed_changed)

        # Nút hiệu chuẩn lọc ồn ao nuôi (adjust_for_ambient_noise)
        self.btn_calibrate = tk.Button(
            panel,
            text="🌊 Hiệu Chuẩn Lọc Ồn Ao (Sục Khí / Quạt Nước)",
            font=("Segoe UI", 9, "bold"),
            bg="#0F2C59",
            fg="#FFFFFF",
            activebackground="#1E3E62",
            activeforeground="#FFFFFF",
            padx=10,
            pady=3,
            relief="ridge",
            cursor="hand2",
            command=self._calibrate_noise
        )
        self.btn_calibrate.pack(side="left", padx=5, pady=5)

        # Thanh trượt độ nhạy lọc ồn (Energy Threshold)
        lbl_thresh = tk.Label(panel, text="Ngưỡng lọc ồn:", font=("Segoe UI", 9), bg="#E8EEF5", fg="#333333")
        lbl_thresh.pack(side="left", padx=(12, 3), pady=5)

        self.scale_noise = tk.Scale(
            panel,
            from_=100,
            to=3500,
            orient="horizontal",
            length=110,
            bg="#E8EEF5",
            highlightthickness=0,
            command=self._on_noise_slider_changed
        )
        self.scale_noise.set(700)
        self.scale_noise.pack(side="left", padx=2, pady=2)

        self.lbl_noise_val = tk.Label(panel, text="700", font=("Segoe UI", 9, "bold"), bg="#E8EEF5", fg="#0F2C59", width=4)
        self.lbl_noise_val.pack(side="left", padx=(0, 8), pady=5)

        # Chọn micro
        lbl_mic_dev = tk.Label(panel, text="Micro:", font=("Segoe UI", 9), bg="#E8EEF5", fg="#333333")
        lbl_mic_dev.pack(side="left", padx=(5, 2), pady=5)

        mic_list = GoogleSpeechVoiceListener.get_microphone_list()
        self.mic_combo = ttk.Combobox(panel, values=mic_list, state="readonly", width=18)
        if mic_list:
            self.mic_combo.current(0)
        self.mic_combo.pack(side="left", padx=2, pady=5)
        self.mic_combo.bind("<<ComboboxSelected>>", self._on_mic_device_changed)

    def _on_tts_toggle(self):
        self.tts.enabled = self.var_tts_enabled.get()
        state = "BẬT" if self.tts.enabled else "TẮT"
        self.lbl_transcript.configure(text=f"Đã {state} tính năng đọc giọng nói phản hồi (TTS).")

    def _on_tts_speed_changed(self, event=None):
        val = self.combo_tts_speed.get()
        match = re.search(r'x(\d+(?:\.\d+)?)', val)
        if match:
            new_rate = float(match.group(1))
            self.tts.set_speed_rate(new_rate)
            self.lbl_transcript.configure(text=f"⚡ Đã đổi tốc độ đọc giọng nói (TTS) sang: x{new_rate}")

    # =========================================================================
    # MỤC 2 & 3: CHẾ ĐỘ 1 - CÂN TRỌNG LƯỢNG & Ô TÍNH SIZE TÔM TRUNG BÌNH
    # =========================================================================
    def _build_tab1_content(self):
        container = tk.Frame(self.tab1_frame, bg=self.COLOR_BG)
        container.pack(fill="both", expand=True, padx=10, pady=10)

        top_frame = tk.Frame(container, bg=self.COLOR_BG)
        top_frame.pack(fill="x", pady=(0, 8))

        # Khối 1: Cú pháp đọc
        guide_box = tk.LabelFrame(
            top_frame,
            text=" 💡 CÚ PHÁP ĐỌC CHẾ ĐỘ 1 (CÂN MẪU) ",
            font=("Segoe UI", 9, "bold"),
            bg="#FFFFFF",
            fg=self.COLOR_NAVY,
            padx=10,
            pady=6
        )
        guide_box.pack(side="left", fill="both", expand=True, padx=(0, 6))

        guide_text = (
            "• Đọc: [Trọng lượng] + [giới tính] (VD: '35 trai', '35 chai', '45 gái', '50 đực', '30 cái', '38 tai', '25 ái')\n"
            "• Xóa dòng cuối: đọc 'xóa', 'hủy', hoặc 'xóa dòng'  |  Xóa sạch về 0: đọc 'xóa hết', 'xóa tất cả', hoặc 'reset'\n"
            "• Giống Đực (Male): 'trai'/'chai'/'đực'/'tai'  |  Giống Cái (Female): 'cái'/'gái'/'ái'"
        )
        lbl_g = tk.Label(guide_box, text=guide_text, font=("Segoe UI", 9), bg="#FFFFFF", fg="#2C3E50", justify="left")
        lbl_g.pack(anchor="w")

        # Khối 2: Ô NHỎ TÍNH SIZE TÔM TRUNG BÌNH (YÊU CẦU 3)
        size_box = tk.LabelFrame(
            top_frame,
            text=" 🦐 SIZE TÔM TRUNG BÌNH ",
            font=("Segoe UI", 9, "bold"),
            bg="#FFFFFF",
            fg="#B71C1C",
            padx=12,
            pady=6
        )
        size_box.pack(side="left", fill="both", padx=(0, 6))

        self.lbl_mode1_size = tk.Label(
            size_box,
            text="0 con/kg",
            font=("Segoe UI", 16, "bold"),
            bg="#FFFFFF",
            fg="#C62828"
        )
        self.lbl_mode1_size.pack(anchor="center")

        self.lbl_mode1_size_sub = tk.Label(
            size_box,
            text="(= Tổng số con / Tổng trọng lượng)",
            font=("Segoe UI", 8, "italic"),
            bg="#FFFFFF",
            fg="#666666"
        )
        self.lbl_mode1_size_sub.pack(anchor="center")

        # Khối 3: Thống kê mẫu cân
        stats_box = tk.LabelFrame(
            top_frame,
            text=" 📊 THỐNG KÊ MẪU ",
            font=("Segoe UI", 9, "bold"),
            bg="#FFFFFF",
            fg=self.COLOR_NAVY,
            padx=12,
            pady=6
        )
        stats_box.pack(side="right", fill="both")

        self.lbl_mode1_total = tk.Label(stats_box, text="Tổng số mẫu: 0 con", font=("Segoe UI", 10, "bold"), bg="#FFFFFF", fg="#0F2C59")
        self.lbl_mode1_total.pack(anchor="w")

        self.lbl_mode1_avg = tk.Label(stats_box, text="Trọng lượng TB: 0.0 g", font=("Segoe UI", 10, "bold"), bg="#FFFFFF", fg="#28A745")
        self.lbl_mode1_avg.pack(anchor="w")

        self.lbl_mode1_ratio = tk.Label(stats_box, text="Male: 0 | Female: 0 | Chưa rõ: 0", font=("Segoe UI", 9), bg="#FFFFFF", fg="#555555")
        self.lbl_mode1_ratio.pack(anchor="w")

        # Bảng dữ liệu mẫu cân
        table_frame = tk.Frame(container, bg="#FFFFFF", highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        table_frame.pack(fill="both", expand=True)

        columns = ("stt", "weight", "gender", "time")
        self.tree_weight = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")

        self.tree_weight.heading("stt", text="STT")
        self.tree_weight.heading("weight", text="Trọng Lượng (g)")
        self.tree_weight.heading("gender", text="Giới Tính")
        self.tree_weight.heading("time", text="Thời Gian Nhận Diện")

        self.tree_weight.column("stt", width=70, anchor="center")
        self.tree_weight.column("weight", width=160, anchor="e")
        self.tree_weight.column("gender", width=140, anchor="center")
        self.tree_weight.column("time", width=180, anchor="center")

        scrollbar_w = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree_weight.yview)
        self.tree_weight.configure(yscrollcommand=scrollbar_w.set)
        self.tree_weight.pack(side="left", fill="both", expand=True)
        scrollbar_w.pack(side="right", fill="y")

        # Thanh nút thao tác Tab 1
        btn_bar = tk.Frame(container, bg=self.COLOR_BG)
        btn_bar.pack(fill="x", pady=(8, 0))

        btn_del_row = tk.Button(
            btn_bar,
            text="🗑️ Xóa dòng chọn",
            font=("Segoe UI", 9),
            bg="#FFFFFF",
            fg="#C82333",
            relief="solid",
            bd=1,
            cursor="hand2",
            command=self._delete_selected_weight_row
        )
        btn_del_row.pack(side="left", padx=3)

        btn_clear_w = tk.Button(
            btn_bar,
            text="🔄 Làm mới bảng cân",
            font=("Segoe UI", 9),
            bg="#FFFFFF",
            fg="#555555",
            relief="solid",
            bd=1,
            cursor="hand2",
            command=self._clear_all_weight_records
        )
        btn_clear_w.pack(side="left", padx=8)

        lbl_manual = tk.Label(btn_bar, text="Nhập tay dự phòng (g):", font=("Segoe UI", 9), bg=self.COLOR_BG, fg="#333333")
        lbl_manual.pack(side="left", padx=(25, 4))
        self.entry_manual_weight = tk.Entry(btn_bar, font=("Segoe UI", 10), width=8)
        self.entry_manual_weight.pack(side="left", padx=2)

        self.combo_manual_gender = ttk.Combobox(btn_bar, values=["Male", "Female", "Chưa rõ"], state="readonly", width=10)
        self.combo_manual_gender.current(0)
        self.combo_manual_gender.pack(side="left", padx=4)

        btn_add_manual = tk.Button(
            btn_bar,
            text="➕ Thêm mẫu",
            font=("Segoe UI", 9, "bold"),
            bg="#007ACC",
            fg="#FFFFFF",
            cursor="hand2",
            command=self._add_manual_weight
        )
        btn_add_manual.pack(side="left", padx=4)

    # =========================================================================
    # CHẾ ĐỘ 2: ĐẾM SỐ LƯỢNG CỘNG DỒN TỐC ĐỘ CAO (CÓ UNDO)
    # =========================================================================
    def _build_tab2_content(self):
        container = tk.Frame(self.tab2_frame, bg=self.COLOR_BG)
        container.pack(fill="both", expand=True, padx=10, pady=10)

        # HỘP HIỂN THỊ KÍCH THƯỚC RẤT LỚN "TỔNG SỐ TÔM"
        self.counter_card = tk.Frame(container, bg="#0F2C59", highlightbackground="#005A9E", highlightthickness=2, padx=15, pady=12)
        self.counter_card.pack(fill="x", pady=(0, 10))

        lbl_counter_title = tk.Label(
            self.counter_card,
            text="TỔNG SỐ LƯỢNG TÔM ĐÃ ĐẾM",
            font=("Segoe UI", 12, "bold"),
            bg="#0F2C59",
            fg="#80D8FF"
        )
        lbl_counter_title.pack(anchor="center")

        # Ô số kích thước cực lớn (Font 56 bold)
        self.lbl_huge_count = tk.Label(
            self.counter_card,
            text="0",
            font=("Segoe UI", 56, "bold"),
            bg="#0F2C59",
            fg="#FFFFFF"
        )
        self.lbl_huge_count.pack(anchor="center", pady=(0, 2))

        self.lbl_counter_sub = tk.Label(
            self.counter_card,
            text="Mẻ vừa cộng gần nhất: Chưa có  |  Tổng số mẻ: 0 mẻ",
            font=("Segoe UI", 11),
            bg="#0F2C59",
            fg="#B0BEC5"
        )
        self.lbl_counter_sub.pack(anchor="center")

        # Khung nút hoàn tác (Undo) và cộng nhanh
        action_bar = tk.Frame(container, bg=self.COLOR_BG)
        action_bar.pack(fill="x", pady=(0, 8))

        self.btn_undo = tk.Button(
            action_bar,
            text="↩️ HOÀN TÁC MẺ VỪA RỒI (HÔ 'XÓA' HOẶC 'HỦY')",
            font=("Segoe UI", 11, "bold"),
            bg=self.COLOR_RED,
            fg="#FFFFFF",
            activebackground=self.COLOR_RED_HOVER,
            activeforeground="#FFFFFF",
            padx=15,
            pady=8,
            cursor="hand2",
            relief="raised",
            bd=2,
            command=self._undo_last_batch
        )
        self.btn_undo.pack(side="left", padx=3)

        btn_reset_count = tk.Button(
            action_bar,
            text="🔄 Đặt lại về 0",
            font=("Segoe UI", 10),
            bg="#FFFFFF",
            fg="#666666",
            relief="solid",
            bd=1,
            padx=10,
            pady=7,
            cursor="hand2",
            command=self._reset_total_shrimp_count
        )
        btn_reset_count.pack(side="left", padx=8)

        lbl_quick = tk.Label(action_bar, text="Cộng nhanh:", font=("Segoe UI", 9, "bold"), bg=self.COLOR_BG, fg="#333333")
        lbl_quick.pack(side="left", padx=(15, 5))

        for qty in [5, 10, 15, 20, 25, 50]:
            b = tk.Button(
                action_bar,
                text=f"+{qty}",
                font=("Segoe UI", 9, "bold"),
                bg="#E3F2FD",
                fg="#0F2C59",
                relief="groove",
                padx=6,
                pady=4,
                cursor="hand2",
                command=lambda q=qty: self._add_batch_amount(q)
            )
            b.pack(side="left", padx=2)

        # Bảng lịch sử các mẻ vừa cộng
        history_frame = tk.Frame(container, bg="#FFFFFF", highlightbackground=self.COLOR_BORDER, highlightthickness=1)
        history_frame.pack(fill="both", expand=True)

        batch_cols = ("stt", "batch_name", "qty", "cumulative", "time", "status")
        self.tree_batch = ttk.Treeview(history_frame, columns=batch_cols, show="headings", selectmode="browse")

        self.tree_batch.heading("stt", text="STT")
        self.tree_batch.heading("batch_name", text="Mã Mẻ")
        self.tree_batch.heading("qty", text="Số Lượng Mẻ (Con)")
        self.tree_batch.heading("cumulative", text="Tổng Tích Lũy")
        self.tree_batch.heading("time", text="Thời Gian")
        self.tree_batch.heading("status", text="Trạng Thái")

        self.tree_batch.column("stt", width=60, anchor="center")
        self.tree_batch.column("batch_name", width=120, anchor="center")
        self.tree_batch.column("qty", width=160, anchor="e")
        self.tree_batch.column("cumulative", width=160, anchor="e")
        self.tree_batch.column("time", width=160, anchor="center")
        self.tree_batch.column("status", width=140, anchor="center")

        scroll_b = ttk.Scrollbar(history_frame, orient="vertical", command=self.tree_batch.yview)
        self.tree_batch.configure(yscrollcommand=scroll_b.set)
        self.tree_batch.pack(side="left", fill="both", expand=True)
        scroll_b.pack(side="right", fill="y")

    # =========================================================================
    # THANH BÁO CÁO DƯỚI CÙNG & XUẤT EXCEL
    # =========================================================================
    def _build_bottom_panel(self):
        bottom_container = tk.Frame(self, bg=self.COLOR_BG)
        bottom_container.pack(fill="x", padx=15, pady=(0, 10))

        action_row = tk.Frame(bottom_container, bg=self.COLOR_BG)
        action_row.pack(fill="x", pady=(0, 6))

        self.btn_export_excel = tk.Button(
            action_row,
            text="📊 XUẤT BÁO CÁO EXCEL (Bao_Cao_Tom.xlsx)",
            font=("Segoe UI", 11, "bold"),
            bg="#1E7E34",
            fg="#FFFFFF",
            activebackground="#155724",
            activeforeground="#FFFFFF",
            padx=18,
            pady=8,
            relief="raised",
            cursor="hand2",
            bd=2,
            command=self._export_excel_report
        )
        self.btn_export_excel.pack(side="left", padx=2)

        btn_open_folder = tk.Button(
            action_row,
            text="📁 Mở Thư Mục File",
            font=("Segoe UI", 10),
            bg="#FFFFFF",
            fg="#0F2C59",
            relief="solid",
            bd=1,
            padx=12,
            pady=7,
            cursor="hand2",
            command=self._open_output_folder
        )
        btn_open_folder.pack(side="left", padx=8)

        # Ô test giọng nói bằng bàn phím
        test_box = tk.Frame(action_row, bg=self.COLOR_BG)
        test_box.pack(side="right")

        lbl_sim = tk.Label(test_box, text="Gõ thử số (Ví dụ: '6', 'sáu', '45 đực'):", font=("Segoe UI", 9), bg=self.COLOR_BG, fg="#555555")
        lbl_sim.pack(side="left", padx=4)

        self.entry_sim_speech = tk.Entry(test_box, font=("Segoe UI", 10), width=20)
        self.entry_sim_speech.pack(side="left", padx=2)
        self.entry_sim_speech.bind("<Return>", lambda e: self._simulate_speech())

        btn_sim = tk.Button(
            test_box,
            text="Gửi Thử",
            font=("Segoe UI", 9, "bold"),
            bg="#007ACC",
            fg="#FFFFFF",
            cursor="hand2",
            command=self._simulate_speech
        )
        btn_sim.pack(side="left", padx=3)

        status_bar = tk.Frame(bottom_container, bg="#FFFFFF", highlightbackground=self.COLOR_BORDER, highlightthickness=1, padx=10, pady=5)
        status_bar.pack(fill="x")

        self.lbl_transcript = tk.Label(
            status_bar,
            text="🎙️ Nhật ký: Sẵn sàng. Bấm 'BẬT MICRO' để bắt đầu đếm rảnh tay.",
            font=("Segoe UI", 9),
            bg="#FFFFFF",
            fg="#333333",
            anchor="w"
        )
        self.lbl_transcript.pack(side="left", fill="x", expand=True)

        self.lbl_autosave_status = tk.Label(
            status_bar,
            text="💾 Tự động lưu: BẬT",
            font=("Segoe UI", 9, "bold"),
            bg="#FFFFFF",
            fg="#28A745"
        )
        self.lbl_autosave_status.pack(side="right")

    # =========================================================================
    # XỬ LÝ SỰ KIỆN MICRO VÀ LỌC ỒN
    # =========================================================================
    def _toggle_microphone(self):
        if not self.voice_listener.is_listening:
            self.voice_listener.start_listening()
            self.btn_toggle_mic.configure(
                text="⏹️ TẮT MICRO (DỪNG LẮNG NGHE)",
                bg=self.COLOR_RED,
                activebackground=self.COLOR_RED_HOVER
            )
            self.lbl_mic_status_badge.configure(
                text="🟢 MICRO ĐANG BẬT (GOOGLE STT - RẢNH TAY)",
                bg="#D4EDDA",
                fg="#155724"
            )
        else:
            self.voice_listener.stop_listening()
            self.btn_toggle_mic.configure(
                text="🎤 BẬT MICRO (RẢNH TAY)",
                bg=self.COLOR_GREEN,
                activebackground=self.COLOR_GREEN_HOVER
            )
            self.lbl_mic_status_badge.configure(
                text="🔴 MICRO ĐANG TẮT",
                bg="#F8D7DA",
                fg="#721C24"
            )

    def _calibrate_noise(self):
        self.lbl_transcript.configure(text="⏳ Đang đo mức ồn sục khí ao nuôi... Vui lòng không nói...")
        self.voice_listener.calibrate_ambient_noise(duration=1.5)

    def _on_noise_adjusted(self, threshold):
        self.msg_queue.put(('NOISE_ADJUSTED', threshold))

    def _on_noise_slider_changed(self, val):
        t = float(val)
        self.lbl_noise_val.configure(text=f"{int(t)}")
        self.voice_listener.set_energy_threshold(t)

    def _on_mic_device_changed(self, event=None):
        idx = self.mic_combo.current()
        self.voice_listener.set_device_index(idx)
        self.lbl_transcript.configure(text=f"Đã chọn Micro: {self.mic_combo.get()}")

    def _on_audio_status_changed(self, status: str, message: str):
        self.msg_queue.put(('STATUS_UPDATE', (status, message)))

    def _on_speech_text_received(self, text: str):
        global is_processing
        # LỖI 3: Khóa an toàn Debounce - Nếu is_processing là True thì lập tức bỏ qua câu lệnh đó
        if is_processing or getattr(self, 'is_processing', False):
            print(f"[DEBOUNCE LOCK] Bỏ qua câu lệnh bị gọi trùng lặp (Double Trigger): \"{text}\"")
            return

        # Nếu False, lập tức đổi thành True để khóa luồng
        is_processing = True
        self.is_processing = True
        self.msg_queue.put(('SPEECH_TEXT', text))

    # =========================================================================
    # VÒNG LẶP HÀNG ĐỢI SỰ KIỆN UI THREAD & LỌC CÂU NÓI CHUYỆN PHIẾM
    # =========================================================================
    def _process_msg_queue(self):
        try:
            while not self.msg_queue.empty():
                evt_type, data = self.msg_queue.get_nowait()

                if evt_type == 'SPEECH_TEXT':
                    self._dispatch_recognized_speech(data, from_queue=True)

                elif evt_type == 'STATUS_UPDATE':
                    status, message = data
                    self.lbl_mic_status_detail.configure(text=message)
                    if status == "LISTENING":
                        self.lbl_transcript.configure(text=f"🎙️ Đang nghe... (Đọc số hoặc lệnh)")
                    elif status == "PROCESSING":
                        self.lbl_transcript.configure(text=f"⚡ {message}")
                    elif status == "ERROR":
                        self.lbl_transcript.configure(text=f"⚠️ {message}")
                    elif status == "READY":
                        self.lbl_transcript.configure(text=f"✅ {message}")

                elif evt_type == 'NOISE_ADJUSTED':
                    val = int(data)
                    self.scale_noise.set(val)
                    self.lbl_noise_val.configure(text=f"{val}")
                    self.lbl_transcript.configure(text=f"✅ Đã thích ứng lọc ồn ao nuôi! Ngưỡng: {val}")

                elif evt_type == 'AUTOSAVE_RESULT':
                    success, msg = data
                    if success:
                        self.lbl_autosave_status.configure(
                            text=f"💾 Đã tự động lưu ({datetime.datetime.now().strftime('%H:%M:%S')})",
                            fg="#28A745"
                        )
                    else:
                        self.lbl_autosave_status.configure(
                            text=f"⚠️ Lỗi lưu: {msg[:25]}",
                            fg="#DC3545"
                        )

        except Exception as e:
            print(f"Lỗi process queue: {e}")

        self.after(60, self._process_msg_queue)

    def _dispatch_recognized_speech(self, text: str, from_queue: bool = False, force: bool = False):
        """
        Phân phối chuỗi giọng nói vào Tab hiện tại đang mở.
        Tiến hành bóc tách dữ liệu, cập nhật giao diện, chèn file Excel,
        rồi đợi khoảng 0.5 giây mới mở khóa (is_processing = False) cho câu tiếp theo.
        """
        global is_processing
        if not from_queue and not force:
            if is_processing or getattr(self, 'is_processing', False):
                print(f"[DEBOUNCE LOCK] Đang bận xử lý, bỏ qua trigger trực tiếp: \"{text}\"")
                return
            is_processing = True
            self.is_processing = True

        try:
            # 1. Kiểm tra bộ lọc câu nói chuyện phiếm ngoài ao
            if not is_relevant_shrimp_command(text):
                self.lbl_transcript.configure(text=f"💬 Bỏ qua câu nói chuyện phiếm: \"{text}\"")
                return

            self.lbl_transcript.configure(text=f"🗣️ Google đã nghe: \"{text}\"")

            clean_cmd = text.lower().strip()
            clean_cmd = re.sub(r'[\.,!?]', '', clean_cmd).strip()

            # -----------------------------------------------------------------
            # CHỨC NĂNG 2: XÓA SẠCH TOÀN BỘ (RESET VỀ 0)
            # Từ khóa kích hoạt: "xóa hết", "xóa tất cả", "reset"
            # -----------------------------------------------------------------
            is_reset_cmd = False
            if clean_cmd in ("xóa hết", "xoa het", "xóa tất cả", "xoa tat ca", "reset", "xóa toàn bộ", "xoa toan bo"):
                is_reset_cmd = True
            elif re.search(r'^(xóa hết|xoa het|xóa tất cả|xoa tat ca|xóa toàn bộ|xoa toan bo|reset)$', clean_cmd):
                is_reset_cmd = True

            if is_reset_cmd:
                self._voice_reset_all_data()
                return

            # -----------------------------------------------------------------
            # CHỨC NĂNG 1: XÓA DÒNG CUỐI CÙNG (SỬA SAI RẢNH TAY)
            # Từ khóa kích hoạt: "xóa", "hủy", "xóa dòng"
            # -----------------------------------------------------------------
            is_undo_cmd = False
            if clean_cmd in ("xóa", "hủy", "xóa dòng", "xoa", "huy", "xoa dong", "hủy dòng", "huy dong", "undo"):
                is_undo_cmd = True
            elif re.search(r'^(xóa|hủy|xóa dòng|xoa|huy|xoa dong|hủy dòng|huy dong|undo)$', clean_cmd):
                is_undo_cmd = True
            elif ("xóa dòng" in clean_cmd or "xoa dong" in clean_cmd or "hủy dòng" in clean_cmd) and not re.search(r'\d', clean_cmd):
                is_undo_cmd = True
            elif ("xóa" in clean_cmd or "hủy" in clean_cmd or "undo" in clean_cmd) and not re.search(r'\d', clean_cmd) and not any(k in clean_cmd for k in ("hết", "tất cả", "toàn bộ", "tat ca", "toan bo")):
                is_undo_cmd = True

            if is_undo_cmd:
                self._voice_delete_last_row()
                return

            active_tab_idx = self.notebook.index(self.notebook.select())

            if active_tab_idx == 0:
                self._handle_mode1_speech(text)
            else:
                self._handle_mode2_speech(text)

        finally:
            # Đợi khoảng 0.5 giây mới mở khóa (is_processing = False) cho câu tiếp theo
            self._schedule_debounce_unlock(0.5)

    def _schedule_debounce_unlock(self, delay_seconds: float = 0.5):
        """Đợi khoảng 0.5 giây mới mở khóa (is_processing = False) cho câu tiếp theo."""
        def _unlock_worker():
            time.sleep(delay_seconds)
            self._unlock_processing()

        threading.Thread(target=_unlock_worker, daemon=True).start()

    def _unlock_processing(self):
        global is_processing
        is_processing = False
        self.is_processing = False

    # =========================================================================
    # LỆNH GIỌNG NÓI: XÓA DÒNG CUỐI & XÓA SẠCH VỀ 0 CHO CẢ 2 CHẾ ĐỘ
    # =========================================================================
    def _voice_delete_last_row(self):
        """
        CHỨC NĂNG 1 - XÓA DÒNG CUỐI CÙNG (SỬA SAI RẢNH TAY):
        - Từ khóa kích hoạt: Nếu văn bản nhận diện được là "xóa", "hủy", hoặc "xóa dòng".
        - Hành động: Phần mềm tự động loại bỏ (Delete) dòng dữ liệu cuối cùng vừa mới chèn
          vào bảng hiển thị và file Excel của chế độ đang mở.
        - Phản hồi: Phát âm thanh cảnh báo hoặc cho AI đọc loa (TTS): "Đã xóa dòng cuối".
        """
        active_tab_idx = self.notebook.index(self.notebook.select())

        if active_tab_idx == 0:
            # Chế độ 1: Cân mẫu
            if not self.weight_records:
                play_audio_feedback('warning')
                self.lbl_transcript.configure(text="⚠️ Bảng cân mẫu trống, không có dòng nào để xóa!")
                return

            play_audio_feedback('undo')
            removed = self.weight_records.pop()
            self._refresh_tree_weight()
            self._update_mode1_stats()
            self._trigger_auto_save()

            self.lbl_transcript.configure(
                text=f"↩️ Đã xóa dòng cuối: Mẫu {removed.get('weight', 0):g}g - {removed.get('gender', 'Chưa rõ')}"
            )
            self.tts.speak_async("Đã xóa dòng cuối")

        else:
            # Chế độ 2: Đếm cộng dồn
            if not self.batch_records:
                play_audio_feedback('warning')
                self.lbl_transcript.configure(text="⚠️ Bảng mẻ đếm trống, không có dòng nào để xóa!")
                return

            play_audio_feedback('undo')
            last_record = self.batch_records.pop()
            qty_to_revert = last_record.get('count', 0)
            batch_name = last_record.get('batch_name', f"Mẻ {self.batch_counter}")

            self.total_batch_count = max(0, self.total_batch_count - qty_to_revert)
            self.batch_counter = max(0, self.batch_counter - 1)

            children = self.tree_batch.get_children()
            if children:
                self.tree_batch.delete(children[-1])

            self._update_huge_count_display(highlight=False)
            self._trigger_auto_save()

            self.lbl_transcript.configure(
                text=f"↩️ Đã xóa dòng cuối: Đã trừ {qty_to_revert} con ({batch_name})! Tổng còn: {self.total_batch_count:,} con"
            )
            self.tts.speak_async("Đã xóa dòng cuối")

    def _voice_reset_all_data(self):
        """
        CHỨC NĂNG 2 - XÓA SẠCH TOÀN BỘ (RESET VỀ 0):
        - Từ khóa kích hoạt: Nếu văn bản nhận diện được là "xóa hết", "xóa tất cả", hoặc "reset".
        - Hành động: Xóa sạch sành sanh tất cả các dòng dữ liệu đang có trong bảng hiển thị,
          đặt lại biến Tổng tích lũy số lượng tôm và số lượng mẻ về lại bằng 0.
        - Phản hồi an toàn: Trước khi xóa sạch, cho AI phát giọng nói đọc loa (TTS):
          "Đang xóa toàn bộ dữ liệu" để cảnh báo người dùng.
        """
        # Phản hồi an toàn qua TTS trước khi xóa sạch
        self.tts.speak_async("Đang xóa toàn bộ dữ liệu")
        play_audio_feedback('undo')

        active_tab_idx = self.notebook.index(self.notebook.select())

        if active_tab_idx == 0:
            # Chế độ 1: Cân mẫu
            self.weight_records.clear()
            self._refresh_tree_weight()
            self._update_mode1_stats()
            self._trigger_auto_save()
            self.lbl_transcript.configure(
                text="🗑️ Đang xóa toàn bộ dữ liệu: Đã làm sạch toàn bộ bảng cân mẫu về trống!"
            )
        else:
            # Chế độ 2: Đếm cộng dồn
            self.total_batch_count = 0
            self.batch_counter = 0
            self.batch_records.clear()
            self.tree_batch.delete(*self.tree_batch.get_children())
            self._update_huge_count_display(highlight=False)
            self._trigger_auto_save()
            self.lbl_transcript.configure(
                text="🗑️ Đang xóa toàn bộ dữ liệu: Đã đặt lại tổng số tôm và số mẻ về 0!"
            )

    # =========================================================================
    # CHẾ ĐỘ 1: CÂN TRỌNG LƯỢNG, GIỚI TÍNH & ĐỌC LẠI QUA TTS
    # =========================================================================
    def _handle_mode1_speech(self, text: str):
        """
        Xử lý dữ liệu giọng nói cho Chế độ 1:
        1. Nhận diện các lệnh điều khiển rảnh tay:
           - "xóa hết", "xóa tất cả", "reset" -> Xóa sạch toàn bộ dữ liệu
           - "xóa", "hủy", "xóa dòng" -> Xóa dòng cuối cùng
        2. Nhận diện giới tính:
           - "cái", "gái", "ái" -> "Female"
           - "trai", "chai", "đực", "tai" -> "Male"
           - Không nghe thấy -> "Chưa rõ"
        3. Bóc tách số trọng lượng (g).
        """
        if not text:
            return

        t_lower = text.lower().strip()
        t_clean_cmd = re.sub(r'[\.,!?]', '', t_lower).strip()

        # Bắt lệnh Xóa Sạch Toàn Bộ (Reset về 0)
        if t_clean_cmd in ("xóa hết", "xoa het", "xóa tất cả", "xoa tat ca", "reset", "xóa toàn bộ", "xoa toan bo"):
            self._voice_reset_all_data()
            return

        # Bắt lệnh Xóa Dòng Cuối (Sửa sai rảnh tay)
        if t_clean_cmd in ("xóa", "hủy", "xóa dòng", "xoa", "huy", "xoa dong", "hủy dòng", "huy dong", "undo"):
            self._voice_delete_last_row()
            return
        if ("xóa dòng" in t_clean_cmd or "xoa dong" in t_clean_cmd) and not re.search(r'\d', t_clean_cmd):
            self._voice_delete_last_row()
            return

        # 1. Nhận diện giới tính trước
        gender = detect_shrimp_gender_mode1(text)

        # 2. Làm sạch từ khóa giới tính để bóc tách số cân chính xác
        t_clean = re.sub(r'(\d+)([a-zA-Zà-ỹÀ-Ỹ]+)', r'\1 \2', text.lower().strip())
        cleaned_text = re.sub(
            r'\b(cái|cai|gái|gai|ái|female|phi\s*meo|phimeo|pheme|femail|phi|mái|mai|trai|chai|đực|duc|tai|male|mail|meo|mê|mel|trống|trong|gram|gam|g|kg)\b',
            ' ',
            t_clean,
            flags=re.IGNORECASE
        )
        cleaned_text = re.sub(r'\s+', ' ', cleaned_text).strip()

        weight, _ = parse_mode1_weighing_and_gender(cleaned_text)
        if weight is None or weight <= 0:
            # Dự phòng trích xuất từ câu gốc
            weight, _ = parse_mode1_weighing_and_gender(text)

        if weight is None or weight <= 0:
            play_audio_feedback('warning')
            self.lbl_transcript.configure(
                text=f"❓ Không nhận diện được số cân từ: \"{text}\". Hãy đọc dạng '35 trai', '35 chai', '45 gái', '50 đực' hoặc '30 cái'."
            )
            return

        play_audio_feedback('beep')

        stt = len(self.weight_records) + 1
        now_time = datetime.datetime.now().strftime("%H:%M:%S")

        record = {
            'stt': stt,
            'weight': weight,
            'gender': gender,
            'time': now_time
        }
        self.weight_records.append(record)

        item_id = self.tree_weight.insert("", "end", values=(stt, f"{weight:g}", gender, now_time))
        self.tree_weight.see(item_id)

        self._update_mode1_stats()
        self._trigger_auto_save()
        self.lbl_transcript.configure(text=f"✅ Đã thêm mẫu: {weight:g}g - {gender} ({now_time})")

        # MỤC 1: ĐỌC LẠI SỐ LIỆU VỪA NHẬP TIẾNG VIỆT (Ví dụ: "Đã lưu 45 gam Male" hoặc "Đã lưu 35 gam Female")
        if gender in ("Male", "Đực"):
            tts_text = f"Đã lưu {weight:g} gam Male"
        elif gender in ("Female", "Cái"):
            tts_text = f"Đã lưu {weight:g} gam Female"
        else:
            tts_text = f"Đã lưu {weight:g} gam"
        self.tts.speak_async(tts_text)

    def _update_mode1_stats(self):
        """
        MỤC 3: Thống kê tự động - Cập nhật số liệu và tính Size tôm trung bình
        bằng cách lấy tổng số con chia cho tổng trọng lượng.
        """
        total_count = len(self.weight_records)
        if total_count == 0:
            self.lbl_mode1_total.configure(text="Tổng số mẫu: 0 con")
            self.lbl_mode1_avg.configure(text="Trọng lượng TB: 0.0 g")
            self.lbl_mode1_ratio.configure(text="Male: 0 | Female: 0 | Chưa rõ: 0")
            self.lbl_mode1_size.configure(text="0 con/kg")
            self.lbl_mode1_size_sub.configure(text="(= Tổng số con / Tổng trọng lượng)")
            return

        total_weight = sum(r['weight'] for r in self.weight_records)
        avg_weight = (total_weight / total_count) if total_count > 0 else 0
        males = sum(1 for r in self.weight_records if r['gender'] in ('Male', 'Đực'))
        females = sum(1 for r in self.weight_records if r['gender'] in ('Female', 'Cái'))
        unknown = sum(1 for r in self.weight_records if r['gender'] not in ('Male', 'Đực', 'Female', 'Cái'))

        pct_m = (males / total_count * 100) if total_count > 0 else 0
        pct_f = (females / total_count * 100) if total_count > 0 else 0

        # TÍNH SIZE TÔM TRUNG BÌNH = TỔNG SỐ CON / TỔNG TRỌNG LƯỢNG
        size_per_g = (total_count / total_weight) if total_weight > 0 else 0
        size_per_kg = size_per_g * 1000

        self.lbl_mode1_total.configure(text=f"Tổng số mẫu: {total_count} con")
        self.lbl_mode1_avg.configure(text=f"Trọng lượng TB: {avg_weight:.1f} g")
        self.lbl_mode1_ratio.configure(text=f"Male: {males} ({pct_m:.0f}%) | Female: {females} ({pct_f:.0f}%) | Chưa rõ: {unknown}")

        # Hiển thị ô Size tôm nổi bật
        self.lbl_mode1_size.configure(text=f"{size_per_kg:.1f} con/kg")
        self.lbl_mode1_size_sub.configure(text=f"(= {total_count} con / {total_weight:.1f}g = {size_per_g:.4f} con/g)")

    def _delete_selected_weight_row(self):
        selected = self.tree_weight.selection()
        if not selected:
            messagebox.showinfo("Thông báo", "Vui lòng bấm chọn dòng muốn xóa trong bảng.")
            return

        item = selected[0]
        values = self.tree_weight.item(item, "values")
        stt_to_del = int(values[0])

        self.tree_weight.delete(item)
        self.weight_records = [r for r in self.weight_records if r['stt'] != stt_to_del]

        for idx, r in enumerate(self.weight_records, 1):
            r['stt'] = idx
        self._refresh_tree_weight()
        self._update_mode1_stats()
        self._trigger_auto_save()
        play_audio_feedback('undo')

    def _clear_all_weight_records(self):
        if not self.weight_records:
            return
        if messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn xóa toàn bộ lịch sử cân mẫu?"):
            self.weight_records.clear()
            self._refresh_tree_weight()
            self._update_mode1_stats()
            self._trigger_auto_save()

    def _refresh_tree_weight(self):
        self.tree_weight.delete(*self.tree_weight.get_children())
        for r in self.weight_records:
            self.tree_weight.insert("", "end", values=(r['stt'], f"{r['weight']:g}", r['gender'], r['time']))

    def _add_manual_weight(self):
        val_str = self.entry_manual_weight.get().strip().replace(',', '.')
        try:
            val = float(val_str)
            if val <= 0:
                raise ValueError()
        except ValueError:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập số trọng lượng hợp lệ lớn hơn 0.")
            return

        gender = self.combo_manual_gender.get()
        stt = len(self.weight_records) + 1
        now_time = datetime.datetime.now().strftime("%H:%M:%S")

        record = {'stt': stt, 'weight': val, 'gender': gender, 'time': now_time}
        self.weight_records.append(record)
        self.tree_weight.insert("", "end", values=(stt, f"{val:g}", gender, now_time))
        self.entry_manual_weight.delete(0, "end")
        self._update_mode1_stats()
        self._trigger_auto_save()
        play_audio_feedback('beep')

        # Đọc TTS
        if gender == "Đực":
            msg = f"Đã lưu {val:g} gam đực"
        elif gender == "Cái":
            msg = f"Đã lưu {val:g} gam cái"
        else:
            msg = f"Đã lưu {val:g} gam"
        self.tts.speak_async(msg)

    # =========================================================================
    # CHẾ ĐỘ 2: ĐẾM SỐ LƯỢNG CỘNG DỒN VỚI BẢNG ÁNH XẠ VÀ ĐỌC TTS
    # =========================================================================
    def _handle_mode2_speech(self, text: str):
        """
        Xử lý giọng nói cho Chế độ 2:
        - LỖI 1: Bảng ánh xạ thủ công (Dictionary) quy đổi tức thì các số từ 1 đến 10.
        - LỖI 2: Hands-free check kiểm tra từ khóa "xóa" hoặc "hủy" để hoàn tác tức thì.
        """
        if not text:
            return

        # -----------------------------------------------------------------
        # BƯỚC 1: Tự động chuyển văn bản nhận diện được thành chữ thường và xóa khoảng trắng
        # -----------------------------------------------------------------
        clean_text = text.lower().strip()
        clean_text = re.sub(r'[\.,!?]', '', clean_text).strip()

        # -----------------------------------------------------------------
        # LỆNH ĐIỀU KHIỂN RẢNH TAY BẰNG GIỌNG NÓI CHO CHẾ ĐỘ 2:
        # 1. XÓA SẠCH TOÀN BỘ (RESET VỀ 0): "xóa hết", "xóa tất cả", "reset"
        # 2. XÓA DÒNG CUỐI CÙNG (SỬA SAI): "xóa", "hủy", "xóa dòng"
        # -----------------------------------------------------------------
        if clean_text in ("xóa hết", "xoa het", "xóa tất cả", "xoa tat ca", "reset", "xóa toàn bộ", "xoa toan bo"):
            self._voice_reset_all_data()
            return

        if clean_text in ("xóa", "hủy", "xóa dòng", "xoa", "huy", "xoa dong", "hủy dòng", "undo"):
            self._voice_delete_last_row()
            return

        if ("xóa dòng" in clean_text or "xoa dong" in clean_text) and not re.search(r'\d', clean_text):
            self._voice_delete_last_row()
            return

        if ("xóa" in clean_text or "hủy" in clean_text or "undo" in clean_text) and not re.search(r'\d', clean_text) and not any(k in clean_text for k in ("hết", "tất cả", "toàn bộ", "tat ca", "toan bo")):
            self._voice_delete_last_row()
            return

        # -----------------------------------------------------------------
        # LỖI 1: BẢNG ÁNH XẠ THỦ CÔNG (DICTIONARY) QUY ĐỔI NGAY LẬP TỨC CÁC SỐ TỪ 1 ĐẾN 10
        # "một"/"mốt" -> 1, "hai" -> 2, "ba" -> 3, "bốn" -> 4, "năm"/"lăm" -> 5,
        # "sáu" -> 6, "bảy" -> 7, "tám" -> 8, "chín" -> 9, "mười"/"chục" -> 10.
        # -----------------------------------------------------------------
        mapping_1_to_10 = {
            "một": 1, "mốt": 1, "mot": 1, "1": 1,
            "hai": 2, "2": 2,
            "ba": 3, "3": 3,
            "bốn": 4, "bon": 4, "tư": 4, "tu": 4, "4": 4,
            "năm": 5, "lăm": 5, "nam": 5, "lam": 5, "nhăm": 5, "5": 5,
            "sáu": 6, "sau": 6, "6": 6,
            "bảy": 7, "bẩy": 7, "bay": 7, "7": 7,
            "tám": 8, "tam": 8, "8": 8,
            "chín": 9, "chin": 9, "9": 9,
            "mười": 10, "chục": 10, "muoi": 10, "mươi": 10, "chuc": 10, "10": 10,
        }

        # 1. Tra cứu trực tiếp từ đơn
        if clean_text in mapping_1_to_10:
            qty = mapping_1_to_10[clean_text]
            self._add_batch_amount(qty)
            return

        # 2. Tra cứu sau khi loại bỏ các từ đệm khẩu ngữ ngoài ao ("mẻ", "thêm", "cộng", "con", "tôm", "bốc")
        words_no_filler = re.sub(r'\b(mẻ|me|thêm|them|cộng|cong|con|tôm|tom|bốc|boc|đợt|dot)\b', '', clean_text).strip()
        words_no_filler = re.sub(r'\s+', ' ', words_no_filler)
        if words_no_filler in mapping_1_to_10:
            qty = mapping_1_to_10[words_no_filler]
            self._add_batch_amount(qty)
            return

        # -----------------------------------------------------------------
        # CÁC SỐ LỚN HƠN 10 (Ví dụ: "mười lăm", "hai mươi", "hai lăm", "ba lăm", "50"...)
        # -----------------------------------------------------------------
        action, qty = map_spoken_text_to_batch(clean_text)
        if action == 'UNDO':
            self._undo_last_batch()
            return

        if action == 'ADD' and qty is not None and qty > 0:
            self._add_batch_amount(qty)
            return

        # Nếu không nhận diện được
        play_audio_feedback('warning')
        self.lbl_transcript.configure(
            text=f"❓ Không nhận diện được mẻ từ: \"{text}\". Chỉ cần đọc: 'năm', 'sáu', 'mười', 'mười lăm' hoặc 'xóa'."
        )

    def _add_batch_amount(self, amount: int):
        if amount <= 0:
            return

        play_audio_feedback('beep')

        self.batch_counter += 1
        self.total_batch_count += amount
        now_time = datetime.datetime.now().strftime("%H:%M:%S")

        batch_name = f"Mẻ {self.batch_counter}"
        record = {
            'stt': len(self.batch_records) + 1,
            'batch_name': batch_name,
            'count': amount,
            'cumulative': self.total_batch_count,
            'time': now_time,
            'action': 'ADD'
        }
        self.batch_records.append(record)

        self._update_huge_count_display(highlight=True)

        item_id = self.tree_batch.insert(
            "",
            "end",
            values=(record['stt'], batch_name, f"+{amount}", f"{self.total_batch_count:,}", now_time, "Đã cộng")
        )
        self.tree_batch.see(item_id)

        self._trigger_auto_save()
        self.lbl_transcript.configure(text=f"✅ {batch_name}: +{amount} con -> Tổng tôm hiện tại: {self.total_batch_count:,} con")

        # MỤC 1: ĐỌC LẠI SỐ LIỆU VỪA CỘNG (Ví dụ: "Đã thêm 5 con, tổng 55")
        tts_text = f"Đã thêm {amount} con, tổng {self.total_batch_count}"
        self.tts.speak_async(tts_text)

    def _undo_last_batch(self):
        """
        LỖI 2: Hàm Hoàn tác (Undo) trừ bớt mẻ vừa cộng gần nhất
        và xóa dòng cuối trong bảng lịch sử.
        """
        if not self.batch_records:
            play_audio_feedback('warning')
            self.lbl_transcript.configure(text="⚠️ Không còn mẻ nào trong lịch sử để hoàn tác!")
            return

        play_audio_feedback('undo')

        # 1. Trừ bớt mẻ vừa cộng gần nhất
        last_record = self.batch_records.pop()
        qty_to_revert = last_record.get('count', 0)
        batch_name = last_record.get('batch_name', f"Mẻ {self.batch_counter}")

        self.total_batch_count = max(0, self.total_batch_count - qty_to_revert)
        self.batch_counter = max(0, self.batch_counter - 1)

        # 2. Xóa dòng cuối trong bảng lịch sử (Treeview)
        children = self.tree_batch.get_children()
        if children:
            self.tree_batch.delete(children[-1])

        # 3. Cập nhật số lớn hiển thị
        self._update_huge_count_display(highlight=False)

        # 4. Tự động lưu file Excel
        self._trigger_auto_save()

        # 5. Thông báo trạng thái và phát giọng đọc tiếng Việt xác nhận
        self.lbl_transcript.configure(
            text=f"↩️ ĐÃ XÓA DÒNG CUỐI: Đã trừ {qty_to_revert} con ({batch_name})! Tổng còn: {self.total_batch_count:,} con"
        )
        self.tts.speak_async("Đã xóa dòng cuối")

    def _update_huge_count_display(self, highlight: bool = True):
        self.lbl_huge_count.configure(text=f"{self.total_batch_count:,}")

        valid_batches = [b for b in self.batch_records if b.get('action') == 'ADD']
        total_valid = len(valid_batches)
        last_qty = valid_batches[-1]['count'] if total_valid > 0 else 0

        self.lbl_counter_sub.configure(
            text=f"Mẻ gần nhất: +{last_qty} con  |  Tổng số mẻ hợp lệ: {total_valid} mẻ"
        )

        if highlight:
            self.lbl_huge_count.configure(fg="#00E676")
            self.after(200, lambda: self.lbl_huge_count.configure(fg="#FFFFFF"))

    def _reset_total_shrimp_count(self):
        if self.total_batch_count == 0 and not self.batch_records:
            return
        if messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn đặt lại Tổng số tôm về 0?"):
            self.total_batch_count = 0
            self.batch_counter = 0
            self.batch_records.clear()
            self.tree_batch.delete(*self.tree_batch.get_children())
            self._update_huge_count_display()
            self._trigger_auto_save()
            play_audio_feedback('beep')
            self.tts.speak_async("Đã đặt lại tổng số tôm về không")

    # =========================================================================
    # TỰ ĐỘNG LƯU VÀ XUẤT BÁO CÁO EXCEL
    # =========================================================================
    def _trigger_auto_save(self):
        pond_name = self.entry_pond.get().strip()
        staff_name = self.entry_staff.get().strip()

        weights_copy = list(self.weight_records)
        batches_copy = list(self.batch_records)
        total_copy = self.total_batch_count

        def _do_save():
            success, final_path, msg = save_shrimp_report(
                self.excel_path,
                pond_name,
                staff_name,
                weights_copy,
                batches_copy,
                total_copy
            )
            # Gửi kết quả về UI Thread một cách an toàn qua hàng đợi msg_queue
            self.msg_queue.put(('AUTOSAVE_RESULT', (success, msg)))

        threading.Thread(target=_do_save, daemon=True).start()

    def _export_excel_report(self):
        pond_name = self.entry_pond.get().strip()
        staff_name = self.entry_staff.get().strip()

        if not pond_name:
            pond_name = "Ao_Mac_Dinh"
        if not staff_name:
            staff_name = "Nhan_Vien"

        timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        safe_pond = "".join(c for c in pond_name if c.isalnum() or c in (' ', '_', '-')).strip().replace(' ', '_')
        default_name = f"Bao_Cao_Tom_{safe_pond}_{timestamp_str}.xlsx"

        file_selected = filedialog.asksaveasfilename(
            parent=self,
            title="Chọn nơi lưu Báo Cáo Excel Tôm",
            initialdir=self.app_dir,
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")]
        )

        if not file_selected:
            return

        success, final_path, msg = save_shrimp_report(
            file_selected,
            pond_name,
            staff_name,
            self.weight_records,
            self.batch_records,
            self.total_batch_count
        )

        if success:
            play_audio_feedback('beep')
            res = messagebox.askyesno(
                "Xuất Excel Thành Công",
                f"Đã xuất báo cáo 2 sheet thành công tại:\n{final_path}\n\nBạn có muốn mở file này ngay bây giờ không?"
            )
            if res:
                self._open_file(final_path)
        else:
            play_audio_feedback('warning')
            messagebox.showerror("Lỗi Xuất File", msg)

    def _open_output_folder(self):
        try:
            os.startfile(self.app_dir)
        except Exception:
            subprocess.Popen(f'explorer "{self.app_dir}"')

    def _open_file(self, file_path: str):
        try:
            os.startfile(file_path)
        except Exception as e:
            messagebox.showinfo("Thông báo", f"Đã lưu file tại:\n{file_path}\nLỗi khi mở: {e}")

    # =========================================================================
    # GIẢ LẬP KIỂM TRA NHANH BẰNG BÀN PHÍM
    # =========================================================================
    def _simulate_speech(self):
        text = self.entry_sim_speech.get().strip()
        if not text:
            return
        self.entry_sim_speech.delete(0, "end")
        self._dispatch_recognized_speech(text)

    def _on_close(self):
        if self.voice_listener.is_listening:
            self.voice_listener.stop_listening()
        self.destroy()

if __name__ == '__main__':
    app = ShrimpSmartCounterApp()
    app.mainloop()
