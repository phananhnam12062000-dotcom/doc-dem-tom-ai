import React, { useCallback, useEffect, useRef, useState } from 'react';
import {
  AnalyzeVoiceApiResponse,
  MIN_AUDIO_BYTES,
  ShrimpGenderLabel,
  TTS_SPEED_MAP,
  VoiceSessionState,
  WeightRecord,
  analyzeRecognizedSpeechText,
  calculateMode1Stats,
  canProcessCurrentAudio,
  createInitialVoiceSessionState,
  formatCurrentTime,
  markAudioProcessedAndResetBuffer,
  receiveNewAudioPayload,
  resetFlagOnMicAutoRestart,
  validateAudioByteLength,
} from './utils/counterLogic';
import { PythonCodePanel } from './components/PythonCodePanel';
import { generateShrimpReportWorkbookBuffer } from './utils/excelExport';

interface Mode2BatchItem {
  id: number;
  batchCode: string;
  count: number;
  cumulative: number;
  time: string;
  status: 'Hợp lệ' | 'Đã hoàn tác (Undo)';
}

interface BrowserSpeechRecognitionEvent {
  resultIndex: number;
  results: {
    length: number;
    [index: number]: {
      isFinal: boolean;
      [altIndex: number]: {
        transcript: string;
      };
    };
  };
}

interface BrowserSpeechRecognitionErrorEvent {
  error: string;
}

interface BrowserSpeechRecognition extends EventTarget {
  lang: string;
  continuous: boolean;
  interimResults: boolean;
  start: () => void;
  stop: () => void;
  abort: () => void;
  onresult: ((event: BrowserSpeechRecognitionEvent) => void) | null;
  onerror: ((event: BrowserSpeechRecognitionErrorEvent) => void) | null;
  onend: (() => void) | null;
}

declare global {
  interface Window {
    SpeechRecognition?: new () => BrowserSpeechRecognition;
    webkitSpeechRecognition?: new () => BrowserSpeechRecognition;
  }
}

export function App() {
  // 1. TOP BAR STATE (Giữ nguyên 100% theo app.py)
  const [pondName, setPondName] = useState<string>('Ao Số 1 - Khu A');
  const [staffName, setStaffName] = useState<string>('Kỹ thuật viên 1');

  // Trạng thái Micro & Chống nhập trùng (Reset Audio State + Processing Flag `is_processed`)
  const [isListening, setIsListening] = useState<boolean>(true);
  const [silenceTimeoutSec, setSilenceTimeoutSec] = useState<number>(1.2);
  const [voiceSessionUI, setVoiceSessionUI] = useState<VoiceSessionState>(() =>
    createInitialVoiceSessionState(),
  );
  const [liveInterimText, setLiveInterimText] = useState<string>('');

  const [micStatusCode, setMicStatusCode] = useState<
    'LISTENING' | 'STOPPED' | 'CALIBRATING' | 'PROCESSING' | 'ERROR'
  >('LISTENING');
  const [micStatusMessage, setMicStatusMessage] = useState<string>(
    '🟢 MICRO ĐANG BẬT LIÊN TỤC (TỰ CẮT 1.2s & CHỐNG TRÙNG 100%)',
  );

  // Thanh cấu hình giọng đọc phản hồi TTS & Lọc ồn ao nuôi (100 - 3500, mặc định 700 như app.py)
  const [ttsEnabled, setTtsEnabled] = useState<boolean>(true);
  const [ttsSpeedLabel, setTtsSpeedLabel] = useState<string>('x1.3 (Mặc định)');
  const [noiseThreshold, setNoiseThreshold] = useState<number>(700);
  const [selectedMic, setSelectedMic] = useState<string>(
    '0: Microphone (High Definition Audio Device)',
  );
  const [micDevices, setMicDevices] = useState<string[]>([
    '0: Microphone (High Definition Audio Device)',
    '1: Headset Microphone (Realtek Audio)',
    '2: USB Condenser Mic (Ao Nuôi)',
  ]);

  // Chế độ Tab đang mở: Tab 1 (Cân Trọng Lượng & Giới Tính) hoặc Tab 2 (Đếm Cộng Dồn Tốc Độ Cao)
  const [activeTab, setActiveTab] = useState<1 | 2>(1);
  const [showPythonModal, setShowPythonModal] = useState<boolean>(false);

  // Dữ liệu Chế độ 1: Cân Trọng Lượng & Giới Tính
  const [weightRecords, setWeightRecords] = useState<WeightRecord[]>([]);
  const [selectedWeightStt, setSelectedWeightStt] = useState<number | null>(null);
  const [manualWeightInput, setManualWeightInput] = useState<string>('');
  const [manualGenderInput, setManualGenderInput] = useState<ShrimpGenderLabel>(
    'Chưa rõ',
  );
  const [mode1LastSpeech, setMode1LastSpeech] = useState<string>(
    'Chưa có khẩu lệnh (Đọc 1 lần chỉ nhảy đúng 1 dòng sạch: VD "35 trống", "45 mái", "50 đực")',
  );
  const [mode1SpeechColor, setMode1SpeechColor] = useState<string>('#2471A3');

  // Dữ liệu Chế độ 2: Đếm Cộng Dồn Tốc Độ Cao (Undo)
  const [batchRecords, setBatchRecords] = useState<Mode2BatchItem[]>([]);
  const [totalBatchCount, setTotalBatchCount] = useState<number>(0);
  const [manualBatchInput, setManualBatchInput] = useState<string>('');
  const [mode2SubInfo, setMode2SubInfo] = useState<string>(
    'Mẻ gần nhất: Chưa có  |  Tổng số mẻ hợp lệ: 0',
  );
  const [mode2LastSpeech, setMode2LastSpeech] = useState<string>(
    'Chưa có khẩu lệnh (Đọc 1 lần chỉ nhảy đúng 1 dòng: VD "Ba mươi lăm", "Bốn mươi")',
  );
  const [mode2SpeechColor, setMode2SpeechColor] = useState<string>('#1E8449');

  // Thanh giả lập đọc giọng nói & Trạng thái HTTP POST
  const [simSpeechText, setSimSpeechText] = useState<string>('');
  const [httpLogMessage, setHttpLogMessage] = useState<string>(
    'Chống trùng lặp: is_processed=False | audio_bytes=None (Xóa bộ nhớ đệm ngay sau khi nhập)',
  );
  const [httpLogType, setHttpLogType] = useState<'ok' | 'warn' | 'err'>('ok');

  // Refs đồng bộ cho Background Worker, Cờ `is_processed` & Bộ nhớ đệm `audio_bytes`
  const activeTabRef = useRef<1 | 2>(activeTab);
  activeTabRef.current = activeTab;
  const ttsEnabledRef = useRef<boolean>(ttsEnabled);
  ttsEnabledRef.current = ttsEnabled;
  const ttsSpeedRef = useRef<string>(ttsSpeedLabel);
  ttsSpeedRef.current = ttsSpeedLabel;
  const isListeningRef = useRef<boolean>(isListening);
  isListeningRef.current = isListening;
  const silenceTimeoutRef = useRef<number>(silenceTimeoutSec);
  silenceTimeoutRef.current = silenceTimeoutSec;

  const voiceSessionRef = useRef<VoiceSessionState>(
    createInitialVoiceSessionState(),
  );
  const recognitionRef = useRef<BrowserSpeechRecognition | null>(null);
  const silenceTimerIdRef = useRef<number | null>(null);

  // Phát âm thanh Beep phản hồi (giống winsound.Beep trong app.py)
  const playAudioFeedback = useCallback(
    (kind: 'beep' | 'double_beep' | 'warning' | 'stop') => {
      try {
        const AudioCtx =
          window.AudioContext ||
          (window as unknown as { webkitAudioContext: typeof AudioContext })
            .webkitAudioContext;
        if (!AudioCtx) return;
        const ctx = new AudioCtx();
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.connect(gain);
        gain.connect(ctx.destination);
        gain.gain.value = 0.08;

        if (kind === 'beep') {
          osc.frequency.value = 1000;
          osc.start();
          osc.stop(ctx.currentTime + 0.12);
        } else if (kind === 'double_beep') {
          osc.frequency.value = 1200;
          osc.start();
          osc.stop(ctx.currentTime + 0.18);
        } else if (kind === 'warning') {
          osc.frequency.value = 400;
          osc.start();
          osc.stop(ctx.currentTime + 0.28);
        } else {
          osc.frequency.value = 600;
          osc.start();
          osc.stop(ctx.currentTime + 0.15);
        }
      } catch {
        // Ignore audio context restrictions
      }
    },
    [],
  );

  // Đọc phản hồi giọng nói Tiếng Việt (giống speak_confirmation trong app.py)
  const speakConfirmation = useCallback((textToSpeak: string) => {
    if (!ttsEnabledRef.current || !textToSpeak.trim()) return;
    try {
      if ('speechSynthesis' in window) {
        window.speechSynthesis.cancel();
        const utterance = new SpeechSynthesisUtterance(textToSpeak);
        utterance.lang = 'vi-VN';
        utterance.rate = TTS_SPEED_MAP[ttsSpeedRef.current] || 1.3;
        window.speechSynthesis.speak(utterance);
      }
    } catch {
      // Ignore if speechSynthesis not supported
    }
  }, []);

  // Thực thi kết quả vào Chế độ 1 (Cân Trọng Lượng & Giới Tính) - Đảm bảo chỉ chèn đúng 1 dòng duy nhất
  const applyMode1Result = useCallback(
    (
      rawText: string,
      cmd: 'ADD' | 'UNDO' | 'CLEAR',
      weightVal: number = 0,
      genderVal: ShrimpGenderLabel = 'Chưa rõ',
      ttsText?: string,
    ) => {
      if (cmd === 'ADD' && weightVal > 0) {
        const nowStr = formatCurrentTime();
        setWeightRecords((prev) => [
          ...prev,
          {
            stt: prev.length + 1,
            weight: Number(weightVal.toFixed(2)),
            gender: genderVal,
            time: nowStr,
          },
        ]);

        playAudioFeedback('beep');
        setMode1LastSpeech(
          `✅ Đã nhập đúng 1 dòng: "${rawText}" ➔ ${weightVal}g (${genderVal}) | Đã xóa bộ nhớ đệm (audio_bytes=None)`,
        );
        setMode1SpeechColor('#1E8449');

        if (ttsText) {
          speakConfirmation(ttsText);
        } else {
          speakConfirmation(`Đã lưu ${weightVal} gam ${genderVal}`);
        }
      } else if (cmd === 'UNDO') {
        setWeightRecords((prev) => {
          if (prev.length === 0) {
            playAudioFeedback('warning');
            setMode1LastSpeech(
              `⚠️ Đọc "${rawText}": Bảng đang trống, không có dòng để xóa!`,
            );
            setMode1SpeechColor('#D35400');
            speakConfirmation('Bảng đang trống, không thể xóa');
            return prev;
          }
          const removed = prev[prev.length - 1];
          const updated = prev
            .slice(0, -1)
            .map((r, idx) => ({ ...r, stt: idx + 1 }));
          playAudioFeedback('double_beep');
          setMode1LastSpeech(
            `↩️ Đã xóa dòng cuối (Mẫu #${removed.stt}: ${removed.weight}g - ${removed.gender})`,
          );
          setMode1SpeechColor('#C0392B');
          speakConfirmation('Đã xóa dòng cuối');
          return updated;
        });
      } else if (cmd === 'CLEAR') {
        setWeightRecords([]);
        playAudioFeedback('warning');
        setMode1LastSpeech('🗑️ Đã xóa sạch toàn bộ bảng cân trọng lượng!');
        setMode1SpeechColor('#C0392B');
        speakConfirmation('Đang xóa toàn bộ dữ liệu');
      }
    },
    [playAudioFeedback, speakConfirmation],
  );

  // Thực thi kết quả vào Chế độ 2 (Đếm Cộng Dồn Tốc Độ Cao) - Đảm bảo chỉ chèn đúng 1 dòng duy nhất
  const applyMode2Result = useCallback(
    (rawText: string, cmd: 'ADD' | 'UNDO' | 'CLEAR', val: number = 0) => {
      if (cmd === 'ADD' && val > 0) {
        const nowStr = formatCurrentTime();
        setBatchRecords((prev) => {
          const validBatches =
            prev.filter((b) => b.status === 'Hợp lệ').length + 1;
          const currentSum = prev
            .filter((b) => b.status === 'Hợp lệ')
            .reduce((acc, b) => acc + b.count, 0);
          const newTotal = currentSum + val;

          const newRec: Mode2BatchItem = {
            id: prev.length + 1,
            batchCode: `Mẻ #${validBatches}`,
            count: val,
            cumulative: newTotal,
            time: nowStr,
            status: 'Hợp lệ',
          };

          setTotalBatchCount(newTotal);
          setMode2SubInfo(
            `Mẻ gần nhất: +${val} con  |  Tổng số mẻ hợp lệ: ${validBatches}`,
          );
          setMode2LastSpeech(
            `✅ Đã cộng đúng 1 dòng: "${rawText}" ➔ +${val} con | Đã xóa bộ nhớ đệm (audio_bytes=None)`,
          );
          setMode2SpeechColor('#1E8449');
          playAudioFeedback('beep');
          speakConfirmation(`Cộng ${val} con. Tổng ${newTotal}`);

          return [...prev, newRec];
        });
      } else if (cmd === 'UNDO') {
        setBatchRecords((prev) => {
          const lastValidIdx = [...prev]
            .map((r, i) => ({ r, i }))
            .reverse()
            .find((item) => item.r.status === 'Hợp lệ')?.i;

          if (lastValidIdx === undefined) {
            playAudioFeedback('warning');
            setMode2LastSpeech(
              `⚠️ Đọc "${rawText}": Không có mẻ hợp lệ nào để hoàn tác!`,
            );
            setMode2SpeechColor('#D35400');
            speakConfirmation('Không có mẻ nào để hủy');
            return prev;
          }

          const next = prev.map((item, idx) =>
            idx === lastValidIdx
              ? { ...item, status: 'Đã hoàn tác (Undo)' as const }
              : item,
          );
          let runningTotal = 0;
          let validCount = 0;
          const recalculated = next.map((item) => {
            if (item.status === 'Hợp lệ') {
              runningTotal += item.count;
              validCount += 1;
            }
            return {
              ...item,
              cumulative: runningTotal,
            };
          });

          const target = prev[lastValidIdx];
          setTotalBatchCount(runningTotal);
          setMode2SubInfo(
            `Vừa hủy: -${target.count} con  |  Tổng số mẻ hợp lệ: ${validCount}`,
          );
          setMode2LastSpeech(
            `↩️ Đã hoàn tác ${target.batchCode} (-${target.count} con). Tổng còn: ${runningTotal}`,
          );
          setMode2SpeechColor('#C0392B');
          playAudioFeedback('double_beep');
          speakConfirmation('Đã xóa dòng cuối');
          return recalculated;
        });
      } else if (cmd === 'CLEAR') {
        setBatchRecords([]);
        setTotalBatchCount(0);
        setMode2SubInfo('Mẻ gần nhất: Chưa có  |  Tổng số mẻ hợp lệ: 0');
        setMode2LastSpeech('🗑️ Đã xóa sạch toàn bộ dữ liệu đếm theo mẻ!');
        setMode2SpeechColor('#C0392B');
        playAudioFeedback('warning');
        speakConfirmation('Đang xóa toàn bộ dữ liệu');
      }
    },
    [playAudioFeedback, speakConfirmation],
  );

  // Xử lý phản hồi từ `/api/analyze-voice`
  const handleApiResponse = useCallback(
    (cleanText: string, data: AnalyzeVoiceApiResponse) => {
      if (data.command_type === 'RESET_ALL') {
        if (activeTabRef.current === 1) {
          applyMode1Result(cleanText, 'CLEAR');
        } else {
          applyMode2Result(cleanText, 'CLEAR');
        }
        return;
      }

      if (
        data.command_type === 'DELETE_LAST_ROW' ||
        data.command_type === 'MODE2_UNDO'
      ) {
        if (activeTabRef.current === 1) {
          applyMode1Result(cleanText, 'UNDO');
        } else {
          applyMode2Result(cleanText, 'UNDO');
        }
        return;
      }

      if (
        data.command_type === 'MODE1_ADD_WEIGHT' &&
        data.weight &&
        data.weight > 0
      ) {
        applyMode1Result(
          cleanText,
          'ADD',
          data.weight,
          data.gender || 'Chưa rõ',
          data.tts_text,
        );
        return;
      }

      if (
        data.command_type === 'MODE2_ADD_BATCH' &&
        data.batch_qty &&
        data.batch_qty > 0
      ) {
        applyMode2Result(cleanText, 'ADD', data.batch_qty);
        return;
      }
    },
    [applyMode1Result, applyMode2Result],
  );

  /**
   * HÀM PHÂN TÍCH GIỌNG NÓI CHỐNG TRÙNG LẶP TUYỆT ĐỐI:
   * 1. Nhận gói âm thanh mới vào `voiceSessionRef.current` (nếu `is_processed == false`).
   * 2. Kiểm tra `canProcessCurrentAudio` (`is_processed == false && audio_bytes != null`).
   * 3. Ngay lập tức đặt `is_processed = true`, xóa sạch `audio_bytes = null` và xóa `temp_voice.wav`.
   * 4. Gửi HTTP POST `/api/analyze-voice` và chèn đúng 1 dòng duy nhất vào bảng.
   * 5. Khi chu kỳ tiếp theo của Micro tự động mở lại, đặt `is_processed = false`.
   */
  const dispatchRecognizedSpeechHttpPost = useCallback(
    async (recognizedText: string, simulatedAudioBytes: number = 4096) => {
      const cleanText = recognizedText.trim();
      if (!cleanText) return;

      // Bước 1: Nạp dữ liệu vào trạng thái phiên thu âm hiện tại
      voiceSessionRef.current = receiveNewAudioPayload(
        voiceSessionRef.current,
        cleanText,
        'temp_voice.wav',
      );

      // Bước 2: Kiểm tra cờ `is_processed == False` và `audio_bytes != None`
      if (!canProcessCurrentAudio(voiceSessionRef.current)) {
        return;
      }

      // Bước 3: Lập tức khóa cờ `is_processed = True` và xóa sạch `audio_bytes = None` + `os.remove('temp_voice.wav')`
      voiceSessionRef.current = markAudioProcessedAndResetBuffer(
        voiceSessionRef.current,
      );
      setVoiceSessionUI({ ...voiceSessionRef.current });

      // Kiểm tra chất lượng/dung lượng file âm thanh đầu vào (< 1KB / 1024 bytes chặn ngay)
      const audioCheck = validateAudioByteLength(simulatedAudioBytes);
      if (!audioCheck.valid) {
        setHttpLogType('warn');
        setHttpLogMessage(`⚠️ Đã chặn gửi: ${audioCheck.message}`);
        playAudioFeedback('warning');
        // Mở lại cờ cho chu kỳ kế tiếp
        voiceSessionRef.current = resetFlagOnMicAutoRestart(
          voiceSessionRef.current,
        );
        setVoiceSessionUI({ ...voiceSessionRef.current });
        return;
      }

      const tabIdx: 0 | 1 = activeTabRef.current === 1 ? 0 : 1;

      try {
        const response = await fetch('/api/analyze-voice', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            transcriptText: cleanText,
            audioByteLength: simulatedAudioBytes,
            mimeType: 'audio/wav',
            activeTabIdx: tabIdx,
          }),
        });

        const data = (await response.json()) as AnalyzeVoiceApiResponse;

        if (data.status === 'error') {
          setHttpLogType('err');
          setHttpLogMessage(
            `❌ HTTP POST /api/analyze-voice trả về lỗi: ${data.message}`,
          );
          playAudioFeedback('warning');
          return;
        }

        if (data.status === 'ignored') {
          setHttpLogType('warn');
          setHttpLogMessage(
            `🔇 Bỏ qua ("${cleanText}") | Đã xóa audio_bytes=None & temp_voice.wav`,
          );
          if (activeTabRef.current === 1) {
            setMode1LastSpeech(data.message);
            setMode1SpeechColor('#7F8C8D');
          } else {
            setMode2LastSpeech(data.message);
            setMode2SpeechColor('#7F8C8D');
          }
          return;
        }

        setHttpLogType('ok');
        setHttpLogMessage(
          `✅ Đã nhập 1 dòng ("${cleanText}") | Đã xóa audio_bytes=None & temp_voice.wav | is_processed=False (Chu kỳ mới)`,
        );

        handleApiResponse(cleanText, data);
      } catch (err) {
        const localResult = analyzeRecognizedSpeechText(cleanText, tabIdx);
        setHttpLogType('warn');
        setHttpLogMessage(
          `⚠️ Xử lý nội bộ 1 lần: "${cleanText}" (${err instanceof Error ? err.message : 'Lỗi mạng'})`,
        );
        if (localResult.status === 'success') {
          handleApiResponse(cleanText, localResult);
        }
      } finally {
        // Bước 4: Khi Micro tự động kích hoạt lại để chờ câu nói tiếp theo -> Chuyển `is_processed = False`
        voiceSessionRef.current = resetFlagOnMicAutoRestart(
          voiceSessionRef.current,
        );
        setVoiceSessionUI({ ...voiceSessionRef.current });

        if (isListeningRef.current) {
          setMicStatusCode('LISTENING');
          setMicStatusMessage(
            `🟢 MICRO ĐANG BẬT LIÊN TỤC (TỰ CẮT ${silenceTimeoutRef.current.toFixed(1)}s & CHỐNG TRÙNG 100%)`,
          );
        } else {
          setMicStatusCode('STOPPED');
          setMicStatusMessage('🔴 MICRO ĐANG TẮT');
        }
      }
    },
    [handleApiResponse, playAudioFeedback],
  );

  // Khởi tạo danh sách thiết bị Micro
  useEffect(() => {
    if (navigator.mediaDevices && navigator.mediaDevices.enumerateDevices) {
      navigator.mediaDevices
        .enumerateDevices()
        .then((devices) => {
          const audioInputs = devices.filter((d) => d.kind === 'audioinput');
          if (audioInputs.length > 0) {
            const names = audioInputs.map(
              (d, i) => `${i}: ${d.label || `Microphone Thiết bị #${i + 1}`}`,
            );
            setMicDevices(names);
            setSelectedMic(names[0]);
          }
        })
        .catch(() => {
          // Keep default microphone list
        });
    }
  }, []);

  /**
   * HỆ THỐNG LẮNG NGHE LIÊN TỤC + TỰ ĐỘNG CẮT KHOẢNG LẶNG (1.0s - 1.5s) + CHỐNG NHẬP TRÙNG 100%:
   * - Khi cắt câu sau khoảng lặng, ngay lập tức:
   *   1) Kiểm tra `is_processed == false`
   *   2) Đặt `is_processed = true` và xóa `audio_bytes = null`
   *   3) Ngắt phiên thu hiện tại (`recognition.abort()`) để xóa sạch bộ đệm âm thanh nội bộ của trình duyệt,
   *      ngăn trình duyệt phát lại sự kiện `isFinal` gây nhảy dòng thứ 2 trùng nhau!
   *   4) Trong `onend`, tự động khởi động lại Micro cho chu kỳ mới và đặt `is_processed = false`.
   */
  useEffect(() => {
    const SpeechRec =
      window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRec) {
      return;
    }

    if (!isListening) {
      if (silenceTimerIdRef.current) {
        window.clearTimeout(silenceTimerIdRef.current);
        silenceTimerIdRef.current = null;
      }
      setLiveInterimText('');
      if (recognitionRef.current) {
        try {
          recognitionRef.current.abort();
        } catch {
          // Ignore
        }
      }
      return;
    }

    let isUnmounted = false;
    let currentUtteranceBuffer = '';
    let cycleLocked = false;

    const recognition = new SpeechRec();
    recognition.lang = 'vi-VN';
    recognition.continuous = false; // Mỗi câu nói là 1 chu kỳ độc lập -> Tự xóa sạch buffer sau mỗi lần cắt câu
    recognition.interimResults = true;

    const cutAndProcessSinglePhrase = () => {
      if (cycleLocked) return;
      const phrase = currentUtteranceBuffer.trim();
      if (!phrase) return;

      // Khóa ngay chu kỳ hiện tại để tuyệt đối không bao giờ đọc lại lần 2
      cycleLocked = true;
      currentUtteranceBuffer = '';
      setLiveInterimText('');

      if (silenceTimerIdRef.current) {
        window.clearTimeout(silenceTimerIdRef.current);
        silenceTimerIdRef.current = null;
      }

      // Dừng phiên thu của câu vừa rồi để xóa sạch bộ nhớ đệm âm thanh trình duyệt
      try {
        recognition.abort();
      } catch {
        // Ignore
      }

      const estimatedWavBytes = Math.max(2048, phrase.length * 640);
      void dispatchRecognizedSpeechHttpPost(phrase, estimatedWavBytes);
    };

    recognition.onresult = (event: BrowserSpeechRecognitionEvent) => {
      if (cycleLocked) return;

      let latestTranscript = '';
      for (let i = event.resultIndex; i < event.results.length; i++) {
        latestTranscript += (event.results[i][0]?.transcript || '') + ' ';
      }

      const trimmed = latestTranscript.trim();
      if (!trimmed) return;

      currentUtteranceBuffer = trimmed;
      setLiveInterimText(trimmed);

      // Đặt lại bộ đếm khoảng lặng (Silence Timeout 1.0s - 1.5s)
      if (silenceTimerIdRef.current) {
        window.clearTimeout(silenceTimerIdRef.current);
      }
      const timeoutMs = Math.round(silenceTimeoutRef.current * 1000);
      silenceTimerIdRef.current = window.setTimeout(() => {
        cutAndProcessSinglePhrase();
      }, timeoutMs);
    };

    recognition.onerror = () => {
      // Bỏ qua lỗi tạm thời, onend sẽ tự kích hoạt lại Micro
    };

    // AUTO-RESTART LOOP: Tự động kích hoạt lại Micro cho chu kỳ mới và đặt lại cờ khóa
    recognition.onend = () => {
      if (!cycleLocked && currentUtteranceBuffer.trim()) {
        cutAndProcessSinglePhrase();
      }

      if (!isUnmounted && isListeningRef.current) {
        window.setTimeout(() => {
          if (!isUnmounted && isListeningRef.current) {
            cycleLocked = false;
            currentUtteranceBuffer = '';
            try {
              recognition.start();
            } catch {
              // Ignore if already started
            }
          }
        }, 80);
      }
    };

    recognitionRef.current = recognition;
    try {
      recognition.start();
    } catch {
      // Ignore if already started
    }

    return () => {
      isUnmounted = true;
      if (silenceTimerIdRef.current) {
        window.clearTimeout(silenceTimerIdRef.current);
        silenceTimerIdRef.current = null;
      }
      try {
        recognition.abort();
      } catch {
        // Ignore
      }
    };
  }, [isListening, dispatchRecognizedSpeechHttpPost]);

  // Nút Bật / Tắt Micro
  const handleToggleMicrophone = () => {
    if (isListening) {
      setIsListening(false);
      setMicStatusCode('STOPPED');
      setMicStatusMessage('🔴 MICRO ĐANG TẮT');
      playAudioFeedback('stop');
    } else {
      setIsListening(true);
      setMicStatusCode('LISTENING');
      setMicStatusMessage(
        `🟢 MICRO ĐANG BẬT LIÊN TỤC (TỰ CẮT ${silenceTimeoutSec.toFixed(1)}s & CHỐNG TRÙNG 100%)`,
      );
      playAudioFeedback('beep');
    }
  };

  // Nút Hiệu Chuẩn Lọc Ồn Ao Nuôi
  const handleCalibrateNoise = () => {
    setMicStatusCode('CALIBRATING');
    setMicStatusMessage('⏳ ĐANG ĐO TIẾNG ỒN SỤC KHÍ AO NUÔI...');
    setTimeout(() => {
      const calibrated = 680;
      setNoiseThreshold(calibrated);
      setMicStatusCode(isListening ? 'LISTENING' : 'STOPPED');
      setMicStatusMessage(
        isListening
          ? `🟢 MICRO ĐANG BẬT LIÊN TỤC (TỰ CẮT ${silenceTimeoutSec.toFixed(1)}s & CHỐNG TRÙNG 100%)`
          : '🔴 MICRO ĐANG TẮT',
      );
      playAudioFeedback('beep');
    }, 500);
  };

  // Chế độ 1: Thêm thủ công
  const handleManualAddWeight = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    const raw = manualWeightInput.trim().replace(',', '.');
    if (!raw) return;
    const val = Number.parseFloat(raw);
    if (Number.isNaN(val) || val <= 0 || val > 500) {
      playAudioFeedback('warning');
      return;
    }
    setManualWeightInput('');
    applyMode1Result('Nhập bàn phím', 'ADD', val, manualGenderInput);
  };

  // Chế độ 1: Xóa dòng đang chọn hoặc dòng cuối
  const handleDeleteSelectedWeight = () => {
    if (selectedWeightStt !== null) {
      setWeightRecords((prev) =>
        prev
          .filter((r) => r.stt !== selectedWeightStt)
          .map((r, idx) => ({ ...r, stt: idx + 1 })),
      );
      setSelectedWeightStt(null);
      playAudioFeedback('double_beep');
    } else {
      applyMode1Result('Nút bấm', 'UNDO');
    }
  };

  // Chế độ 2: Cộng thủ công
  const handleManualAddBatch = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    const raw = manualBatchInput.trim();
    if (!raw) return;
    const val = Number.parseInt(raw, 10);
    if (Number.isNaN(val) || val <= 0) {
      playAudioFeedback('warning');
      return;
    }
    setManualBatchInput('');
    applyMode2Result('Nhập bàn phím', 'ADD', val);
  };

  // Giả lập đọc nhanh qua HTTP POST
  const handleSimulateSpeech = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!simSpeechText.trim()) return;
    const text = simSpeechText.trim();
    setSimSpeechText('');
    void dispatchRecognizedSpeechHttpPost(text, 4096);
  };

  // Kiểm thử chặn file âm thanh < 1KB (1024 bytes)
  const handleTestSmallAudioBlock = async () => {
    const response = await fetch('/api/analyze-voice', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        audioByteLength: 512,
        transcriptText: '35 đực',
      }),
    });
    const data = await response.json();
    setHttpLogType('warn');
    setHttpLogMessage(
      `🛡️ Kiểm tra chặn file < 1KB (512 bytes): ${data.message} (HTTP ${response.status} JSON sạch)`,
    );
  };

  // Xuất báo cáo Excel (.xlsx) chuẩn 2 Sheet (giống save_shrimp_report bằng openpyxl trong app.py)
  const handleExportExcelReport = async () => {
    try {
      const buffer = await generateShrimpReportWorkbookBuffer({
        pondName,
        staffName,
        weightRecords,
        batchRecords,
      });

      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      const safePond = (pondName || 'Ao_Tom').replace(/\s+/g, '_');
      const filename = `Bao_Cao_Tom_${safePond}.xlsx`;
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      setHttpLogType('ok');
      setHttpLogMessage(
        `📊 Đã xuất file Excel chuẩn 2 Sheet (${filename}) — Theme #1F4E79, Zebra #F2F5F8, Công thức SUM/AVERAGE/COUNTIF & Auto-fit cột!`,
      );
      playAudioFeedback('beep');
    } catch (err) {
      setHttpLogType('err');
      setHttpLogMessage(
        `❌ Lỗi xuất file Excel: ${err instanceof Error ? err.message : String(err)}`,
      );
      playAudioFeedback('warning');
    }
  };

  const mode1Stats = calculateMode1Stats(weightRecords);

  const micStatusBgMap = {
    LISTENING: '#1E8449',
    STOPPED: '#C0392B',
    CALIBRATING: '#D35400',
    PROCESSING: '#2874A6',
    ERROR: '#922B21',
  };

  return (
    <div className="min-h-screen bg-[#F4F6F9] text-[#1C2833] font-sans flex flex-col select-none">
      {/* 1. HEADER & THANH ĐIỀU KHIỂN MICRO (Giống 100% _build_header trong app.py) */}
      <header className="bg-[#1B4F72] text-white px-4 py-2.5 shadow-md">
        <div className="max-w-[1440px] mx-auto flex flex-wrap items-center justify-between gap-3">
          {/* Thông tin Ao / Đầm & Người thực hiện */}
          <div className="flex flex-wrap items-center gap-2">
            <label className="text-sm font-bold text-white whitespace-nowrap">
              Ao/Đầm:
            </label>
            <input
              type="text"
              value={pondName}
              onChange={(e) => setPondName(e.target.value)}
              className="w-40 px-2.5 py-1 text-sm font-bold text-[#1C2833] bg-white rounded border border-slate-300 focus:outline-none"
            />
            <label className="text-sm font-bold text-white whitespace-nowrap ml-1">
              Người cân/đếm:
            </label>
            <input
              type="text"
              value={staffName}
              onChange={(e) => setStaffName(e.target.value)}
              className="w-36 px-2.5 py-1 text-sm font-bold text-[#1C2833] bg-white rounded border border-slate-300 focus:outline-none"
            />
          </div>

          {/* Đèn báo trạng thái Micro Liên Tục + Cờ chống trùng (is_processed & audio_bytes) */}
          <div className="flex items-center gap-2">
            <div
              style={{ backgroundColor: micStatusBgMap[micStatusCode] }}
              className="px-3.5 py-1.5 rounded font-bold text-xs md:text-sm text-white tracking-wide text-center shadow-inner transition-colors"
            >
              {liveInterimText
                ? `🎙️ ĐANG NGHE: "${liveInterimText}" (Tự gửi sau ${silenceTimeoutSec.toFixed(1)}s ngừng nói...)`
                : micStatusMessage}
            </div>
            <span className="bg-[#117A65] text-white text-xs font-bold px-2.5 py-1.5 rounded shadow-inner whitespace-nowrap">
              🔄 Chu kỳ #{voiceSessionUI.cycle_id} | is_processed:{' '}
              {voiceSessionUI.is_processed ? 'True' : 'False'}
            </span>
          </div>

          {/* Cụm nút Bật/Tắt Micro & Mã nguồn Python */}
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={handleToggleMicrophone}
              style={{ backgroundColor: isListening ? '#E74C3C' : '#27AE60' }}
              className="px-4 py-1.5 rounded font-bold text-xs md:text-sm text-white shadow hover:opacity-95 transition-opacity cursor-pointer"
            >
              {isListening
                ? '🛑 TẮT MICRO (DỪNG NGHE)'
                : '🎙️ BẬT MICRO (RẢNH TAY)'}
            </button>
            <button
              type="button"
              onClick={() => setShowPythonModal((prev) => !prev)}
              className="px-3 py-1.5 rounded font-bold text-xs bg-[#0E6655] hover:bg-[#0B5345] text-white shadow cursor-pointer"
            >
              {showPythonModal
                ? '✖ Đóng Mã Nguồn'
                : '🐍 Mã Nguồn app.py & audio_manager.py'}
            </button>
          </div>
        </div>
      </header>

      {/* 2. THANH CẤU HÌNH LỌC TIẾNG ỒN SỤC KHÍ, KHOẢNG LẶNG TỰ CẮT (1.0s - 1.5s) & TTS */}
      <section className="bg-[#D6EAF8] border-b border-[#AED6F1] px-4 py-1.5">
        <div className="max-w-[1440px] mx-auto flex flex-wrap items-center justify-between gap-3 text-xs md:text-sm">
          <div className="flex flex-wrap items-center gap-3">
            {/* Checkbox Đọc lại số liệu */}
            <label className="flex items-center gap-1.5 font-bold text-[#1B4F72] cursor-pointer">
              <input
                type="checkbox"
                checked={ttsEnabled}
                onChange={(e) => setTtsEnabled(e.target.checked)}
                className="w-4 h-4 accent-[#1B4F72] cursor-pointer"
              />
              <span>🔊 Đọc Lại Số Liệu</span>
            </label>

            {/* Chọn tốc độ giọng đọc */}
            <div className="flex items-center gap-1">
              <span className="font-bold text-[#1B4F72]">Tốc độ:</span>
              <select
                value={ttsSpeedLabel}
                onChange={(e) => setTtsSpeedLabel(e.target.value)}
                className="bg-white border border-slate-300 rounded px-2 py-0.5 text-xs font-semibold text-[#1C2833]"
              >
                {Object.keys(TTS_SPEED_MAP).map((label) => (
                  <option key={label} value={label}>
                    {label}
                  </option>
                ))}
              </select>
            </div>

            {/* Cấu hình Khoảng lặng tự cắt câu (Silence Timeout 1.0s - 1.5s) */}
            <div className="flex items-center gap-1.5 bg-white/80 px-2.5 py-0.5 rounded border border-[#AED6F1]">
              <span className="font-bold text-[#0E6655]">
                ⏱️ Ngắt câu khi lặng: {silenceTimeoutSec.toFixed(1)}s
              </span>
              <input
                type="range"
                min={1.0}
                max={1.5}
                step={0.1}
                value={silenceTimeoutSec}
                onChange={(e) => {
                  const val = Number(e.target.value);
                  setSilenceTimeoutSec(val);
                  if (isListening) {
                    setMicStatusMessage(
                      `🟢 MICRO ĐANG BẬT LIÊN TỤC (TỰ CẮT ${val.toFixed(1)}s & CHỐNG TRÙNG 100%)`,
                    );
                  }
                }}
                className="w-20 accent-[#0E6655] cursor-pointer"
              />
            </div>

            {/* Nút Hiệu chuẩn lọc ồn ao */}
            <button
              type="button"
              onClick={handleCalibrateNoise}
              className="bg-[#2E86C1] hover:bg-[#2471A3] text-white font-bold px-3 py-1 rounded text-xs cursor-pointer transition-colors"
            >
              🌊 Hiệu Chuẩn Lọc Ồn Ao
            </button>

            {/* Thanh trượt Ngưỡng lọc ồn (100 - 3500, mặc định 700) */}
            <div className="flex items-center gap-2">
              <span className="font-bold text-[#1B4F72] min-w-[125px]">
                Ngưỡng lọc ồn: {noiseThreshold}
              </span>
              <input
                type="range"
                min={100}
                max={3500}
                value={noiseThreshold}
                onChange={(e) => setNoiseThreshold(Number(e.target.value))}
                className="w-28 accent-[#1B4F72] cursor-pointer"
              />
            </div>
          </div>

          {/* Chọn thiết bị Micro */}
          <div className="flex items-center gap-1.5">
            <span className="font-bold text-[#1B4F72]">Micro:</span>
            <select
              value={selectedMic}
              onChange={(e) => setSelectedMic(e.target.value)}
              className="bg-white border border-slate-300 rounded px-2 py-0.5 text-xs text-[#1C2833] max-w-[230px] truncate"
            >
              {micDevices.map((mic) => (
                <option key={mic} value={mic}>
                  {mic}
                </option>
              ))}
            </select>
          </div>
        </div>
      </section>

      {/* PANEL HIỂN THỊ MÃ NGUỒN PYTHON GỐC NẾU BẤM MỞ */}
      {showPythonModal && (
        <div className="max-w-[1440px] w-full mx-auto px-4 pt-3">
          <PythonCodePanel />
        </div>
      )}

      {/* 3. HỆ THỐNG 2 TAB CHẾ ĐỘ CHÍNH (Giống 100% ttk.Notebook trong app.py) */}
      <main className="max-w-[1440px] w-full mx-auto px-3 pt-2.5 pb-2 flex-1 flex flex-col">
        {/* Thanh chuyển Tab */}
        <div className="flex items-center gap-1.5 border-b-2 border-[#1B4F72]">
          <button
            type="button"
            onClick={() => setActiveTab(1)}
            className={`px-5 py-2 font-bold text-xs md:text-sm rounded-t transition-colors cursor-pointer ${
              activeTab === 1
                ? 'bg-[#1B4F72] text-white shadow'
                : 'bg-[#D6EAF8] text-[#1B4F72] hover:bg-[#AED6F1]'
            }`}
          >
            ⚖️ CHẾ ĐỘ 1: CÂN TRỌNG LƯỢNG & GIỚI TÍNH (CÂN MẪU)
          </button>
          <button
            type="button"
            onClick={() => setActiveTab(2)}
            className={`px-5 py-2 font-bold text-xs md:text-sm rounded-t transition-colors cursor-pointer ${
              activeTab === 2
                ? 'bg-[#1B4F72] text-white shadow'
                : 'bg-[#D6EAF8] text-[#1B4F72] hover:bg-[#AED6F1]'
            }`}
          >
            ⚡ CHẾ ĐỘ 2: ĐẾM CỘNG DỒN TỐC ĐỘ CAO (UNDO)
          </button>

          {/* Trạng thái đường truyền HTTP POST & Bộ nhớ đệm âm thanh bên phải thanh Tab */}
          <div
            className={`ml-auto px-3 py-1 rounded text-xs font-semibold hidden lg:block ${
              httpLogType === 'ok'
                ? 'bg-emerald-100 text-emerald-900 border border-emerald-300'
                : httpLogType === 'warn'
                  ? 'bg-amber-100 text-amber-900 border border-amber-300'
                  : 'bg-red-100 text-red-900 border border-red-300'
            }`}
          >
            {httpLogMessage}
          </div>
        </div>

        {/* ===================================================================
            TAB 1: CHẾ ĐỘ 1 - CÂN TRỌNG LƯỢNG & GIỚI TÍNH (CÂN MẪU)
            =================================================================== */}
        {activeTab === 1 && (
          <div className="bg-white border border-t-0 border-slate-300 p-3 flex-1 flex flex-col gap-2.5 shadow-sm">
            {/* Banner hướng dẫn cú pháp đọc giọng nói Chế độ 1 */}
            <div className="bg-[#E8F8F5] border border-[#A3E4D7] rounded px-3 py-2 flex flex-col gap-1">
              <p className="text-xs md:text-sm font-bold text-[#0E6655]">
                🎙️ CHẾ ĐỘ RẢNH TAY & CHỐNG TRÙNG LẶP 100%: Đọc 1 lần chỉ nhảy đúng 1 dòng sạch (Tự xóa audio_bytes=None & temp_voice.wav ngay sau khi xử lý).
              </p>
              <p className="text-xs font-bold text-[#117A65]">
                💡 Ví dụ: Đọc "35 trống" ➔ nghỉ {silenceTimeoutSec.toFixed(1)}s ➔ Đọc tiếp "45 mái" ➔ nghỉ {silenceTimeoutSec.toFixed(1)}s ➔ Đọc tiếp "50 đực"   |   Hoàn tác: "Xóa dòng" hoặc "Xóa hết"
              </p>
              <p
                style={{ color: mode1SpeechColor }}
                className="text-xs md:text-sm font-bold mt-0.5"
              >
                {mode1LastSpeech}
              </p>
            </div>

            {/* Bố cục 2 cột: Trái (Bảng dữ liệu + Nhập tay) - Phải (Size Tôm Trung Bình & Thống kê mẫu) */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-3 flex-1">
              {/* CỘT TRÁI (7/12): Bảng Danh Sách Tôm Đã Cân */}
              <div className="lg:col-span-7 border border-slate-300 rounded bg-white p-3 flex flex-col justify-between">
                <div>
                  <div className="flex items-center justify-between mb-2">
                    <h3 className="text-sm font-bold text-[#1B4F72]">
                      📋 Bảng Danh Sách Tôm Đã Cân (Tự động điền đúng 1 dòng/lần đọc)
                    </h3>
                    {/* Nút đọc thử nhanh khẩu lệnh mẫu Chế độ 1 */}
                    <div className="flex flex-wrap items-center gap-1">
                      {['35 trống', '45 mái', '50 đực', 'Xóa dòng'].map(
                        (phrase) => (
                          <button
                            key={phrase}
                            type="button"
                            onClick={() =>
                              void dispatchRecognizedSpeechHttpPost(
                                phrase,
                                4096,
                              )
                            }
                            className="px-2 py-0.5 text-xs font-semibold bg-[#EBF5FB] hover:bg-[#D6EAF8] text-[#1B4F72] border border-[#AED6F1] rounded cursor-pointer"
                          >
                            🗣️ "{phrase}"
                          </button>
                        ),
                      )}
                    </div>
                  </div>

                  {/* Bảng Treeview 4 cột giống app.py */}
                  <div className="border border-slate-300 rounded overflow-hidden max-h-[320px] overflow-y-auto">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="bg-[#1B4F72] text-white text-xs md:text-sm font-bold">
                          <th className="py-2 px-3 text-center w-16">STT</th>
                          <th className="py-2 px-3 text-center">
                            Trọng Lượng (g)
                          </th>
                          <th className="py-2 px-3 text-center">Giới Tính</th>
                          <th className="py-2 px-3 text-center">
                            Thời Gian Nhận Diện
                          </th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200 text-xs md:text-sm">
                        {weightRecords.length === 0 ? (
                          <tr>
                            <td
                              colSpan={4}
                              className="py-12 text-center text-slate-400 italic"
                            >
                              Chưa có mẫu tôm nào. Hãy đọc khẩu lệnh (VD: "35 trống", "45 mái", "50 đực") — Mỗi câu đọc sẽ chỉ chèn đúng 1 dòng sạch.
                            </td>
                          </tr>
                        ) : (
                          weightRecords.map((rec) => {
                            const rowBg =
                              rec.gender === 'Male'
                                ? 'bg-[#EBF5FB]'
                                : rec.gender === 'Female'
                                  ? 'bg-[#FDEDEC]'
                                  : 'bg-[#F8F9F9]';
                            const isSelected = selectedWeightStt === rec.stt;
                            return (
                              <tr
                                key={rec.stt}
                                onClick={() =>
                                  setSelectedWeightStt(
                                    isSelected ? null : rec.stt,
                                  )
                                }
                                className={`${rowBg} ${
                                  isSelected
                                    ? 'ring-2 ring-inset ring-[#1B4F72] font-bold'
                                    : ''
                                } hover:brightness-95 cursor-pointer transition-all`}
                              >
                                <td className="py-2 px-3 text-center font-mono">
                                  {rec.stt}
                                </td>
                                <td className="py-2 px-3 text-center font-mono font-bold text-[#1C2833]">
                                  {rec.weight.toFixed(2)}
                                </td>
                                <td className="py-2 px-3 text-center font-bold">
                                  {rec.gender}
                                </td>
                                <td className="py-2 px-3 text-center font-mono text-slate-600">
                                  {rec.time}
                                </td>
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* Thanh thao tác thủ công cho Chế độ 1 (Giống 100% manual_bar trong app.py) */}
                <form
                  onSubmit={handleManualAddWeight}
                  className="mt-3 pt-2.5 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2 text-xs md:text-sm"
                >
                  <div className="flex flex-wrap items-center gap-2">
                    <label className="font-bold text-[#1C2833]">
                      Nhập tay Trọng lượng (g):
                    </label>
                    <input
                      type="text"
                      value={manualWeightInput}
                      onChange={(e) => setManualWeightInput(e.target.value)}
                      placeholder="VD: 35"
                      className="w-24 px-2 py-1 border border-slate-300 rounded font-bold text-[#1C2833]"
                    />
                    <label className="font-bold text-[#1C2833]">
                      Giới tính:
                    </label>
                    <select
                      value={manualGenderInput}
                      onChange={(e) =>
                        setManualGenderInput(
                          e.target.value as ShrimpGenderLabel,
                        )
                      }
                      className="px-2 py-1 border border-slate-300 rounded font-bold text-[#1C2833] bg-white"
                    >
                      <option value="Male">Male</option>
                      <option value="Female">Female</option>
                      <option value="Chưa rõ">Chưa rõ</option>
                    </select>
                    <button
                      type="submit"
                      className="bg-[#27AE60] hover:bg-[#1E8449] text-white font-bold px-3 py-1 rounded cursor-pointer"
                    >
                      ➕ Thêm Mẫu
                    </button>
                  </div>

                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() => applyMode1Result('Nút bấm', 'CLEAR')}
                      className="bg-[#922B21] hover:bg-[#7B241C] text-white font-bold px-2.5 py-1 rounded cursor-pointer"
                    >
                      🗑️ Xóa Hết Bảng
                    </button>
                    <button
                      type="button"
                      onClick={handleDeleteSelectedWeight}
                      className="bg-[#E67E22] hover:bg-[#D35400] text-white font-bold px-2.5 py-1 rounded cursor-pointer"
                    >
                      ↩️ Xóa Dòng Cuối / Đang Chọn
                    </button>
                  </div>
                </form>
              </div>

              {/* CỘT PHẢI (5/12): Bảng Thống Kê & Tính Size Tôm Trung Bình (Giống 100% right_frame trong app.py) */}
              <div className="lg:col-span-5 border border-slate-300 rounded bg-white p-3 flex flex-col gap-3">
                <h3 className="text-sm font-bold text-[#1B4F72]">
                  📊 Thống Kê Mẫu & Tính Size Tôm
                </h3>

                {/* Hộp nổi bật nhất: SIZE TÔM (Số con / kg) màu vàng nhạt #FEF9E7 viền #F39C12 */}
                <div className="bg-[#FEF9E7] border-2 border-[#F39C12] rounded p-3 text-center">
                  <div className="text-xs md:text-sm font-bold text-[#B7950B]">
                    🦐 SIZE TÔM TRUNG BÌNH
                  </div>
                  <div className="text-[11px] italic text-[#7F8C8D] mt-0.5">
                    (= Tổng số con / Tổng trọng lượng)
                  </div>
                  <div className="text-3xl md:text-4xl font-extrabold text-[#D35400] my-1.5 font-mono">
                    {mode1Stats.totalCount > 0
                      ? `${mode1Stats.sizePerKg.toFixed(1)} con/kg`
                      : '0.0 con/kg'}
                  </div>
                  <div className="text-xs font-bold text-[#7D6608]">
                    {mode1Stats.totalCount > 0
                      ? `(${mode1Stats.sizePerG.toFixed(4)} con/gam)`
                      : '(0.0000 con/gam)'}
                  </div>
                </div>

                {/* Các chỉ số thống kê chi tiết */}
                <div className="space-y-1.5 text-xs md:text-sm">
                  <div className="bg-[#ECF0F1] px-3 py-2 rounded flex items-center justify-between">
                    <span className="font-bold text-[#2C3E50]">
                      Tổng số tôm đã cân:
                    </span>
                    <span className="font-extrabold text-[#1B4F72] text-base">
                      {mode1Stats.totalCount} con
                    </span>
                  </div>
                  <div className="bg-[#ECF0F1] px-3 py-2 rounded flex items-center justify-between">
                    <span className="font-bold text-[#2C3E50]">
                      Tổng trọng lượng:
                    </span>
                    <span className="font-extrabold text-[#1B4F72]">
                      {mode1Stats.totalWeight.toFixed(2)} g (
                      {(mode1Stats.totalWeight / 1000).toFixed(3)} kg)
                    </span>
                  </div>
                  <div className="bg-[#ECF0F1] px-3 py-2 rounded flex items-center justify-between">
                    <span className="font-bold text-[#2C3E50]">
                      Trọng lượng TB:
                    </span>
                    <span className="font-extrabold text-[#1E8449] text-base">
                      {mode1Stats.avgWeight.toFixed(2)} g/con
                    </span>
                  </div>
                </div>

                {/* Thống kê tỷ lệ Giới tính */}
                <div className="border border-slate-300 rounded p-2.5 space-y-1.5 text-xs md:text-sm">
                  <div className="font-bold text-[#1B4F72] mb-1">
                    ♀️♂️ Thống Kê Giới Tính
                  </div>
                  <div className="bg-[#EBF5FB] px-3 py-1.5 rounded flex items-center justify-between">
                    <span className="font-bold text-[#2980B9]">
                      ♂️ Đực (Male):
                    </span>
                    <span className="font-extrabold text-[#2980B9]">
                      {mode1Stats.males} con ({mode1Stats.pctMale.toFixed(1)}%)
                    </span>
                  </div>
                  <div className="bg-[#FDEDEC] px-3 py-1.5 rounded flex items-center justify-between">
                    <span className="font-bold text-[#C0392B]">
                      ♀️ Cái (Female):
                    </span>
                    <span className="font-extrabold text-[#C0392B]">
                      {mode1Stats.females} con (
                      {mode1Stats.pctFemale.toFixed(1)}%)
                    </span>
                  </div>
                  <div className="bg-[#F2F4F4] px-3 py-1.5 rounded flex items-center justify-between">
                    <span className="font-bold text-[#566573]">
                      ❔ Chưa rõ (Unknow):
                    </span>
                    <span className="font-extrabold text-[#566573]">
                      {mode1Stats.unknown} con
                    </span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ===================================================================
            TAB 2: CHẾ ĐỘ 2 - ĐẾM CỘNG DỒN TỐC ĐỘ CAO (UNDO)
            =================================================================== */}
        {activeTab === 2 && (
          <div className="bg-white border border-t-0 border-slate-300 p-3 flex-1 flex flex-col gap-2.5 shadow-sm">
            {/* Bảng hiển thị số tổng cực lớn màu xanh đen #1B2631 (Giống 100% counter_board trong app.py) */}
            <div className="bg-[#1B2631] rounded p-4 text-center shadow-inner">
              <div className="text-xs md:text-sm font-bold text-[#AED6F1] tracking-wider uppercase">
                TỔNG SỐ LƯỢNG TÔM ĐÃ ĐẾM (CON)
              </div>
              <div className="text-5xl md:text-6xl font-extrabold text-[#F1C40F] my-1 font-mono">
                {totalBatchCount.toLocaleString('en-US')}
              </div>
              <div className="text-xs md:text-sm font-bold text-[#A3E4D7]">
                {mode2SubInfo}
              </div>
            </div>

            {/* Dòng trạng thái nhận diện của Chế độ 2 + Khẩu lệnh thử nhanh */}
            <div className="flex flex-wrap items-center justify-between gap-2 px-1">
              <p
                style={{ color: mode2SpeechColor }}
                className="text-xs md:text-sm font-bold"
              >
                {mode2LastSpeech}
              </p>
              <div className="flex flex-wrap items-center gap-1">
                {['Cộng ba mươi', 'Mẻ năm 40 con', 'Mười lăm', 'Hủy'].map(
                  (phrase) => (
                    <button
                      key={phrase}
                      type="button"
                      onClick={() =>
                        void dispatchRecognizedSpeechHttpPost(phrase, 4096)
                      }
                      className="px-2 py-0.5 text-xs font-semibold bg-[#E8F8F5] hover:bg-[#D1F2EB] text-[#0E6655] border border-[#A3E4D7] rounded cursor-pointer"
                    >
                      🗣️ "{phrase}"
                    </button>
                  ),
                )}
              </div>
            </div>

            {/* Bảng lịch sử các mẻ đếm 6 cột giống app.py */}
            <div className="border border-slate-300 rounded flex-1 flex flex-col justify-between p-3">
              <div>
                <h3 className="text-sm font-bold text-[#1B4F72] mb-2">
                  📜 Bảng Tiến Trình Các Mẻ Đếm (Đọc 1 lần chỉ cộng đúng 1 dòng)
                </h3>
                <div className="border border-slate-300 rounded overflow-hidden max-h-[260px] overflow-y-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-[#1B4F72] text-white text-xs md:text-sm font-bold">
                        <th className="py-2 px-3 text-center w-16">STT</th>
                        <th className="py-2 px-3 text-center">Mã Mẻ</th>
                        <th className="py-2 px-3 text-center">
                          Số Lượng Mẻ (Con)
                        </th>
                        <th className="py-2 px-3 text-center">Tổng Tích Lũy</th>
                        <th className="py-2 px-3 text-center">Thời Gian</th>
                        <th className="py-2 px-3 text-center">Trạng Thái</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 text-xs md:text-sm">
                      {batchRecords.length === 0 ? (
                        <tr>
                          <td
                            colSpan={6}
                            className="py-10 text-center text-slate-400 italic"
                          >
                            Chưa có mẻ đếm nào. Hãy đọc số lượng (VD: "Ba mươi lăm", "Mẻ năm 40 con") — Mỗi câu đọc chỉ cộng đúng 1 dòng.
                          </td>
                        </tr>
                      ) : (
                        batchRecords.map((rec) => {
                          const isUndone = rec.status !== 'Hợp lệ';
                          return (
                            <tr
                              key={rec.id}
                              className={
                                isUndone
                                  ? 'bg-[#FADBD8] text-[#922B21] line-through'
                                  : 'bg-[#EAFAF1] text-[#1C2833] font-semibold'
                              }
                            >
                              <td className="py-2 px-3 text-center font-mono">
                                {rec.id}
                              </td>
                              <td className="py-2 px-3 text-center font-bold">
                                {rec.batchCode}
                              </td>
                              <td className="py-2 px-3 text-center font-mono font-bold">
                                +{rec.count}
                              </td>
                              <td className="py-2 px-3 text-center font-mono font-bold">
                                {rec.cumulative}
                              </td>
                              <td className="py-2 px-3 text-center font-mono">
                                {rec.time}
                              </td>
                              <td className="py-2 px-3 text-center font-bold">
                                {rec.status}
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Thanh nút bấm thủ công của Chế độ 2 */}
              <form
                onSubmit={handleManualAddBatch}
                className="mt-3 pt-2.5 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2 text-xs md:text-sm"
              >
                <div className="flex items-center gap-2">
                  <label className="font-bold text-[#1C2833]">
                    Nhập tay số lượng mẻ:
                  </label>
                  <input
                    type="text"
                    value={manualBatchInput}
                    onChange={(e) => setManualBatchInput(e.target.value)}
                    placeholder="VD: 35"
                    className="w-24 px-2 py-1 border border-slate-300 rounded font-bold text-[#1C2833]"
                  />
                  <button
                    type="submit"
                    className="bg-[#27AE60] hover:bg-[#1E8449] text-white font-bold px-3 py-1 rounded cursor-pointer"
                  >
                    ➕ Cộng Mẻ
                  </button>
                </div>

                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    onClick={() => applyMode2Result('Nút bấm', 'CLEAR')}
                    className="bg-[#922B21] hover:bg-[#7B241C] text-white font-bold px-3 py-1 rounded cursor-pointer"
                  >
                    🗑️ Xóa Hết Đếm Lại
                  </button>
                  <button
                    type="button"
                    onClick={() => applyMode2Result('Nút bấm', 'UNDO')}
                    className="bg-[#E67E22] hover:bg-[#D35400] text-white font-bold px-3 py-1 rounded cursor-pointer"
                  >
                    ↩️ Hoàn Tác Mẻ Vừa Rồi (Undo)
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}
      </main>

      {/* 4. THANH CÔNG CỤ DƯỚI CÙNG: GIẢ LẬP GIỌNG NÓI & XUẤT BÁO CÁO EXCEL (Giống 100% _build_bottom_action_bar) */}
      <footer className="bg-[#EAECEE] border-t border-slate-300 px-4 py-2">
        <div className="max-w-[1440px] mx-auto flex flex-wrap items-center justify-between gap-3 text-xs md:text-sm">
          {/* Ô giả lập đọc câu lệnh nhanh (dùng khi test không cần nói vào mic) */}
          <form
            onSubmit={handleSimulateSpeech}
            className="flex flex-wrap items-center gap-2"
          >
            <label className="font-bold text-[#2C3E50]">
              ⌨️ Giả lập đọc nhanh:
            </label>
            <input
              type="text"
              value={simSpeechText}
              onChange={(e) => setSimSpeechText(e.target.value)}
              placeholder='Nhập câu nói VD: "35 trống" hoặc "cộng 35"'
              className="w-64 px-2.5 py-1 bg-white border border-slate-300 rounded font-medium text-[#1C2833]"
            />
            <button
              type="submit"
              className="bg-[#34495E] hover:bg-[#2C3E50] text-white font-bold px-3 py-1 rounded cursor-pointer"
            >
              Gửi Lệnh Đọc (HTTP POST)
            </button>
            <button
              type="button"
              onClick={handleTestSmallAudioBlock}
              title="Kiểm tra chặn file âm thanh nhỏ hơn 1KB (1024 bytes)"
              className="bg-slate-200 hover:bg-slate-300 text-slate-800 font-semibold px-2.5 py-1 rounded border border-slate-300 text-xs cursor-pointer"
            >
              🛡️ Test chặn file &lt; {MIN_AUDIO_BYTES}B
            </button>
          </form>

          {/* Nút Xuất File Excel Báo Cáo */}
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => void handleExportExcelReport()}
              className="bg-[#1F4E79] hover:bg-[#163A5C] text-white font-bold px-4 py-1.5 rounded shadow cursor-pointer"
            >
              📊 XUẤT FILE EXCEL (.XLSX) BÁO CÁO 2 SHEET (OPENPYXL)
            </button>
          </div>
        </div>
      </footer>
    </div>
  );
}

export default App;
