/**
 * Lõi xử lý giọng nói tiếng Việt gốc từ app.py & vietnamese_parser.py (Google STT vi-VN)
 * Tuyệt đối giữ nguyên quy tắc nhận diện từ khóa, bảng ánh xạ số, lọc câu phiếm ngoài ao,
 * và công thức tính Size tôm trung bình (= Tổng số con / Tổng trọng lượng * 1000).
 */

export type ShrimpGenderLabel = 'Male' | 'Female' | 'Chưa rõ';

export interface WeightRecord {
  stt: number;
  weight: number;
  gender: ShrimpGenderLabel;
  time: string;
}

export interface BatchRecord {
  stt: number;
  batch_name: string;
  count: number;
  cumulative: number;
  time: string;
  action: 'ADD';
}

export interface Mode1Stats {
  totalCount: number;
  totalWeight: number;
  avgWeight: number;
  sizePerG: number;
  sizePerKg: number;
  males: number;
  females: number;
  unknown: number;
  pctMale: number;
  pctFemale: number;
}

export interface AnalyzeVoiceApiResponse {
  status: 'success' | 'ignored' | 'error';
  message: string;
  recognized_text?: string;
  command_type?:
    | 'RESET_ALL'
    | 'DELETE_LAST_ROW'
    | 'MODE1_ADD_WEIGHT'
    | 'MODE2_ADD_BATCH'
    | 'MODE2_UNDO'
    | 'IGNORED_CHATTER'
    | 'UNRECOGNIZED';
  weight?: number | null;
  gender?: ShrimpGenderLabel;
  batch_qty?: number | null;
  tts_text?: string;
}

export const MIN_AUDIO_BYTES = 1024; // 1KB minimum threshold

// =============================================================================
// BẢNG ÁNH XẠ DỊCH TỪ CHỮ SANG SỐ TIẾNG VIỆT TỰ ĐỘNG CHO CHẾ ĐỘ 2 (GỐC TỪ APP.PY)
// =============================================================================
export const VIETNAMESE_WORD_TO_NUMBER: Record<string, number> = {
  // 0 - 9
  'không': 0, khong: 0, zero: 0,
  'một': 1, mot: 1, 'mốt': 1,
  hai: 2,
  ba: 3,
  'bốn': 4, bon: 4, 'tư': 4, tu: 4,
  'năm': 5, nam: 5, 'lăm': 5, lam: 5, 'nhăm': 5,
  'sáu': 6, sau: 6,
  'bảy': 7, bay: 7, 'bẩy': 7,
  'tám': 8, tam: 8,
  'chín': 9, chin: 9,

  // 10 - 19
  'mười': 10, muoi: 10, 'mươi': 10,
  'mười một': 11, 'mười hai': 12, 'mười ba': 13,
  'mười bốn': 14, 'mười tư': 14,
  'mười lăm': 15, 'mười năm': 15, 'mười sáu': 16,
  'mười bảy': 17, 'mười bẩy': 17, 'mười tám': 18, 'mười chín': 19,

  // 20 - 29
  'hai mươi': 20, 'hai chục': 20,
  'hai mốt': 21, 'hai một': 21, 'hai hai': 22, 'hai ba': 23,
  'hai tư': 24, 'hai bốn': 24,
  'hai lăm': 25, 'hai năm': 25, 'hai sáu': 26,
  'hai bảy': 27, 'hai bẩy': 27, 'hai tám': 28, 'hai chín': 29,

  // 30 - 39
  'ba mươi': 30, 'ba chục': 30,
  'ba mốt': 31, 'ba hai': 32, 'ba ba': 33,
  'ba tư': 34, 'ba bốn': 34,
  'ba lăm': 35, 'ba năm': 35, 'ba sáu': 36,
  'ba bảy': 37, 'ba bẩy': 37, 'ba tám': 38, 'ba chín': 39,

  // 40 - 49
  'bốn mươi': 40, 'bốn chục': 40,
  'bốn mốt': 41, 'bốn hai': 42, 'bốn ba': 43,
  'bốn tư': 44, 'bốn bốn': 44,
  'bốn lăm': 45, 'bốn năm': 45, 'bốn sáu': 46,
  'bốn bảy': 47, 'bốn bẩy': 47, 'bốn tám': 48, 'bốn chín': 49,

  // 50 - 59
  'năm mươi': 50, 'năm chục': 50,
  'năm mốt': 51, 'năm hai': 52, 'năm ba': 53,
  'năm tư': 54, 'năm bốn': 54,
  'năm lăm': 55, 'năm năm': 55, 'năm sáu': 56,
  'năm bảy': 57, 'năm bẩy': 57, 'năm tám': 58, 'năm chín': 59,

  // 60 - 100
  'sáu mươi': 60, 'sáu chục': 60, 'sáu lăm': 65,
  'bảy mươi': 70, 'bảy chục': 70, 'bảy lăm': 75,
  'tám mươi': 80, 'tám chục': 80, 'tám lăm': 85,
  'chín mươi': 90, 'chín chục': 90, 'chín lăm': 95,
  'một trăm': 100, 'trăm': 100,
};

const MAPPING_1_TO_10: Record<string, number> = {
  'một': 1, 'mốt': 1, mot: 1, '1': 1,
  hai: 2, '2': 2,
  ba: 3, '3': 3,
  'bốn': 4, bon: 4, 'tư': 4, tu: 4, '4': 4,
  'năm': 5, 'lăm': 5, nam: 5, lam: 5, 'nhăm': 5, '5': 5,
  'sáu': 6, sau: 6, '6': 6,
  'bảy': 7, 'bẩy': 7, bay: 7, '7': 7,
  'tám': 8, tam: 8, '8': 8,
  'chín': 9, chin: 9, '9': 9,
  'mười': 10, 'chục': 10, muoi: 10, 'mươi': 10, chuc: 10, '10': 10,
};

/**
 * Kiểm tra dung lượng file âm thanh trước khi gửi (chặn nếu rỗng hoặc < 1KB = 1024 bytes).
 */
export function validateAudioByteLength(byteLength: number): {
  valid: boolean;
  message?: string;
} {
  if (!byteLength || byteLength < MIN_AUDIO_BYTES) {
    return {
      valid: false,
      message: `File âm thanh rỗng hoặc dung lượng quá nhỏ (${byteLength || 0} bytes < 1KB). Hệ thống đã chặn không gửi lên API để tránh lỗi máy chủ.`,
    };
  }
  return { valid: true };
}

/**
 * Nhận diện giới tính tôm tại Chế độ 1 (Cân mẫu) - giữ nguyên từ app.py
 */
export function detect_shrimp_gender_mode1(text: string): ShrimpGenderLabel {
  if (!text) return 'Chưa rõ';

  let t = text
    .toLowerCase()
    .trim()
    .replace(/(\d+)([a-zA-Zà-ỹÀ-Ỹ]+)/g, '$1 $2')
    .replace(/([a-zA-Zà-ỹÀ-Ỹ]+)(\d+)/g, '$1 $2');

  const tokens = t.split(/[\s,.;:!?-]+/).filter(Boolean);

  // 1. KIỂM TRA GIỐNG CÁI (FEMALE): "cái", "gái", "ái", "mái", "female"...
  const femaleKeywords = new Set([
    'cái',
    'cai',
    'gái',
    'gai',
    'ái',
    'female',
    'phimeo',
    'pheme',
    'femail',
    'phi',
    'mái',
    'mai',
  ]);
  if (tokens.some((w) => femaleKeywords.has(w)) || t.includes('phi meo')) {
    return 'Female';
  }

  // 2. KIỂM TRA GIỐNG ĐỰC (MALE): "trai", "chai", "đực", "tai", "trống", "male"...
  const maleKeywords = new Set([
    'trai',
    'chai',
    'đực',
    'duc',
    'tai',
    'male',
    'mail',
    'meo',
    'mê',
    'mel',
    'trống',
    'trong',
  ]);
  if (tokens.some((w) => maleKeywords.has(w))) {
    return 'Male';
  }

  return 'Chưa rõ';
}

/**
 * MỤC 2: LOGIC LỌC CÂU NÓI CHUYỆN PHIẾM NGOÀI AO - giữ nguyên từ app.py
 */
export function is_relevant_shrimp_command(text: string): boolean {
  if (!text) return false;
  const t = text.toLowerCase().trim();

  if (['xóa', 'xoa', 'hủy', 'huy', 'undo', 'reset'].some((k) => t.includes(k))) {
    return true;
  }

  if (/\d/.test(t)) {
    return true;
  }

  const tokens = t.split(/[\s,.;:!?-]+/).filter(Boolean);
  const numWords = new Set([
    'không',
    'khong',
    'một',
    'mot',
    'mốt',
    'hai',
    'ba',
    'bốn',
    'bon',
    'tư',
    'tu',
    'năm',
    'nam',
    'lăm',
    'lam',
    'nhăm',
    'nham',
    'sáu',
    'sau',
    'bảy',
    'bay',
    'bẩy',
    'bey',
    'tám',
    'tam',
    'chín',
    'chin',
    'mười',
    'muoi',
    'mươi',
    'chục',
    'trăm',
    'tram',
  ]);
  if (tokens.some((w) => numWords.has(w))) {
    return true;
  }

  const kwWords = new Set([
    'trai',
    'chai',
    'đực',
    'duc',
    'tai',
    'cái',
    'cai',
    'gái',
    'gai',
    'ái',
    'trống',
    'mái',
    'xóa',
    'xoa',
    'hủy',
    'huy',
    'undo',
    'reset',
    'bỏ',
    'bo',
    'trừ',
    'tru',
    'male',
    'female',
    'mail',
    'meo',
    'mê',
    'mel',
    'pheme',
    'femail',
    'phi',
  ]);
  if (tokens.some((w) => kwWords.has(w)) || t.includes('phi meo')) {
    return true;
  }

  return false;
}

/**
 * Bóc tách số từ câu nói tiếng Việt (dùng cho cả Chế độ 1 và Chế độ 2)
 */
export function parseSpokenNumberValue(text: string): number | null {
  if (!text) return null;
  const cleaned = text
    .toLowerCase()
    .trim()
    .replace(/,/g, '.');

  // Check decimal or integer digits first
  const digitMatch = cleaned.match(/(\d+(?:\.\d+)?)/);
  if (digitMatch) {
    const val = parseFloat(digitMatch[1]);
    if (!Number.isNaN(val) && val > 0) return val;
  }

  const normWords = cleaned.replace(/[^a-zA-Zà-ỹÀ-Ỹ0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
  if (normWords in VIETNAMESE_WORD_TO_NUMBER) {
    const val = VIETNAMESE_WORD_TO_NUMBER[normWords];
    if (val > 0) return val;
  }

  const words = normWords.split(' ').filter(Boolean);
  let total = 0;
  let curr = 0;
  for (const w of words) {
    if (w in VIETNAMESE_WORD_TO_NUMBER) {
      const n = VIETNAMESE_WORD_TO_NUMBER[w];
      if (n === 100) {
        total += (curr > 0 ? curr : 1) * 100;
        curr = 0;
      } else if (n === 10 && curr > 0) {
        curr = curr * 10;
      } else if (curr > 0 && curr % 10 === 0) {
        curr += n;
      } else if (curr === 0) {
        curr = n;
      } else {
        curr = curr * 10 + n;
      }
    }
  }
  total += curr;
  return total > 0 ? total : null;
}

/**
 * Hàm parse_mode1_weighing_and_gender tương đương vietnamese_parser.py
 */
export function parse_mode1_weighing_and_gender(text: string): {
  weight: number | null;
  gender: ShrimpGenderLabel;
} {
  const gender = detect_shrimp_gender_mode1(text);
  const tClean = text
    .toLowerCase()
    .trim()
    .replace(/(\d+)([a-zA-Zà-ỹÀ-Ỹ]+)/g, '$1 $2')
    .replace(/([a-zA-Zà-ỹÀ-Ỹ]+)(\d+)/g, '$1 $2');

  const stripped = tClean
    .split(/\s+/)
    .filter(
      (w) =>
        ![
          'cái',
          'cai',
          'gái',
          'gai',
          'ái',
          'female',
          'phimeo',
          'pheme',
          'femail',
          'phi',
          'mái',
          'mai',
          'trai',
          'chai',
          'đực',
          'duc',
          'tai',
          'male',
          'mail',
          'meo',
          'mê',
          'mel',
          'trống',
          'trong',
          'gram',
          'gam',
          'g',
          'kg',
          'con',
          'tôm',
          'mẫu',
        ].includes(w)
    )
    .join(' ');

  const weight = parseSpokenNumberValue(stripped) ?? parseSpokenNumberValue(text);
  return { weight, gender };
}

/**
 * Áp dụng bảng ánh xạ dịch từ chữ sang số tự động cho Chế độ 2 - giữ nguyên từ app.py
 */
export function map_spoken_text_to_batch(speech_text: string): {
  action: 'ADD' | 'UNDO' | null;
  qty: number | null;
} {
  if (!speech_text) return { action: null, qty: null };

  let text = speech_text.toLowerCase().trim();
  text = text.replace(/[,.](?!\d)/g, ' ');
  text = text.replace(/[\-_/\\:;!?()[\]{}"'~`@#$%^&*+=|]/g, ' ');
  text = text.replace(/\s+/g, ' ').trim();

  const tokens = text.split(' ');
  const undoWords = new Set(['xóa', 'xoa', 'hủy', 'huy', 'undo', 'bỏ', 'bo', 'trừ', 'tru', 'lùi', 'lui', 'nhầm', 'nham']);
  if (tokens.some((w) => undoWords.has(w)) || text.includes('quay lại')) {
    return { action: 'UNDO', qty: null };
  }

  if (text in MAPPING_1_TO_10) {
    return { action: 'ADD', qty: MAPPING_1_TO_10[text] };
  }

  const fillerWords = new Set(['mẻ', 'me', 'thêm', 'them', 'cộng', 'cong', 'con', 'tôm', 'tom', 'bốc', 'boc', 'đợt', 'dot']);
  const cleaned = tokens.filter((w) => !fillerWords.has(w)).join(' ').trim();

  if (cleaned in MAPPING_1_TO_10) {
    return { action: 'ADD', qty: MAPPING_1_TO_10[cleaned] };
  }

  if (cleaned in VIETNAMESE_WORD_TO_NUMBER && VIETNAMESE_WORD_TO_NUMBER[cleaned] > 0) {
    return { action: 'ADD', qty: VIETNAMESE_WORD_TO_NUMBER[cleaned] };
  }

  const numVal = parseSpokenNumberValue(cleaned);
  if (numVal !== null && numVal > 0) {
    return { action: 'ADD', qty: Math.round(numVal) };
  }

  return { action: null, qty: null };
}

/**
 * Xử lý câu lệnh văn bản nhận diện được từ Google STT theo đúng luồng _dispatch_recognized_speech của app.py
 */
export function analyzeRecognizedSpeechText(
  rawText: string,
  activeTabIdx: 0 | 1
): AnalyzeVoiceApiResponse {
  const text = (rawText || '').trim();
  if (!text) {
    return {
      status: 'ignored',
      message: 'Không nghe rõ âm thanh hoặc câu nói trống.',
      recognized_text: '',
      command_type: 'IGNORED_CHATTER',
    };
  }

  // 1. Lọc câu nói chuyện phiếm ngoài ao
  if (!is_relevant_shrimp_command(text)) {
    return {
      status: 'ignored',
      message: `💬 Bỏ qua câu nói chuyện phiếm: "${text}"`,
      recognized_text: text,
      command_type: 'IGNORED_CHATTER',
    };
  }

  const cleanCmd = text
    .toLowerCase()
    .trim()
    .replace(/[.,!?]/g, '')
    .trim();

  // 2. Kiểm tra lệnh XÓA SẠCH TOÀN BỘ (RESET VỀ 0): "xóa hết", "xóa tất cả", "reset"
  const resetSet = new Set([
    'xóa hết',
    'xoa het',
    'xóa tất cả',
    'xoa tat ca',
    'reset',
    'xóa toàn bộ',
    'xoa toan bo',
  ]);
  if (resetSet.has(cleanCmd)) {
    return {
      status: 'success',
      message: '🗑️ Đang xóa toàn bộ dữ liệu: Đã đặt lại bảng hiện tại về 0!',
      recognized_text: text,
      command_type: 'RESET_ALL',
      tts_text: 'Đang xóa toàn bộ dữ liệu',
    };
  }

  // 3. Kiểm tra lệnh XÓA DÒNG CUỐI CÙNG (SỬA SAI RẢNH TAY): "xóa", "hủy", "xóa dòng"
  const undoSet = new Set([
    'xóa',
    'hủy',
    'xóa dòng',
    'xoa',
    'huy',
    'xoa dong',
    'hủy dòng',
    'huy dong',
    'undo',
  ]);
  const hasDigits = /\d/.test(cleanCmd);
  const isUndo =
    undoSet.has(cleanCmd) ||
    ((cleanCmd.includes('xóa dòng') ||
      cleanCmd.includes('xoa dong') ||
      cleanCmd.includes('hủy dòng')) &&
      !hasDigits) ||
    ((cleanCmd.includes('xóa') || cleanCmd.includes('hủy') || cleanCmd.includes('undo')) &&
      !hasDigits &&
      !['hết', 'tất cả', 'toàn bộ', 'tat ca', 'toan bo'].some((k) => cleanCmd.includes(k)));

  if (isUndo) {
    return {
      status: 'success',
      message: '↩️ Đã xóa dòng cuối!',
      recognized_text: text,
      command_type: 'DELETE_LAST_ROW',
      tts_text: 'Đã xóa dòng cuối',
    };
  }

  // 4. Phân phối theo Tab đang mở (0 = Chế độ 1: Cân mẫu, 1 = Chế độ 2: Đếm cộng dồn)
  if (activeTabIdx === 0) {
    const { weight, gender } = parse_mode1_weighing_and_gender(text);
    if (weight === null || weight <= 0) {
      return {
        status: 'ignored',
        message: `❓ Không nhận diện được số cân từ: "${text}". Hãy đọc dạng '35 trai', '35 chai', '45 gái', '50 đực' hoặc '30 cái'.`,
        recognized_text: text,
        command_type: 'UNRECOGNIZED',
      };
    }

    const tts_text =
      gender === 'Male'
        ? `Đã lưu ${weight} gam Male`
        : gender === 'Female'
        ? `Đã lưu ${weight} gam Female`
        : `Đã lưu ${weight} gam`;

    return {
      status: 'success',
      message: `✅ Đã thêm mẫu: ${weight}g - ${gender}`,
      recognized_text: text,
      command_type: 'MODE1_ADD_WEIGHT',
      weight,
      gender,
      tts_text,
    };
  } else {
    const { action, qty } = map_spoken_text_to_batch(cleanCmd);
    if (action === 'UNDO') {
      return {
        status: 'success',
        message: '↩️ Đã xóa dòng cuối!',
        recognized_text: text,
        command_type: 'MODE2_UNDO',
        tts_text: 'Đã xóa dòng cuối',
      };
    }
    if (action === 'ADD' && qty !== null && qty > 0) {
      return {
        status: 'success',
        message: `✅ Đã cộng thêm +${qty} con`,
        recognized_text: text,
        command_type: 'MODE2_ADD_BATCH',
        batch_qty: qty,
      };
    }

    return {
      status: 'ignored',
      message: `❓ Không nhận diện được mẻ từ: "${text}". Chỉ cần đọc: 'năm', 'sáu', 'mười', 'mười lăm' hoặc 'xóa'.`,
      recognized_text: text,
      command_type: 'UNRECOGNIZED',
    };
  }
}

/**
 * MỤC 3: Thống kê tự động Chế độ 1 - giữ nguyên công thức từ app.py:
 * Size tôm trung bình = Tổng số con / Tổng trọng lượng * 1000 (con/kg)
 */
export function calculateMode1Stats(records: WeightRecord[]): Mode1Stats {
  const totalCount = records.length;
  if (totalCount === 0) {
    return {
      totalCount: 0,
      totalWeight: 0,
      avgWeight: 0,
      sizePerG: 0,
      sizePerKg: 0,
      males: 0,
      females: 0,
      unknown: 0,
      pctMale: 0,
      pctFemale: 0,
    };
  }

  const totalWeight = records.reduce((sum, r) => sum + r.weight, 0);
  const avgWeight = totalWeight / totalCount;
  const males = records.filter((r) => r.gender === 'Male').length;
  const females = records.filter((r) => r.gender === 'Female').length;
  const unknown = records.filter((r) => r.gender !== 'Male' && r.gender !== 'Female').length;

  const pctMale = (males / totalCount) * 100;
  const pctFemale = (females / totalCount) * 100;

  const sizePerG = totalWeight > 0 ? totalCount / totalWeight : 0;
  const sizePerKg = sizePerG * 1000;

  return {
    totalCount,
    totalWeight,
    avgWeight,
    sizePerG,
    sizePerKg,
    males,
    females,
    unknown,
    pctMale,
    pctFemale,
  };
}

export const TTS_SPEED_MAP: Record<string, number> = {
  'x1.0 (Bình thường)': 1.0,
  'x1.2 (Khá nhanh)': 1.2,
  'x1.3 (Mặc định)': 1.3,
  'x1.5 (Nhanh)': 1.5,
  'x1.8 (Rất nhanh)': 1.8,
};

export function formatCurrentTime(): string {
  const now = new Date();
  const h = String(now.getHours()).padStart(2, '0');
  const m = String(now.getMinutes()).padStart(2, '0');
  const s = String(now.getSeconds()).padStart(2, '0');
  return `${h}:${m}:${s}`;
}

/**
 * Quản lý trạng thái chống nhập trùng dữ liệu (Anti-Duplicate Input State Machine):
 * 1. Reset Audio State: Xóa bộ nhớ đệm âm thanh (`audio_bytes = null`) và xóa file tạm (`temp_voice.wav`)
 *    ngay sau khi `/api/analyze-voice` xử lý xong.
 * 2. Processing Flag (`is_processed`): Mặc định `false`. Khi có âm thanh mới gửi lên, kiểm tra
 *    `if (!state.is_processed)` mới phân tích và chuyển `is_processed = true`.
 *    Chỉ khi Micro tự động kích hoạt lại sang chu kỳ mới mới đặt `is_processed = false`.
 */
export interface VoiceSessionState {
  audio_bytes: string | null;
  is_processed: boolean;
  temp_file_path: string | null;
  temp_file_removed: boolean;
  cycle_id: number;
}

export function createInitialVoiceSessionState(): VoiceSessionState {
  return {
    audio_bytes: null,
    is_processed: false,
    temp_file_path: null,
    temp_file_removed: true,
    cycle_id: 1,
  };
}

export function receiveNewAudioPayload(
  state: VoiceSessionState,
  audioPayload: string,
  tempFilePath: string = 'temp_voice.wav',
): VoiceSessionState {
  if (state.is_processed) {
    // Đang ở trạng thái đã xử lý của chu kỳ hiện tại -> Không nhận đè lại chính gói âm thanh cũ
    return state;
  }
  return {
    ...state,
    audio_bytes: audioPayload,
    temp_file_path: tempFilePath,
    temp_file_removed: false,
  };
}

export function canProcessCurrentAudio(state: VoiceSessionState): boolean {
  return (
    state.is_processed === false &&
    state.audio_bytes !== null &&
    state.audio_bytes.trim().length > 0
  );
}

export function markAudioProcessedAndResetBuffer(
  state: VoiceSessionState,
): VoiceSessionState {
  return {
    ...state,
    is_processed: true,
    audio_bytes: null, // 1. Xóa bộ nhớ đệm âm thanh ngay sau khi sử dụng
    temp_file_path: null,
    temp_file_removed: true, // Đã xóa temp_voice.wav
  };
}

export function resetFlagOnMicAutoRestart(
  state: VoiceSessionState,
): VoiceSessionState {
  return {
    ...state,
    is_processed: false, // 2. Chuyển is_processed về False khi Micro tự bật lại đón câu mới
    audio_bytes: null,
    temp_file_path: null,
    temp_file_removed: true,
    cycle_id: state.cycle_id + 1,
  };
}


