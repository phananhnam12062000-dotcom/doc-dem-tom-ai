import test from 'node:test';
import assert from 'node:assert/strict';
import {
  analyzeRecognizedSpeechText,
  calculateMode1Stats,
  canProcessCurrentAudio,
  createInitialVoiceSessionState,
  detect_shrimp_gender_mode1,
  is_relevant_shrimp_command,
  map_spoken_text_to_batch,
  markAudioProcessedAndResetBuffer,
  receiveNewAudioPayload,
  resetFlagOnMicAutoRestart,
  validateAudioByteLength,
} from './counterLogic';

test('detect_shrimp_gender_mode1 matches original app.py rules', () => {
  assert.equal(detect_shrimp_gender_mode1('35 trai'), 'Male');
  assert.equal(detect_shrimp_gender_mode1('35chai'), 'Male');
  assert.equal(detect_shrimp_gender_mode1('50 đực'), 'Male');
  assert.equal(detect_shrimp_gender_mode1('38 tai'), 'Male');
  assert.equal(detect_shrimp_gender_mode1('45 gái'), 'Female');
  assert.equal(detect_shrimp_gender_mode1('30 cái'), 'Female');
  assert.equal(detect_shrimp_gender_mode1('25 ái'), 'Female');
  assert.equal(detect_shrimp_gender_mode1('40'), 'Chưa rõ');
});

test('is_relevant_shrimp_command filters out pond chatter and keeps valid commands', () => {
  assert.equal(is_relevant_shrimp_command('hôm nay trời nắng đẹp quá'), false);
  assert.equal(is_relevant_shrimp_command('35 trai'), true);
  assert.equal(is_relevant_shrimp_command('mười lăm'), true);
  assert.equal(is_relevant_shrimp_command('xóa dòng'), true);
  assert.equal(is_relevant_shrimp_command('xóa hết'), true);
});

test('map_spoken_text_to_batch maps Vietnamese words and undo commands in Mode 2', () => {
  assert.deepEqual(map_spoken_text_to_batch('sáu'), { action: 'ADD', qty: 6 });
  assert.deepEqual(map_spoken_text_to_batch('mười lăm'), { action: 'ADD', qty: 15 });
  assert.deepEqual(map_spoken_text_to_batch('ba lăm'), { action: 'ADD', qty: 35 });
  assert.deepEqual(map_spoken_text_to_batch('hủy'), { action: 'UNDO', qty: null });
  assert.deepEqual(map_spoken_text_to_batch('xóa'), { action: 'UNDO', qty: null });
});

test('calculateMode1Stats computes sizePerKg (= total_count / total_weight * 1000)', () => {
  const stats = calculateMode1Stats([
    { stt: 1, weight: 40, gender: 'Male', time: '08:00:00' },
    { stt: 2, weight: 40, gender: 'Female', time: '08:00:05' },
  ]);
  assert.equal(stats.totalCount, 2);
  assert.equal(stats.totalWeight, 80);
  assert.equal(stats.avgWeight, 40);
  assert.equal(stats.sizePerKg, 25.0);
});

test('validateAudioByteLength blocks empty or <1KB audio payloads', () => {
  assert.equal(validateAudioByteLength(0).valid, false);
  assert.equal(validateAudioByteLength(512).valid, false);
  assert.equal(validateAudioByteLength(2048).valid, true);
});

test('analyzeRecognizedSpeechText dispatches Mode 1, Mode 2, Undo, and Reset commands accurately', () => {
  const m1 = analyzeRecognizedSpeechText('45 gái', 0);
  assert.equal(m1.status, 'success');
  assert.equal(m1.command_type, 'MODE1_ADD_WEIGHT');
  assert.equal(m1.weight, 45);
  assert.equal(m1.gender, 'Female');

  const m2 = analyzeRecognizedSpeechText('mười lăm', 1);
  assert.equal(m2.status, 'success');
  assert.equal(m2.command_type, 'MODE2_ADD_BATCH');
  assert.equal(m2.batch_qty, 15);

  const undo = analyzeRecognizedSpeechText('xóa dòng', 0);
  assert.equal(undo.command_type, 'DELETE_LAST_ROW');
});

test('VoiceSessionState prevents duplicate input via Reset Audio State and is_processed flag', () => {
  let state = createInitialVoiceSessionState();
  assert.equal(state.is_processed, false);
  assert.equal(state.audio_bytes, null);
  assert.equal(canProcessCurrentAudio(state), false);

  // 1. New audio arrives ("35 trống")
  state = receiveNewAudioPayload(state, '35 trống', 'temp_voice.wav');
  assert.equal(canProcessCurrentAudio(state), true);
  assert.equal(state.is_processed, false);
  assert.equal(state.temp_file_removed, false);

  // 2. Immediately after /api/analyze-voice processes it -> mark processed & reset audio_bytes = null + remove temp_voice.wav
  state = markAudioProcessedAndResetBuffer(state);
  assert.equal(state.is_processed, true);
  assert.equal(state.audio_bytes, null);
  assert.equal(state.temp_file_removed, true);
  assert.equal(canProcessCurrentAudio(state), false);

  // 3. Even if a duplicate rerun/callback tries to resend before mic restarts, it is blocked!
  const duplicateAttempt = receiveNewAudioPayload(state, '35 trống', 'temp_voice.wav');
  assert.equal(canProcessCurrentAudio(duplicateAttempt), false);

  // 4. When Microphone auto-restarts for next cycle, is_processed resets to false
  state = resetFlagOnMicAutoRestart(state);
  assert.equal(state.is_processed, false);
  assert.equal(state.audio_bytes, null);
  assert.equal(state.cycle_id, 2);
});
