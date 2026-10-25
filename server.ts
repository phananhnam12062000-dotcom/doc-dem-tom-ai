import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import {
  MIN_AUDIO_BYTES,
  analyzeRecognizedSpeechText,
  validateAudioByteLength,
} from './src/utils/counterLogic.ts';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json({ limit: '25mb' }));

  app.get('/api/health', (_req, res) => {
    res.status(200).json({
      status: 'ok',
      protocol: 'HTTP POST',
      websocket: 'disabled',
      engine: 'Google Speech Recognition (vi-VN) + Original Rule Parser',
      minAudioBytes: MIN_AUDIO_BYTES,
    });
  });

  /**
   * Endpoint HTTP POST tiêu chuẩn `/api/analyze-voice`
   * - Giữ nguyên bộ phân tích giọng nói gốc (Google STT vi-VN + Bộ lọc từ vựng của app.py & audio_manager.py)
   * - Tuyệt đối không gọi Gemini API
   * - Kiểm tra dung lượng file âm thanh >= 1KB (1024 bytes) trước khi xử lý
   * - Bọc toàn bộ trong try ... catch, in `Error details: ...` ra terminal và luôn trả về JSON sạch (HTTP 200)
   */
  app.post('/api/analyze-voice', async (req, res) => {
    try {
      const body = (req.body || {}) as {
        audioBase64?: string;
        audioByteLength?: number;
        mimeType?: string;
        transcriptText?: string;
        activeTabIdx?: 0 | 1;
      };

      const rawTranscript = (body.transcriptText || '').trim();
      const rawAudioBase64 = (body.audioBase64 || '').trim();
      const activeTabIdx: 0 | 1 = body.activeTabIdx === 1 ? 1 : 0;

      let byteLength =
        typeof body.audioByteLength === 'number' ? body.audioByteLength : 0;

      if (rawAudioBase64) {
        const cleanBase64 = rawAudioBase64.includes(',')
          ? rawAudioBase64.split(',')[1]
          : rawAudioBase64;
        const audioBuffer = Buffer.from(cleanBase64, 'base64');
        byteLength = audioBuffer.byteLength;
      }

      // Kiểm tra chất lượng/dung lượng file âm thanh đầu vào (< 1KB / 1024 bytes)
      if (rawAudioBase64 || typeof body.audioByteLength === 'number' || !rawTranscript) {
        const validation = validateAudioByteLength(byteLength);
        if (!validation.valid) {
          console.log(`Error details: ${validation.message}`);
          res.status(200).json({
            status: 'error',
            message: validation.message,
            audioBytes: byteLength,
          });
          return;
        }
      }

      if (!rawTranscript) {
        res.status(200).json({
          status: 'ignored',
          message:
            'File âm thanh hợp lệ (>= 1KB) đã nhận qua HTTP POST nhưng chưa có chuỗi nhận diện giọng nói từ Google STT.',
          audioBytes: byteLength,
        });
        return;
      }

      const analysis = analyzeRecognizedSpeechText(rawTranscript, activeTabIdx);

      res.status(200).json({
        ...analysis,
        protocol: 'HTTP_POST',
        audioBytes: byteLength,
      });
    } catch (e: unknown) {
      const errMessage = e instanceof Error ? e.message : String(e);
      console.error(`Error details: ${errMessage}`);
      res.status(200).json({
        status: 'error',
        message: errMessage,
      });
    }
  });

  if (process.env.NODE_ENV !== 'production') {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: {
        middlewareMode: true,
        hmr: false,
      },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(__dirname, 'dist');
    app.use(express.static(distPath));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://localhost:${PORT} (HTTP POST mode, WebSocket HMR disabled)`);
  });
}

startServer();
