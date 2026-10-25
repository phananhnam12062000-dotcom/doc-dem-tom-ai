import React, { useState } from 'react';
import { Check, Copy, Download, FileCode, Terminal } from 'lucide-react';
import {
  AUDIO_MANAGER_PY_CODE,
  EXCEL_MANAGER_PY_CODE,
  STREAMLIT_APP_PY_CODE,
  STREAMLIT_REQUIREMENTS_TXT,
} from '../data/streamlitSource';

export const PythonCodePanel: React.FC = () => {
  const [activeFile, setActiveFile] = useState<
    'app.py' | 'excel_manager.py' | 'audio_manager.py'
  >('app.py');
  const [copiedKey, setCopiedKey] = useState<
    'app.py' | 'excel_manager.py' | 'audio_manager.py' | 'requirements.txt' | 'cmd' | null
  >(null);

  const handleCopy = async (
    content: string,
    key: 'app.py' | 'excel_manager.py' | 'audio_manager.py' | 'requirements.txt' | 'cmd',
  ) => {
    try {
      await navigator.clipboard.writeText(content);
      setCopiedKey(key);
      setTimeout(() => setCopiedKey(null), 2000);
    } catch {
      const textarea = document.createElement('textarea');
      textarea.value = content;
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
      setCopiedKey(key);
      setTimeout(() => setCopiedKey(null), 2000);
    }
  };

  const handleDownloadFile = (filename: string, content: string) => {
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const installCommand =
    'pip install openpyxl SpeechRecognition pyttsx3 gTTS pygame pandas requests PyAudio && python app.py';

  const displayedCode =
    activeFile === 'app.py'
      ? STREAMLIT_APP_PY_CODE
      : activeFile === 'excel_manager.py'
        ? EXCEL_MANAGER_PY_CODE
        : AUDIO_MANAGER_PY_CODE;

  return (
    <div className="space-y-4">
      <section className="bg-white border border-slate-300 rounded-lg p-4 shadow-sm">
        <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3 pb-3 border-b border-slate-200">
          <div>
            <h2 className="text-base font-bold text-[#1B4F72]">
              Mã Nguồn Python (app.py, excel_manager.py & audio_manager.py)
            </h2>
            <p className="text-xs text-slate-600 mt-1">
              Đã tích hợp xuất báo cáo Excel 2 Sheet bằng{' '}
              <code className="font-mono text-slate-800 bg-slate-100 px-1.5 py-0.5 rounded">
                openpyxl
              </code>{' '}
              (Theme Navy Blue{' '}
              <code className="font-mono text-slate-800 bg-slate-100 px-1.5 py-0.5 rounded">
                #1F4E79
              </code>
              , dòng xen kẽ{' '}
              <code className="font-mono text-slate-800 bg-slate-100 px-1.5 py-0.5 rounded">
                #F2F5F8
              </code>
              , công thức động{' '}
              <code className="font-mono text-slate-800 bg-slate-100 px-1.5 py-0.5 rounded">
                SUM / AVERAGE / COUNTIF / IFERROR
              </code>{' '}
              &amp; Auto-fit cột).
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={() =>
                handleDownloadFile('requirements.txt', STREAMLIT_REQUIREMENTS_TXT)
              }
              className="px-3 py-1.5 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded transition-colors flex items-center gap-1.5 cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              requirements.txt
            </button>
            <button
              type="button"
              onClick={() =>
                handleDownloadFile('excel_manager.py', EXCEL_MANAGER_PY_CODE)
              }
              className="px-3 py-1.5 text-xs font-semibold text-white bg-[#1F4E79] hover:bg-[#163A5C] rounded transition-colors flex items-center gap-1.5 cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              Tải excel_manager.py
            </button>
            <button
              type="button"
              onClick={() =>
                handleDownloadFile('audio_manager.py', AUDIO_MANAGER_PY_CODE)
              }
              className="px-3 py-1.5 text-xs font-semibold text-white bg-[#0E6655] hover:bg-[#0B5345] rounded transition-colors flex items-center gap-1.5 cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              Tải audio_manager.py
            </button>
            <button
              type="button"
              onClick={() => handleDownloadFile('app.py', STREAMLIT_APP_PY_CODE)}
              className="px-3 py-1.5 text-xs font-semibold text-white bg-[#1B4F72] hover:bg-[#154360] rounded transition-colors flex items-center gap-1.5 cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              Tải app.py
            </button>
          </div>
        </div>

        <div className="mt-3 bg-slate-900 text-slate-100 rounded p-3 flex items-center justify-between gap-3">
          <div className="flex items-center gap-2 text-xs font-mono text-sky-300 overflow-x-auto">
            <Terminal className="w-4 h-4 text-slate-400 shrink-0" />
            <span>{installCommand}</span>
          </div>
          <button
            type="button"
            onClick={() => handleCopy(installCommand, 'cmd')}
            className="text-xs text-amber-300 hover:text-white flex items-center gap-1 shrink-0 cursor-pointer"
          >
            {copiedKey === 'cmd' ? (
              <>
                <Check className="w-3.5 h-3.5" /> Đã chép
              </>
            ) : (
              <>
                <Copy className="w-3.5 h-3.5" /> Sao chép lệnh
              </>
            )}
          </button>
        </div>
      </section>

      <section className="bg-slate-900 border border-slate-800 rounded-lg overflow-hidden">
        <div className="flex items-center justify-between px-4 py-2.5 bg-slate-950 border-b border-slate-800">
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => setActiveFile('app.py')}
              className={`px-3 py-1 rounded text-xs font-mono font-bold flex items-center gap-1.5 cursor-pointer ${
                activeFile === 'app.py'
                  ? 'bg-[#1B4F72] text-white'
                  : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
              }`}
            >
              <FileCode className="w-3.5 h-3.5" />
              app.py
            </button>
            <button
              type="button"
              onClick={() => setActiveFile('excel_manager.py')}
              className={`px-3 py-1 rounded text-xs font-mono font-bold flex items-center gap-1.5 cursor-pointer ${
                activeFile === 'excel_manager.py'
                  ? 'bg-[#1F4E79] text-white'
                  : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
              }`}
            >
              <FileCode className="w-3.5 h-3.5" />
              excel_manager.py (openpyxl)
            </button>
            <button
              type="button"
              onClick={() => setActiveFile('audio_manager.py')}
              className={`px-3 py-1 rounded text-xs font-mono font-bold flex items-center gap-1.5 cursor-pointer ${
                activeFile === 'audio_manager.py'
                  ? 'bg-[#0E6655] text-white'
                  : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
              }`}
            >
              <FileCode className="w-3.5 h-3.5" />
              audio_manager.py
            </button>
          </div>
          <button
            type="button"
            onClick={() => handleCopy(displayedCode, activeFile)}
            className="px-3 py-1 text-xs font-semibold text-slate-200 bg-slate-800 hover:bg-slate-700 rounded transition-colors flex items-center gap-1.5 cursor-pointer"
          >
            {copiedKey === activeFile ? (
              <>
                <Check className="w-3.5 h-3.5 text-emerald-400" />
                Đã sao chép {activeFile}
              </>
            ) : (
              <>
                <Copy className="w-3.5 h-3.5" />
                Sao chép {activeFile}
              </>
            )}
          </button>
        </div>
        <div className="p-4 overflow-x-auto max-h-[520px] overflow-y-auto">
          <pre className="font-mono text-xs leading-relaxed text-slate-100 select-all">
            <code>{displayedCode}</code>
          </pre>
        </div>
      </section>
    </div>
  );
};
