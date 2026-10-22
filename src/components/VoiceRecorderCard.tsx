'use client';

import React, { useState, useEffect, useRef } from 'react';
import {
  Mic,
  Square,
  Sparkles,
  RotateCcw,
  Volume2,
  Wand2,
  MessageSquare,
  CheckSquare,
  LayoutGrid,
  Mail,
  Loader2,
} from 'lucide-react';

interface VoiceRecorderCardProps {
  language: string;
  aiProvider: string;
  isProcessing: boolean;
  onProcessTranscript: (transcript: string, durationSeconds: number) => Promise<void>;
}

const QUICK_EXAMPLES = [
  {
    label: 'Multi-Action Combo',
    text: 'Send a WhatsApp to Roland saying the HVAC quote is ready for review, and create a Planner task to order copper piping by Friday, and add an Outlook task to call the supplier tomorrow morning',
  },
  {
    label: 'WhatsApp Message',
    text: 'Send a WhatsApp to site manager saying delivery truck has arrived at gate 2 please sign the waybill',
  },
  {
    label: 'Planner + Outlook',
    text: 'Create a Planner task to finalize project handover documents urgent due tomorrow, and email the client saying we are on schedule for commissioning',
  },
];

export default function VoiceRecorderCard({
  language,
  aiProvider,
  isProcessing,
  onProcessTranscript,
}: VoiceRecorderCardProps) {
  const [isRecording, setIsRecording] = useState(false);
  const [transcript, setTranscript] = useState('');
  const [interimText, setInterimText] = useState('');
  const [duration, setDuration] = useState(0);
  const [speechSupported, setSpeechSupported] = useState(true);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [audioLevels, setAudioLevels] = useState<number[]>(
    new Array(18).fill(15)
  );

  const recognitionRef = useRef<any>(null);
  const timerRef = useRef<NodeJS.Timeout | null>(null);
  const waveRef = useRef<NodeJS.Timeout | null>(null);
  const isRecordingRef = useRef(false);

  useEffect(() => {
    if (typeof window !== 'undefined') {
      const SpeechRec =
        (window as any).SpeechRecognition ||
        (window as any).webkitSpeechRecognition;
      if (!SpeechRec) {
        setSpeechSupported(false);
      }
    }
    return () => {
      stopAllTimers();
      if (recognitionRef.current) {
        try {
          recognitionRef.current.stop();
        } catch {
          // ignore
        }
      }
    };
  }, []);

  const stopAllTimers = () => {
    if (timerRef.current) clearInterval(timerRef.current);
    if (waveRef.current) clearInterval(waveRef.current);
    setAudioLevels(new Array(18).fill(15));
  };

  const startRecording = () => {
    setErrorMsg(null);
    const SpeechRec =
      typeof window !== 'undefined'
        ? (window as any).SpeechRecognition ||
          (window as any).webkitSpeechRecognition
        : null;

    if (!SpeechRec) {
      setErrorMsg(
        'Browser speech recognition is not supported in this browser. Please use Chrome, Edge, or Safari, or type your command below.'
      );
      return;
    }

    try {
      const recognition = new SpeechRec();
      recognition.continuous = true;
      recognition.interimResults = true;
      recognition.lang = language || 'en-ZA';

      recognition.onresult = (event: any) => {
        let finalChunk = '';
        let interimChunk = '';
        for (let i = event.resultIndex; i < event.results.length; i++) {
          const res = event.results[i];
          if (res.isFinal) {
            finalChunk += res[0].transcript + ' ';
          } else {
            interimChunk += res[0].transcript;
          }
        }
        if (finalChunk) {
          setTranscript((prev) => `${prev} ${finalChunk}`.replace(/\s+/g, ' ').trim());
        }
        setInterimText(interimChunk);
      };

      recognition.onerror = (event: any) => {
        if (event.error === 'no-speech') return;
        console.warn('Speech recognition notice:', event.error);
        if (event.error === 'not-allowed') {
          setErrorMsg(
            'Microphone access was denied. Please allow microphone permissions in your browser address bar.'
          );
          stopRecording();
        }
      };

      recognition.onend = () => {
        // Restart seamlessly if user hasn't clicked stop
        if (isRecordingRef.current) {
          try {
            recognition.start();
          } catch {
            // ignore
          }
        }
      };

      recognition.start();
      recognitionRef.current = recognition;
      isRecordingRef.current = true;
      setIsRecording(true);
      setDuration(0);

      timerRef.current = setInterval(() => {
        setDuration((d) => d + 1);
      }, 1000);

      waveRef.current = setInterval(() => {
        setAudioLevels(
          Array.from({ length: 18 }, () => Math.floor(Math.random() * 75) + 20)
        );
      }, 120);
    } catch (err: any) {
      setErrorMsg(err?.message || 'Could not start microphone.');
    }
  };

  const stopRecording = () => {
    isRecordingRef.current = false;
    setIsRecording(false);
    stopAllTimers();
    if (recognitionRef.current) {
      try {
        recognitionRef.current.stop();
      } catch {
        // ignore
      }
    }
    if (interimText) {
      setTranscript((prev) => `${prev} ${interimText}`.replace(/\s+/g, ' ').trim());
      setInterimText('');
    }
  };

  const handleExtract = async () => {
    if (isRecording) {
      stopRecording();
    }
    const combined = `${transcript} ${interimText}`.replace(/\s+/g, ' ').trim();
    if (!combined) return;
    await onProcessTranscript(combined, duration);
    setTranscript('');
    setInterimText('');
    setDuration(0);
  };

  const formatTimer = (sec: number) => {
    const m = Math.floor(sec / 60);
    const s = sec % 60;
    return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
  };

  const combinedText = `${transcript}${interimText ? ' ' + interimText : ''}`.trim();

  return (
    <div className="relative overflow-hidden rounded-2xl border border-slate-800 bg-gradient-to-br from-slate-900 via-slate-900 to-indigo-950/60 p-6 shadow-xl">
      {/* Top Status Row */}
      <div className="flex flex-wrap items-center justify-between gap-3 mb-5">
        <div className="flex items-center gap-2.5">
          <div
            className={`flex h-9 w-9 items-center justify-center rounded-xl ${
              isRecording
                ? 'bg-rose-500/20 text-rose-400 animate-pulse ring-2 ring-rose-500/40'
                : 'bg-indigo-500/20 text-indigo-400'
            }`}
          >
            <Volume2 className="h-5 w-5" />
          </div>
          <div>
            <h2 className="text-base font-semibold text-white flex items-center gap-2">
              Voice Command Center
              <span className="rounded-full bg-indigo-500/20 border border-indigo-500/30 px-2.5 py-0.5 text-[11px] font-medium text-indigo-300">
                {aiProvider === 'smart-local'
                  ? 'Smart Multi-Intent Engine'
                  : aiProvider === 'gemini'
                  ? 'Gemini AI'
                  : 'OpenAI GPT'}
              </span>
            </h2>
            <p className="text-xs text-slate-400">
              Speak naturally to send WhatsApps, create Outlook tasks, generate Planner tasks, or draft emails in one go.
            </p>
          </div>
        </div>

        {/* Supported Intents Pills */}
        <div className="hidden lg:flex items-center gap-1.5 text-[11px]">
          <span className="inline-flex items-center gap-1 rounded-lg bg-emerald-500/10 border border-emerald-500/20 px-2.5 py-1 text-emerald-300">
            <MessageSquare className="h-3 w-3" /> WhatsApp API
          </span>
          <span className="inline-flex items-center gap-1 rounded-lg bg-sky-500/10 border border-sky-500/20 px-2.5 py-1 text-sky-300">
            <CheckSquare className="h-3 w-3" /> Outlook Tasks
          </span>
          <span className="inline-flex items-center gap-1 rounded-lg bg-purple-500/10 border border-purple-500/20 px-2.5 py-1 text-purple-300">
            <LayoutGrid className="h-3 w-3" /> MS Planner
          </span>
          <span className="inline-flex items-center gap-1 rounded-lg bg-amber-500/10 border border-amber-500/20 px-2.5 py-1 text-amber-300">
            <Mail className="h-3 w-3" /> Outlook Email
          </span>
        </div>
      </div>

      {/* Main Recording Controls & Waveform */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 items-stretch">
        {/* Left: Mic Button & Waveform */}
        <div className="lg:col-span-4 flex flex-col items-center justify-center rounded-xl border border-slate-800/90 bg-slate-950/60 p-5">
          <button
            type="button"
            onClick={isRecording ? stopRecording : startRecording}
            disabled={isProcessing}
            className={`group relative flex h-20 w-20 items-center justify-center rounded-full transition-all duration-300 shadow-lg ${
              isRecording
                ? 'bg-rose-600 hover:bg-rose-500 text-white ring-8 ring-rose-500/20 scale-105'
                : 'bg-indigo-600 hover:bg-indigo-500 text-white hover:scale-105 ring-4 ring-indigo-500/20'
            }`}
          >
            {isRecording ? (
              <Square className="h-8 w-8 fill-current" />
            ) : (
              <Mic className="h-9 w-9" />
            )}
          </button>

          <div className="mt-3 text-center">
            <div className="text-sm font-semibold text-white">
              {isRecording ? (
                <span className="text-rose-400 flex items-center gap-1.5 justify-center">
                  <span className="h-2 w-2 rounded-full bg-rose-500 animate-ping" />
                  Listening... ({formatTimer(duration)})
                </span>
              ) : (
                'Tap Microphone to Speak'
              )}
            </div>
            <div className="text-[11px] text-slate-400 mt-0.5">
              {speechSupported
                ? `Voice Locale: ${language}`
                : 'Type or paste command on the right'}
            </div>
          </div>

          {/* Animated Waveform Bars */}
          <div className="mt-4 flex h-8 items-end justify-center gap-1 w-full px-4">
            {audioLevels.map((lvl, idx) => (
              <div
                key={idx}
                className={`w-1.5 rounded-full transition-all duration-100 ${
                  isRecording ? 'bg-rose-400' : 'bg-slate-800'
                }`}
                style={{ height: `${isRecording ? lvl : 18}%` }}
              />
            ))}
          </div>
        </div>

        {/* Right: Live Transcript Editor & Action Triggers */}
        <div className="lg:col-span-8 flex flex-col justify-between rounded-xl border border-slate-800/90 bg-slate-950/60 p-4">
          <div>
            <div className="flex items-center justify-between mb-2">
              <span className="text-xs font-medium text-slate-400">
                Live Voice Transcript / Command Input
              </span>
              {combinedText && (
                <button
                  type="button"
                  onClick={() => {
                    setTranscript('');
                    setInterimText('');
                  }}
                  className="inline-flex items-center gap-1 text-xs text-slate-400 hover:text-slate-200"
                >
                  <RotateCcw className="h-3 w-3" /> Clear
                </button>
              )}
            </div>

            <textarea
              rows={3}
              value={
                isRecording && interimText
                  ? `${transcript} ${interimText}`.trim()
                  : transcript
              }
              onChange={(e) => setTranscript(e.target.value)}
              placeholder='Try saying: "Send a WhatsApp to John saying the site inspection is tomorrow at 10am, and create a Planner task to prepare safety files by Friday, and add an Outlook task to call Eskom..."'
              className="w-full rounded-lg border border-slate-800 bg-slate-900/90 p-3 text-sm text-slate-100 placeholder-slate-500 focus:border-indigo-500 focus:outline-none focus:ring-1 focus:ring-indigo-500 resize-none"
            />
          </div>

          {errorMsg && (
            <div className="mt-2 rounded-lg bg-rose-500/10 border border-rose-500/30 px-3 py-1.5 text-xs text-rose-300">
              {errorMsg}
            </div>
          )}

          {/* Quick Example Prompts + Extract Button */}
          <div className="mt-3 flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-800/80">
            <div className="flex flex-wrap items-center gap-1.5">
              <span className="text-[11px] text-slate-500 mr-1 flex items-center gap-1">
                <Wand2 className="h-3 w-3" /> Try example:
              </span>
              {QUICK_EXAMPLES.map((ex, i) => (
                <button
                  key={i}
                  type="button"
                  onClick={() => setTranscript(ex.text)}
                  className="rounded-md bg-slate-800/90 hover:bg-slate-700/90 border border-slate-700/70 px-2 py-1 text-[11px] text-slate-300 transition"
                >
                  {ex.label}
                </button>
              ))}
            </div>

            <button
              type="button"
              onClick={handleExtract}
              disabled={!combinedText || isProcessing}
              className="inline-flex items-center gap-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 disabled:opacity-40 disabled:pointer-events-none px-4 py-2 text-xs font-semibold text-white shadow-md shadow-indigo-600/20 transition"
            >
              {isProcessing ? (
                <>
                  <Loader2 className="h-4 w-4 animate-spin" />
                  Extracting Actions...
                </>
              ) : (
                <>
                  <Sparkles className="h-4 w-4" />
                  Turn Voice Note into Actions
                </>
              )}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
