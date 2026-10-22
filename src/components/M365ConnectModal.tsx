'use client';

import React, { useState, useEffect, useRef } from 'react';
import {
  DeviceCodeSession,
  graphService,
  M365AccountProfile,
} from '@/lib/graphService';
import {
  ShieldCheck,
  Copy,
  Check,
  ExternalLink,
  Loader2,
  X,
  RefreshCw,
} from 'lucide-react';

interface M365ConnectModalProps {
  isOpen: boolean;
  onClose: () => void;
  onConnected: (account: M365AccountProfile) => void;
}

export default function M365ConnectModal({
  isOpen,
  onClose,
  onConnected,
}: M365ConnectModalProps) {
  const [session, setSession] = useState<DeviceCodeSession | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [copied, setCopied] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const pollTimerRef = useRef<NodeJS.Timeout | null>(null);

  const stopPolling = () => {
    if (pollTimerRef.current) {
      clearInterval(pollTimerRef.current);
      pollTimerRef.current = null;
    }
  };

  const startDeviceFlow = async () => {
    stopPolling();
    setIsLoading(true);
    setError(null);
    setSession(null);

    try {
      const dc = await graphService.startDeviceCodeLogin();
      setSession(dc);

      // Automatically copy code to clipboard if allowed
      if (typeof navigator !== 'undefined' && navigator.clipboard) {
        try {
          await navigator.clipboard.writeText(dc.userCode);
          setCopied(true);
          setTimeout(() => setCopied(false), 3000);
        } catch {
          // ignore
        }
      }

      // Poll every 3.5 seconds for completion
      pollTimerRef.current = setInterval(async () => {
        try {
          const result = await graphService.pollDeviceCodeLogin(dc.deviceCode);
          if (result.status === 'connected' && result.account) {
            stopPolling();
            onConnected(result.account);
            onClose();
          } else if (result.status === 'error') {
            stopPolling();
            setError(result.error || 'Sign-in code expired. Click Refresh Code below.');
          }
        } catch {
          // ignore transient poll error
        }
      }, 3500);
    } catch (err: any) {
      setError(err?.message || 'Could not start Microsoft 365 sign-in.');
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    if (isOpen) {
      startDeviceFlow();
    } else {
      stopPolling();
    }
    return () => stopPolling();
  }, [isOpen]);

  if (!isOpen) return null;

  const handleCopyAndOpen = async () => {
    if (!session) return;
    if (typeof navigator !== 'undefined' && navigator.clipboard) {
      try {
        await navigator.clipboard.writeText(session.userCode);
        setCopied(true);
        setTimeout(() => setCopied(false), 3000);
      } catch {
        // ignore
      }
    }
    window.open(
      session.verificationUri || 'https://login.microsoft.com/device',
      '_blank',
      'noopener,noreferrer'
    );
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/75 backdrop-blur-sm p-4">
      <div className="w-full max-w-md rounded-2xl border border-slate-800 bg-slate-900 p-6 shadow-2xl">
        {/* Header */}
        <div className="flex items-center justify-between border-b border-slate-800 pb-4">
          <div className="flex items-center gap-2.5">
            <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-sky-500/20 text-sky-400">
              <ShieldCheck className="h-5 w-5" />
            </div>
            <div>
              <h3 className="text-base font-semibold text-white">
                Connect Microsoft 365
              </h3>
              <p className="text-xs text-slate-400">
                One-time Cloud Link (No Redirect URI Required)
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="rounded-lg p-1.5 text-slate-400 hover:bg-slate-800 hover:text-white"
          >
            <X className="h-5 w-5" />
          </button>
        </div>

        {/* Body */}
        <div className="mt-5 space-y-5 text-center">
          {isLoading ? (
            <div className="py-10 flex flex-col items-center gap-3">
              <Loader2 className="h-8 w-8 animate-spin text-sky-400" />
              <p className="text-xs text-slate-400">
                Generating your secure Microsoft 365 activation code...
              </p>
            </div>
          ) : error ? (
            <div className="space-y-4 py-4">
              <div className="rounded-xl bg-rose-500/10 border border-rose-500/30 p-3 text-xs text-rose-300">
                {error}
              </div>
              <button
                type="button"
                onClick={startDeviceFlow}
                className="inline-flex items-center gap-1.5 rounded-xl bg-sky-600 hover:bg-sky-500 px-4 py-2 text-xs font-semibold text-white"
              >
                <RefreshCw className="h-4 w-4" /> Generate New Code
              </button>
            </div>
          ) : session ? (
            <>
              <p className="text-xs text-slate-300 leading-relaxed">
                1. Click the button below to copy this code and open Microsoft&apos;s verification page.<br />
                2. Paste the code and approve sign-in — this window will connect automatically and keep you signed in permanently!
              </p>

              {/* Big Code Box */}
              <div className="flex items-center justify-center gap-3 rounded-2xl border border-sky-500/40 bg-slate-950 py-4 px-6">
                <span className="font-mono text-2xl font-bold tracking-widest text-sky-300 select-all">
                  {session.userCode}
                </span>
                <button
                  type="button"
                  onClick={async () => {
                    await navigator.clipboard.writeText(session.userCode);
                    setCopied(true);
                    setTimeout(() => setCopied(false), 2500);
                  }}
                  className="rounded-lg bg-slate-800 hover:bg-slate-700 p-2 text-slate-200 transition"
                  title="Copy Code"
                >
                  {copied ? (
                    <Check className="h-4 w-4 text-emerald-400" />
                  ) : (
                    <Copy className="h-4 w-4" />
                  )}
                </button>
              </div>

              {copied && (
                <div className="text-[11px] font-medium text-emerald-400">
                  Code copied to clipboard!
                </div>
              )}

              <button
                type="button"
                onClick={handleCopyAndOpen}
                className="w-full inline-flex items-center justify-center gap-2 rounded-xl bg-sky-600 hover:bg-sky-500 px-5 py-3 text-sm font-semibold text-white shadow-lg shadow-sky-600/25 transition"
              >
                Copy Code &amp; Open Microsoft Sign-In
                <ExternalLink className="h-4 w-4" />
              </button>

              <div className="flex items-center justify-center gap-2 pt-2 text-xs text-slate-400">
                <Loader2 className="h-3.5 w-3.5 animate-spin text-sky-400" />
                Waiting for you to complete sign-in in the Microsoft tab...
              </div>
            </>
          ) : null}
        </div>
      </div>
    </div>
  );
}
