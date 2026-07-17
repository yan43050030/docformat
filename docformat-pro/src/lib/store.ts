import { create } from 'zustand';
import type { Preset, ProcessMode } from '@/lib/formatting/types';
import { BUILTIN_PRESETS } from '@/lib/formatting/presets';

export type ThemeId = 'paper' | 'dark' | 'ink' | 'teal';

interface AppState {
  // Navigation
  activePage: 'home' | 'presets' | 'theme' | 'log';
  setActivePage: (page: 'home' | 'presets' | 'theme' | 'log') => void;

  // Presets
  presets: Preset[];
  activePresetId: string;
  editingPreset: Preset | null;
  setActivePresetId: (id: string) => void;
  setEditingPreset: (preset: Preset | null) => void;
  addPreset: (preset: Preset) => void;
  removePreset: (id: string) => void;
  updatePreset: (preset: Preset) => void;
  duplicatePreset: (id: string) => void;

  // Processing
  processMode: ProcessMode;
  setProcessMode: (mode: ProcessMode) => void;
  selectedFiles: string[];
  setSelectedFiles: (files: string[]) => void;
  isProcessing: boolean;
  setIsProcessing: (v: boolean) => void;
  progress: number;
  setProgress: (v: number) => void;

  // Preview
  showPreview: boolean;
  setShowPreview: (v: boolean) => void;

  // Theme
  theme: ThemeId;
  setTheme: (t: ThemeId) => void;

  // Logs
  logs: LogEntry[];
  addLog: (entry: Omit<LogEntry, 'id' | 'timestamp'>) => void;
  clearLogs: () => void;
}

export interface LogEntry {
  id: string;
  timestamp: number;
  level: 'info' | 'success' | 'warning' | 'error';
  message: string;
  file?: string;
}

export const useAppStore = create<AppState>((set, get) => ({
  activePage: 'home',
  setActivePage: (page) => set({ activePage: page }),

  presets: [...BUILTIN_PRESETS],
  activePresetId: 'official',
  editingPreset: null,
  setActivePresetId: (id) => set({ activePresetId: id }),
  setEditingPreset: (preset) => set({ editingPreset: preset }),
  addPreset: (preset) => set((s) => ({ presets: [...s.presets, preset] })),
  removePreset: (id) => {
    const { presets, activePresetId } = get();
    const filtered = presets.filter((p) => p.id !== id);
    set({
      presets: filtered,
      activePresetId: activePresetId === id ? 'official' : activePresetId,
    });
  },
  updatePreset: (preset) =>
    set((s) => ({
      presets: s.presets.map((p) => (p.id === preset.id ? preset : p)),
    })),
  duplicatePreset: (id) => {
    const { presets } = get();
    const source = presets.find((p) => p.id === id);
    if (!source) return;
    const copy: Preset = JSON.parse(JSON.stringify(source));
    copy.id = `user_${Date.now()}`;
    copy.name = `${source.name} (副本)`;
    copy.scope = 'user';
    set((s) => ({ presets: [...s.presets, copy] }));
  },

  processMode: 'full',
  setProcessMode: (mode) => set({ processMode: mode }),
  selectedFiles: [],
  setSelectedFiles: (files) => set({ selectedFiles: files }),
  isProcessing: false,
  setIsProcessing: (v) => set({ isProcessing: v }),
  progress: 0,
  setProgress: (v) => set({ progress: v }),

  showPreview: true,
  setShowPreview: (v) => set({ showPreview: v }),

  theme: 'paper',
  setTheme: (t) => set({ theme: t }),

  logs: [],
  addLog: (entry) =>
    set((s) => ({
      logs: [
        {
          ...entry,
          id: crypto.randomUUID(),
          timestamp: Date.now(),
        },
        ...s.logs,
      ].slice(0, 200),
    })),
  clearLogs: () => set({ logs: [] }),
}));
