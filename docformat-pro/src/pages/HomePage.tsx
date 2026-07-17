import { useState, useCallback } from 'react';
import { useAppStore } from '@/lib/store';
import { Button } from '@/components/ui/button';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { RadioGroup } from '@/components/ui/radio-group';
import { Select } from '@/components/ui/select';
import { FileDropZone } from '@/components/home/FileDropZone';
import { ProgressBar } from '@/components/home/ProgressBar';
import { FileList } from '@/components/home/FileList';
import type { ProcessMode } from '@/lib/formatting/types';
import { Play, FileOutput, Wand2, Stethoscope, Pilcrow, ClipboardPaste } from 'lucide-react';

const MODE_OPTIONS: { label: string; value: ProcessMode; description: string; icon: typeof Wand2 }[] = [
  { label: '智能一键处理', value: 'full', description: '标点修复 + 排版规范 + 样式清洗，一步到位', icon: Wand2 },
  { label: '格式诊断', value: 'diagnose', description: '仅分析文档问题，不修改文件内容', icon: Stethoscope },
  { label: '标点修复', value: 'punctuation', description: '仅修复中英文标点混用，保留原有段落格式', icon: Pilcrow },
  { label: 'AI 粘贴生成', value: 'ai_paste', description: '粘贴 AI 生成的文本或 Markdown，自动生成规范 docx', icon: ClipboardPaste },
];

export function HomePage() {
  const {
    processMode, setProcessMode,
    presets, activePresetId, setActivePresetId,
    selectedFiles, setSelectedFiles,
    isProcessing, setIsProcessing,
    progress, setProgress,
    addLog,
  } = useAppStore();

  const [outputSuffix, setOutputSuffix] = useState('_processed');

  const handleSelectFiles = useCallback(() => {
    const input = document.createElement('input');
    input.type = 'file';
    input.multiple = true;
    input.accept = '.docx,.doc,.wps';
    input.onchange = () => {
      const files = Array.from(input.files || []);
      const paths = files.map((f) => (f as any).path || f.name);
      setSelectedFiles(paths);
    };
    input.click();
  }, [setSelectedFiles]);

  const handleProcess = useCallback(async () => {
    if (selectedFiles.length === 0) return;
    setIsProcessing(true);
    setProgress(0);

    for (let i = 0; i < selectedFiles.length; i++) {
      const file = selectedFiles[i];
      addLog({ level: 'info', message: `正在处理: ${file}` });

      // Simulate processing — in production this calls the real formatting engine
      await new Promise((r) => setTimeout(r, 600));

      setProgress(Math.round(((i + 1) / selectedFiles.length) * 100));
      addLog({ level: 'success', message: `已完成: ${file}` });
    }

    setIsProcessing(false);
    addLog({ level: 'info', message: `全部处理完成，共 ${selectedFiles.length} 个文件` });
  }, [selectedFiles, addLog, setIsProcessing, setProgress]);

  const removeFile = (index: number) => {
    setSelectedFiles(selectedFiles.filter((_, i) => i !== index));
  };

  const activePreset = presets.find((p) => p.id === activePresetId);
  const presetOptions = presets.map((p) => ({ label: p.name, value: p.id }));

  return (
    <div className="max-w-3xl mx-auto px-6 py-8 space-y-6">
      <div>
        <h1 className="text-xl font-bold text-[var(--color-ink)]">公文格式处理</h1>
        <p className="text-sm text-[var(--color-ink-muted)] mt-1">
          一键规范化 Word 文档排版，遵循 GB/T 9704-2012 标准
        </p>
      </div>

      {/* File Input */}
      <Card>
        <FileDropZone
          selectedFiles={selectedFiles}
          onSelectFiles={handleSelectFiles}
          onRemoveFile={removeFile}
        />
        {selectedFiles.length > 0 && (
          <FileList files={selectedFiles} onRemove={removeFile} className="mt-3" />
        )}
      </Card>

      {/* Mode & Preset */}
      <Card>
        <div className="grid grid-cols-2 gap-6">
          <div>
            <h3 className="text-sm font-semibold text-[var(--color-ink)] mb-3">处理模式</h3>
            <RadioGroup
              name="processMode"
              value={processMode}
              onChange={(v) => setProcessMode(v as ProcessMode)}
              options={MODE_OPTIONS}
            />
          </div>
          <div className="space-y-4">
            <Select
              label="排版预设方案"
              options={presetOptions}
              value={activePresetId}
              onChange={(e) => setActivePresetId(e.target.value)}
            />
            <div>
              <label className="text-xs font-medium text-[var(--color-ink-light)] block mb-1.5">
                输出后缀
              </label>
              <div className="flex items-center gap-2">
                <span className="text-xs text-[var(--color-ink-muted)]">原文件名</span>
                <input
                  value={outputSuffix}
                  onChange={(e) => setOutputSuffix(e.target.value)}
                  className="w-32 px-2 py-1 text-xs bg-[var(--color-paper)] border border-[var(--color-border-medium)] rounded text-center text-[var(--color-vermillion)] font-medium"
                />
                <span className="text-xs text-[var(--color-ink-muted)]">.docx</span>
              </div>
            </div>
            {activePreset && (
              <div className="flex flex-wrap gap-1.5">
                <Badge variant="default">{activePreset.page.top}/{activePreset.page.bottom}/{activePreset.page.left}/{activePreset.page.right} cm</Badge>
                <Badge variant="success">{activePreset.elements.body.fontCn} {activePreset.elements.body.size}pt</Badge>
                <Badge variant="default">行距 {activePreset.elements.body.lineSpacing}pt</Badge>
              </div>
            )}
          </div>
        </div>
      </Card>

      {/* Process */}
      <div className="flex items-center gap-4">
        <Button
          size="lg"
          onClick={handleProcess}
          disabled={isProcessing || selectedFiles.length === 0}
        >
          {isProcessing ? (
            <>处理中...</>
          ) : (
            <><Play size={16} /> 开始处理</>
          )}
        </Button>
        {selectedFiles.length > 0 && (
          <span className="text-sm text-[var(--color-ink-muted)]">
            已选择 {selectedFiles.length} 个文件
          </span>
        )}
      </div>

      {/* Progress */}
      {isProcessing && (
        <Card>
          <ProgressBar progress={progress} />
        </Card>
      )}

      {/* Output placeholder */}
      {!isProcessing && selectedFiles.length === 0 && (
        <Card className="text-center py-12 border-dashed">
          <FileOutput size={36} className="mx-auto text-[var(--color-border-medium)] mb-3" />
          <p className="text-sm text-[var(--color-ink-muted)]">
            处理后的文件将在原文件旁自动生成，原文件不会被覆盖
          </p>
          <p className="text-xs text-[var(--color-ink-muted)] mt-1">
            支持 .docx / .doc / .wps 格式，可批量处理多个文件
          </p>
        </Card>
      )}
    </div>
  );
}
