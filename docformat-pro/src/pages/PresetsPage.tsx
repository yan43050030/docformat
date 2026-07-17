import { useState } from 'react';
import { useAppStore } from '@/lib/store';
import { Button } from '@/components/ui/button';
import { Card } from '@/components/ui/card';
import { Select } from '@/components/ui/select';
import { Switch } from '@/components/ui/switch';
import { createDefaultUserPreset } from '@/lib/formatting/presets';
import {
  ELEMENT_LABELS, FONT_SIZES, CN_FONTS, EN_FONTS,
  type Preset, type ElementType, type ElementFormat,
} from '@/lib/formatting/types';
import {
  Plus, Copy, Trash2, Download, Upload, ChevronDown, ChevronRight, FileText,
} from 'lucide-react';

export function PresetsPage() {
  const {
    presets, activePresetId, setActivePresetId,
    addPreset, removePreset, updatePreset, duplicatePreset,
  } = useAppStore();

  const preset = presets.find((p) => p.id === activePresetId);
  const [expandedElements, setExpandedElements] = useState<Set<string>>(
    new Set(['page', 'title', 'body']),
  );

  const toggleExpand = (key: string) => {
    setExpandedElements((prev) => {
      const next = new Set(prev);
      next.has(key) ? next.delete(key) : next.add(key);
      return next;
    });
  };

  const handleNew = () => {
    const name = `自定义格式 ${presets.filter((p) => p.scope === 'user').length + 1}`;
    addPreset(createDefaultUserPreset(name));
  };

  const handleImport = () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    input.onchange = async () => {
      const file = input.files?.[0];
      if (!file) return;
      try {
        const text = await file.text();
        const data = JSON.parse(text);
        if (Array.isArray(data)) {
          data.forEach((p: Preset) => addPreset({ ...p, scope: 'user' }));
        } else {
          addPreset({ ...data, scope: 'user' });
        }
      } catch {
        alert('导入失败：文件格式不正确');
      }
    };
    input.click();
  };

  const handleExport = () => {
    if (!preset) return;
    const blob = new Blob([JSON.stringify(preset, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${preset.name}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const updateElement = (type: ElementType, field: keyof ElementFormat, value: unknown) => {
    if (!preset) return;
    updatePreset({
      ...preset,
      elements: {
        ...preset.elements,
        [type]: { ...preset.elements[type], [field]: value },
      },
    });
  };

  const presetOptions = presets.map((p) => ({ label: `${p.name} ${p.scope === 'builtin' ? '(内置)' : ''}`, value: p.id }));

  if (!preset) return null;
  const isBuiltin = preset.scope === 'builtin';

  return (
    <div className="max-w-3xl mx-auto px-6 py-8 space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-xl font-bold text-[var(--color-ink)]">预设方案管理</h1>
          <p className="text-sm text-[var(--color-ink-muted)] mt-1">可视化编辑排版参数，所有修改自动保存</p>
        </div>
      </div>

      {/* Preset Selector */}
      <Card>
        <div className="flex items-center gap-3">
          <Select
            options={presetOptions}
            value={activePresetId}
            onChange={(e) => setActivePresetId(e.target.value)}
            className="flex-1"
          />
          <Button variant="secondary" size="sm" onClick={handleNew}><Plus size={14} /> 新建</Button>
          <Button variant="secondary" size="sm" onClick={() => duplicatePreset(activePresetId)}><Copy size={14} /> 复制</Button>
          {!isBuiltin && (
            <Button variant="danger" size="sm" onClick={() => removePreset(activePresetId)}><Trash2 size={14} /> 删除</Button>
          )}
          <div className="w-px h-6 bg-[var(--color-border-light)]" />
          <Button variant="ghost" size="sm" onClick={handleImport}><Upload size={14} /> 导入</Button>
          <Button variant="ghost" size="sm" onClick={handleExport}><Download size={14} /> 导出</Button>
        </div>
        {isBuiltin && (
          <p className="text-xs text-[var(--color-ink-muted)] mt-2">内置预设不可删除，可复制后自定义修改</p>
        )}
      </Card>

      {/* Page Settings */}
      <AccordionSection
        title="页面设置"
        icon={<FileText size={16} />}
        expanded={expandedElements.has('page')}
        onToggle={() => toggleExpand('page')}
      >
        <div className="grid grid-cols-4 gap-3">
          {(['top', 'bottom', 'left', 'right'] as const).map((key) => (
            <div key={key} className="flex flex-col gap-1">
              <label className="text-[10px] font-medium text-[var(--color-ink-muted)] uppercase">
                {{ top: '上边距', bottom: '下边距', left: '左边距', right: '右边距' }[key]}
              </label>
              <div className="flex items-center gap-1">
                <input
                  type="number"
                  step="0.1"
                  min="0"
                  max="10"
                  value={preset.page[key]}
                  onChange={(e) => {
                    const v = parseFloat(e.target.value);
                    if (!isNaN(v)) updatePreset({ ...preset, page: { ...preset.page, [key]: v } });
                  }}
                  className="w-full px-2 py-1.5 text-sm bg-[var(--color-paper)] border border-[var(--color-border-medium)] rounded-lg text-center focus:outline-none focus:ring-2 focus:ring-[var(--color-vermillion)]/20"
                />
                <span className="text-xs text-[var(--color-ink-muted)]">cm</span>
              </div>
            </div>
          ))}
        </div>
        <div className="grid grid-cols-3 gap-4 mt-4 pt-4 border-t border-[var(--color-border-light)]">
          <Select
            label="页码字体"
            options={CN_FONTS.map((f) => ({ label: f, value: f }))}
            value={preset.pageNumber.font}
            onChange={(e) => updatePreset({ ...preset, pageNumber: { ...preset.pageNumber, font: e.target.value } })}
          />
          <Select
            label="页码样式"
            options={[
              { label: '两侧横线 - 1 -', value: 'dash' },
              { label: '纯数字 1', value: 'plain' },
              { label: '第1页', value: 'page_text' },
              { label: '1/10', value: 'page_total' },
            ]}
            value={preset.pageNumber.style}
            onChange={(e) => updatePreset({ ...preset, pageNumber: { ...preset.pageNumber, style: e.target.value as any } })}
          />
          <Select
            label="页码位置"
            options={[
              { label: '外侧', value: 'outside' },
              { label: '居中', value: 'center' },
              { label: '右侧', value: 'right' },
              { label: '左侧', value: 'left' },
            ]}
            value={preset.pageNumber.position}
            onChange={(e) => updatePreset({ ...preset, pageNumber: { ...preset.pageNumber, position: e.target.value as any } })}
          />
        </div>
      </AccordionSection>

      {/* Element Editors */}
      {(Object.entries(ELEMENT_LABELS) as [ElementType, string][]).map(([type, label]) => {
        const el = preset.elements[type];
        return (
          <AccordionSection
            key={type}
            title={label}
            expanded={expandedElements.has(type)}
            onToggle={() => toggleExpand(type)}
          >
            <div className="grid grid-cols-4 gap-3">
              <Select
                label="中文字体"
                options={CN_FONTS.map((f) => ({ label: f, value: f }))}
                value={el.fontCn}
                onChange={(e) => updateElement(type, 'fontCn', e.target.value)}
              />
              <Select
                label="英文字体"
                options={EN_FONTS.map((f) => ({ label: f, value: f }))}
                value={el.fontEn}
                onChange={(e) => updateElement(type, 'fontEn', e.target.value)}
              />
              <Select
                label="字号"
                options={FONT_SIZES.map((s) => ({ label: s.label, value: s.value }))}
                value={el.size}
                onChange={(e) => updateElement(type, 'size', Number(e.target.value))}
              />
              <Select
                label="对齐"
                options={[
                  { label: '左对齐', value: 'left' },
                  { label: '居中', value: 'center' },
                  { label: '右对齐', value: 'right' },
                  { label: '两端对齐', value: 'justify' },
                ]}
                value={el.align}
                onChange={(e) => updateElement(type, 'align', e.target.value)}
              />
            </div>
            <div className="grid grid-cols-4 gap-3 mt-3">
              <div className="flex flex-col gap-1.5">
                <label className="text-xs font-medium text-[var(--color-ink-light)]">首行缩进</label>
                <div className="flex items-center gap-1">
                  <input
                    type="number"
                    min="0"
                    max="100"
                    value={el.indent}
                    onChange={(e) => updateElement(type, 'indent', Number(e.target.value))}
                    className="w-full px-2 py-1.5 text-sm bg-[var(--color-paper)] border border-[var(--color-border-medium)] rounded-lg text-center focus:outline-none focus:ring-2 focus:ring-[var(--color-vermillion)]/20"
                  />
                  <span className="text-xs text-[var(--color-ink-muted)]">pt</span>
                </div>
              </div>
              <div className="flex flex-col gap-1.5">
                <label className="text-xs font-medium text-[var(--color-ink-light)]">行距</label>
                <div className="flex items-center gap-1">
                  <input
                    type="number"
                    min="10"
                    max="100"
                    value={el.lineSpacing}
                    onChange={(e) => updateElement(type, 'lineSpacing', Number(e.target.value))}
                    className="w-full px-2 py-1.5 text-sm bg-[var(--color-paper)] border border-[var(--color-border-medium)] rounded-lg text-center focus:outline-none focus:ring-2 focus:ring-[var(--color-vermillion)]/20"
                  />
                  <span className="text-xs text-[var(--color-ink-muted)]">pt</span>
                </div>
              </div>
              <div className="flex flex-col gap-1.5">
                <label className="text-xs font-medium text-[var(--color-ink-light)]">段前间距</label>
                <div className="flex items-center gap-1">
                  <input
                    type="number"
                    min="0"
                    max="50"
                    value={el.spaceBefore}
                    onChange={(e) => updateElement(type, 'spaceBefore', Number(e.target.value))}
                    className="w-full px-2 py-1.5 text-sm bg-[var(--color-paper)] border border-[var(--color-border-medium)] rounded-lg text-center focus:outline-none focus:ring-2 focus:ring-[var(--color-vermillion)]/20"
                  />
                  <span className="text-xs text-[var(--color-ink-muted)]">pt</span>
                </div>
              </div>
              <div className="flex flex-col gap-1.5">
                <label className="text-xs font-medium text-[var(--color-ink-light)]">段后间距</label>
                <div className="flex items-center gap-1">
                  <input
                    type="number"
                    min="0"
                    max="50"
                    value={el.spaceAfter}
                    onChange={(e) => updateElement(type, 'spaceAfter', Number(e.target.value))}
                    className="w-full px-2 py-1.5 text-sm bg-[var(--color-paper)] border border-[var(--color-border-medium)] rounded-lg text-center focus:outline-none focus:ring-2 focus:ring-[var(--color-vermillion)]/20"
                  />
                  <span className="text-xs text-[var(--color-ink-muted)]">pt</span>
                </div>
              </div>
            </div>
            <div className="mt-3 pt-3 border-t border-[var(--color-border-light)]">
              <Switch
                checked={el.bold}
                onChange={(v) => updateElement(type, 'bold', v)}
                label="加粗"
                id={`bold-${type}`}
              />
            </div>
          </AccordionSection>
        );
      })}

      {/* Table Settings */}
      <AccordionSection
        title="表格格式"
        expanded={expandedElements.has('table')}
        onToggle={() => toggleExpand('table')}
      >
        <div className="grid grid-cols-3 gap-3">
          <Select
            label="表格字体"
            options={CN_FONTS.map((f) => ({ label: f, value: f }))}
            value={preset.table.fontCn}
            onChange={(e) => updatePreset({ ...preset, table: { ...preset.table, fontCn: e.target.value } })}
          />
          <Select
            label="表格字号"
            options={FONT_SIZES.slice(6).map((s) => ({ label: s.label, value: s.value }))}
            value={preset.table.size}
            onChange={(e) => updatePreset({ ...preset, table: { ...preset.table, size: Number(e.target.value) } })}
          />
          <div className="flex flex-col gap-1.5">
            <label className="text-xs font-medium text-[var(--color-ink-light)]">行距</label>
            <div className="flex items-center gap-1">
              <input
                type="number" min="10" max="60"
                value={preset.table.lineSpacing}
                onChange={(e) => updatePreset({ ...preset, table: { ...preset.table, lineSpacing: Number(e.target.value) } })}
                className="w-full px-2 py-1.5 text-sm bg-[var(--color-paper)] border border-[var(--color-border-medium)] rounded-lg text-center focus:outline-none focus:ring-2 focus:ring-[var(--color-vermillion)]/20"
              />
              <span className="text-xs text-[var(--color-ink-muted)]">pt</span>
            </div>
          </div>
        </div>
        <div className="flex gap-6 mt-3 pt-3 border-t border-[var(--color-border-light)]">
          <Switch
            checked={preset.table.headerBold}
            onChange={(v) => updatePreset({ ...preset, table: { ...preset.table, headerBold: v } })}
            label="表头加粗"
            id="table-header-bold"
          />
          <Switch
            checked={preset.table.smartAlign}
            onChange={(v) => updatePreset({ ...preset, table: { ...preset.table, smartAlign: v } })}
            label="智能列对齐（数字靠右/短文本居中）"
            id="table-smart-align"
          />
        </div>
      </AccordionSection>

      {/* Advanced */}
      <AccordionSection
        title="高级选项"
        expanded={expandedElements.has('advanced')}
        onToggle={() => toggleExpand('advanced')}
      >
        <div className="space-y-3">
          <Select
            label="空格处理策略"
            options={[
              { label: '删除全部多余空格', value: 'remove_all' },
              { label: '规范化空格（保留必要空格）', value: 'normalize' },
              { label: '保持原样', value: 'keep' },
            ]}
            value={preset.options.spaceHandling}
            onChange={(e) => updatePreset({ ...preset, options: { ...preset.options, spaceHandling: e.target.value as any } })}
          />
          <div className="flex gap-6">
            <Switch
              checked={preset.options.firstLineBold}
              onChange={(v) => updatePreset({ ...preset, options: { ...preset.options, firstLineBold: v } })}
              label="首行加粗"
              id="first-line-bold"
            />
            <Switch
              checked={preset.options.boldSerial}
              onChange={(v) => updatePreset({ ...preset, options: { ...preset.options, boldSerial: v } })}
              label={'序号加粗（如“一、”“1.”）'}
              id="bold-serial"
            />
          </div>
        </div>
      </AccordionSection>
    </div>
  );
}

function AccordionSection({
  title,
  children,
  expanded,
  onToggle,
  icon,
}: {
  title: string;
  children: React.ReactNode;
  expanded: boolean;
  onToggle: () => void;
  icon?: React.ReactNode;
}) {
  return (
    <Card>
      <button
        onClick={onToggle}
        className="w-full flex items-center justify-between cursor-pointer group"
      >
        <div className="flex items-center gap-2">
          {icon}
          <span className="text-sm font-semibold text-[var(--color-ink)]">{title}</span>
        </div>
        {expanded ? <ChevronDown size={16} className="text-[var(--color-ink-muted)]" /> : <ChevronRight size={16} className="text-[var(--color-ink-muted)]" />}
      </button>
      {expanded && <div className="mt-4">{children}</div>}
    </Card>
  );
}
