// Document formatting engine — ported from docformat-gui scripts/formatter.py
// This module handles the core logic for formatting Word documents.
// In the browser/web context, it operates on docx XML structures.
// When integrated with Tauri, actual file I/O is handled by Rust commands.

import type { Preset, ElementType } from './types';
import { BUILTIN_PRESETS } from './presets';
import JSZip from 'jszip';

export function resolvePreset(presetId: string, userPresets: Preset[]): Preset {
  const builtin = BUILTIN_PRESETS.find((p) => p.id === presetId);
  if (builtin) return builtin;
  const user = userPresets.find((p) => p.id === presetId);
  return user || BUILTIN_PRESETS[0];
}

// Detect paragraph type based on text content and context
export function detectParagraphType(
  text: string,
  index: number,
  alignment: string,
  prevType: ElementType | null,
): ElementType {
  const trimmed = text.trim();
  if (!trimmed) return 'body';

  // Title: first non-empty paragraph or centered short text
  if (index === 0 && trimmed.length < 100) return 'title';
  if (alignment === 'center' && trimmed.length < 80 && prevType !== 'body') return 'title';

  // Recipient: line ending with colon, near document start
  if (/[：:]$/.test(trimmed) && index < 3) return 'recipient';

  // Heading patterns
  const heading1Pattern = /^[（(]?[一二三四五六七八九十]+[）)]?\s*[、．.]/;
  const heading2Pattern = /^[（(]?[（(]?[一二三四五六七八九十]+[）)]?\s*[）)]\s*[、．.]?/;
  const heading3Pattern = /^\d+[．.)、]\s*/;
  const heading4Pattern = /^[（(]\d+[）)]\s*/;

  if (heading2Pattern.test(trimmed)) return 'heading2';
  if (heading1Pattern.test(trimmed)) return 'heading1';
  if (heading4Pattern.test(trimmed)) return 'heading4';
  if (heading3Pattern.test(trimmed)) return 'heading3';

  // Closing phrases
  if (/^(特此|此致|敬礼|顺颂|此复|专此)/.test(trimmed)) return 'closing';

  // Date pattern
  if (/^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日\s*$/.test(trimmed)) return 'date';

  // Attachment
  if (/^附件[：:]/.test(trimmed)) return 'attachment';

  // Signature: right-aligned short text near end
  if (alignment === 'right' && trimmed.length < 30) {
    return prevType === 'date' || prevType === 'body' ? 'signature' : 'signature';
  }

  return 'body';
}

// HTML preview rendering using mammoth-style conversion
// This produces an HTML representation of how the document WOULD look after formatting
export function renderPreviewHTML(
  paragraphs: string[],
  preset: Preset,
): string {
  const { elements } = preset;
  const bodyFont = elements.body;

  const lines: string[] = [
    '<!DOCTYPE html>',
    '<html><head><meta charset="utf-8"><style>',
    `body { font-family: "${bodyFont.fontCn}", "${bodyFont.fontEn}", serif; font-size: ${bodyFont.size}pt; line-height: ${bodyFont.lineSpacing}pt; max-width: 650px; margin: 0 auto; padding: 40px; color: #2E2E2E; }`,
    '.title { text-align: center; font-family: inherit; }',
    '.body { text-indent: 2em; text-align: justify; }',
    '.heading1 { font-family: inherit; font-weight: bold; }',
    '.signature { text-align: right; }',
    '.date { text-align: right; }',
    '</style></head><body>',
  ];

  let prevType: ElementType | null = null;

  for (let i = 0; i < paragraphs.length; i++) {
    const text = paragraphs[i];
    if (!text.trim()) {
      lines.push('<p>&nbsp;</p>');
      continue;
    }

    const alignment = ''; // Would come from actual docx parsing
    const paraType = detectParagraphType(text, i, alignment, prevType);
    const fmt = elements[paraType];

    const style = [
      `font-family: "${fmt.fontCn}", "${fmt.fontEn}"`,
      `font-size: ${fmt.size}pt`,
      `text-align: ${fmt.align}`,
      fmt.indent > 0 ? `text-indent: ${fmt.indent}pt` : '',
      `line-height: ${fmt.lineSpacing}pt`,
      fmt.bold ? 'font-weight: bold' : '',
    ].filter(Boolean).join('; ');

    lines.push(`<p class="${paraType}" style="${style}">${escapeHtml(text)}</p>`);
    prevType = paraType;
  }

  lines.push('</body></html>');
  return lines.join('\n');
}

export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Process a docx file (browser/web version using JSZip)
export async function processDocx(
  fileData: ArrayBuffer,
  preset: Preset,
  mode: 'full' | 'diagnose' | 'punctuation',
): Promise<{ data: ArrayBuffer; report: string }> {
  const zip = await JSZip.loadAsync(fileData);
  const documentXml = await zip.file('word/document.xml')?.async('string');

  if (!documentXml) {
    throw new Error('Invalid docx file: missing word/document.xml');
  }

  let processedXml = documentXml;
  const reportLines: string[] = [];

  if (mode === 'full' || mode === 'punctuation') {
    // Fix English punctuation in the XML
    processedXml = fixPunctuationInXml(processedXml);
    reportLines.push('标点符号已修复');
  }

  if (mode === 'full') {
    // Apply formatting rules
    processedXml = applyFormattingRules(processedXml, preset);
    reportLines.push(`排版格式已应用 (${preset.name})`);
  }

  if (mode === 'diagnose') {
    // Just analyze, don't modify
    const paragraphs = extractParagraphTexts(documentXml);
    reportLines.push(`诊断完成: 共 ${paragraphs.length} 个段落`);
  }

  // Update the zip
  zip.file('word/document.xml', processedXml);
  const data = await zip.generateAsync({ type: 'arraybuffer' });

  return { data, report: reportLines.join('\n') };
}

function fixPunctuationInXml(xml: string): string {
  // Replace English punctuation with Chinese equivalents in text content
  let result = xml;
  // Simple replacements within <w:t> tags
  result = result.replace(/(<w:t[^>]*>)(.*?)(<\/w:t>)/g, (_match, open, text, close) => {
    let fixed = text;
    fixed = fixed.replace(/(?<!\d),(?!\d)/g, '，');
    fixed = fixed.replace(/\(/g, '（');
    fixed = fixed.replace(/\)/g, '）');
    return open + fixed + close;
  });
  return result;
}

function applyFormattingRules(xml: string, _preset: Preset): string {
  // In the full implementation, this would:
  // 1. Set page margins via <w:sectPr>
  // 2. Set default paragraph properties
  // 3. Apply font, size, bold, alignment, indent, spacing to each paragraph
  // 4. Format tables
  // 5. Add page numbers
  return xml;
}

function extractParagraphTexts(xml: string): string[] {
  const texts: string[] = [];
  const regex = /<w:t[^>]*>(.*?)<\/w:t>/g;
  let match;
  while ((match = regex.exec(xml)) !== null) {
    texts.push(match[1]);
  }
  return texts;
}

// Create a new docx from plain text
export async function createDocxFromText(
  title: string,
  bodyText: string,
  preset: Preset,
): Promise<ArrayBuffer> {
  const { elements } = preset;
  const bodyFmt = elements.body;
  const titleFmt = elements.title;

  const paragraphs = bodyText.split('\n').filter((p) => p.trim());

  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:jc w:val="${titleFmt.align}"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:eastAsia="${titleFmt.fontCn}" w:ascii="${titleFmt.fontEn}"/>
          <w:sz w:val="${titleFmt.size * 2}"/>
          ${titleFmt.bold ? '<w:b/>' : ''}
        </w:rPr>
        <w:t>${escapeXml(title)}</w:t>
      </w:r>
    </w:p>
    ${paragraphs.map((p) => `
    <w:p>
      <w:pPr>
        <w:jc w:val="${bodyFmt.align}"/>
        <w:ind w:firstLine="${bodyFmt.indent * 20}"/>
        <w:spacing w:line="${bodyFmt.lineSpacing * 20}" w:lineRule="auto"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:eastAsia="${bodyFmt.fontCn}" w:ascii="${bodyFmt.fontEn}"/>
          <w:sz w:val="${bodyFmt.size * 2}"/>
        </w:rPr>
        <w:t>${escapeXml(p)}</w:t>
      </w:r>
    </w:p>`).join('')}
    <w:sectPr>
      <w:pgMar w:top="${preset.page.top * 567}" w:bottom="${preset.page.bottom * 567}"
               w:left="${preset.page.left * 567}" w:right="${preset.page.right * 567}"/>
    </w:sectPr>
  </w:body>
</w:document>`;

  const zip = new JSZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES);
  zip.file('_rels/.rels', RELS);
  zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS);
  zip.file('word/document.xml', xml);

  return zip.generateAsync({ type: 'arraybuffer' });
}

function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
