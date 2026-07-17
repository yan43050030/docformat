// Document analysis engine — ported from docformat-gui scripts/analyzer.py

export interface PunctuationIssue {
  type: 'english_paren' | 'english_quote' | 'english_comma' | 'english_period' | 'nonstandard_ellipsis' | 'chinese_period_fix';
  text: string;
  position: number;
  suggestion: string;
}

export interface NumberingIssue {
  type: 'mixed_style' | 'inconsistent';
  text: string;
  detail: string;
}

export interface ParagraphIssue {
  type: 'missing_indent' | 'inconsistent_spacing';
  text: string;
  detail: string;
}

export interface FontIssue {
  type: 'too_many_fonts' | 'too_many_sizes';
  detail: string;
}

export interface AnalysisResult {
  punctuation: PunctuationIssue[];
  numbering: NumberingIssue[];
  paragraphs: ParagraphIssue[];
  fonts: FontIssue[];
  summary: string;
}

const ENGLISH_PARENS = /[()]/g;
const ENGLISH_QUOTES = /["']/g;
const CHINESE_PERIOD_AFTER_CHINESE = /([一-鿿])\s*\.(?=\s|$)/g;

export function analyzePunctuation(text: string): PunctuationIssue[] {
  const issues: PunctuationIssue[] = [];

  // Check for English parentheses
  let match: RegExpExecArray | null;
  const parenRegex = new RegExp(ENGLISH_PARENS.source, 'g');
  while ((match = parenRegex.exec(text)) !== null) {
    issues.push({
      type: 'english_paren',
      text: match[0],
      position: match.index,
      suggestion: match[0] === '(' ? '（' : '）',
    });
  }

  // Check for English quotes
  const quoteRegex = new RegExp(ENGLISH_QUOTES.source, 'g');
  while ((match = quoteRegex.exec(text)) !== null) {
    issues.push({
      type: 'english_quote',
      text: match[0],
      position: match.index,
      suggestion: match[0] === '"' ? '“' : '‘',
    });
  }

  // Check for non-standard ellipsis (multiple periods)
  if (/\.{3,}/.test(text) && !/……/.test(text)) {
    issues.push({
      type: 'nonstandard_ellipsis',
      text: '...',
      position: text.indexOf('...'),
      suggestion: '……',
    });
  }

  // Check for Chinese character followed by English period
  const periodRegex = new RegExp(CHINESE_PERIOD_AFTER_CHINESE.source, 'g');
  while ((match = periodRegex.exec(text)) !== null) {
    issues.push({
      type: 'chinese_period_fix',
      text: match[0],
      position: match.index,
      suggestion: match[0].replace('.', '。'),
    });
  }

  return issues;
}

export function analyzeNumbering(paragraphs: string[]): NumberingIssue[] {
  const issues: NumberingIssue[] = [];
  let hasChineseNum = false;
  let hasArabicNum = false;

  const chinesePattern = /^[（(]?[一二三四五六七八九十]+[）)]?\s*[、．.]/;
  const arabicPattern = /^\d+[．.)、]\s*/;

  for (const para of paragraphs) {
    if (chinesePattern.test(para)) hasChineseNum = true;
    if (arabicPattern.test(para)) hasArabicNum = true;
    if (hasChineseNum && hasArabicNum) break;
  }

  if (hasChineseNum && hasArabicNum) {
    issues.push({
      type: 'mixed_style',
      text: '',
      detail: '文档中同时存在中文序号（一、）和阿拉伯序号（1.），建议统一',
    });
  }

  return issues;
}

const NO_INDENT_PATTERNS = [
  /^附件[：:]/,
  /^联系人[：:]/,
  /^抄送[：:]/,
  /^主送[：:]/,
  /^抄报[：:]/,
];

export function analyzeParagraphFormat(paragraphs: string[], alignments: string[]): ParagraphIssue[] {
  const issues: ParagraphIssue[] = [];

  for (let i = 0; i < paragraphs.length; i++) {
    const text = paragraphs[i];
    const alignment = alignments[i] || '';

    if (!text.trim()) continue;

    const isCentered = alignment === 'center' || alignment === 'CENTER';
    const isNoIndent = NO_INDENT_PATTERNS.some((p) => p.test(text));

    if (!isCentered && !isNoIndent) {
      // Check for missing first-line indent
      // In the actual implementation, this checks paragraph formatting properties
    }
  }

  return issues;
}

export function analyzeFonts(fontNames: string[], fontSizes: number[]): FontIssue[] {
  const issues: FontIssue[] = [];
  const uniqueFonts = new Set(fontNames);
  const uniqueSizes = new Set(fontSizes);

  if (uniqueFonts.size > 4) {
    issues.push({
      type: 'too_many_fonts',
      detail: `文档使用了 ${uniqueFonts.size} 种不同字体: ${[...uniqueFonts].join(', ')}`,
    });
  }

  if (uniqueSizes.size > 5) {
    issues.push({
      type: 'too_many_sizes',
      detail: `文档使用了 ${uniqueSizes.size} 种不同字号`,
    });
  }

  return issues;
}

export function generateReport(results: {
  punctuation: PunctuationIssue[];
  numbering: NumberingIssue[];
  paragraphs: ParagraphIssue[];
  fonts: FontIssue[];
}): string {
  const lines: string[] = [];
  let totalIssues = 0;

  if (results.punctuation.length > 0) {
    lines.push(`【标点符号】发现 ${results.punctuation.length} 个问题`);
    for (const issue of results.punctuation) {
      lines.push(`  - ${issue.type}: "${issue.text}" → 建议 "${issue.suggestion}"`);
    }
    totalIssues += results.punctuation.length;
    lines.push('');
  }

  if (results.numbering.length > 0) {
    lines.push(`【序号格式】发现 ${results.numbering.length} 个问题`);
    for (const issue of results.numbering) {
      lines.push(`  - ${issue.detail}`);
    }
    totalIssues += results.numbering.length;
    lines.push('');
  }

  if (results.paragraphs.length > 0) {
    lines.push(`【段落格式】发现 ${results.paragraphs.length} 个问题`);
    for (const issue of results.paragraphs) {
      lines.push(`  - ${issue.detail}`);
    }
    totalIssues += results.paragraphs.length;
    lines.push('');
  }

  if (results.fonts.length > 0) {
    lines.push(`【字体】发现 ${results.fonts.length} 个问题`);
    for (const issue of results.fonts) {
      lines.push(`  - ${issue.detail}`);
    }
    totalIssues += results.fonts.length;
    lines.push('');
  }

  if (totalIssues === 0) {
    lines.push('未发现格式问题，文档格式符合规范。');
  } else {
    lines.unshift(`\n格式诊断报告 — 共发现 ${totalIssues} 个问题:\n`);
  }

  return lines.join('\n');
}
