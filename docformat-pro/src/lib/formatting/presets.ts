import type { Preset, ElementFormat } from './types';

function el(overrides: Partial<ElementFormat> = {}): ElementFormat {
  return {
    fontCn: '仿宋_GB2312',
    fontEn: 'Times New Roman',
    size: 16,
    bold: false,
    align: 'justify',
    indent: 32,
    lineSpacing: 28,
    spaceBefore: 0,
    spaceAfter: 0,
    ...overrides,
  };
}

export const BUILTIN_PRESETS: Preset[] = [
  {
    id: 'official',
    name: '公文格式',
    scope: 'builtin',
    page: { top: 3.7, bottom: 3.5, left: 2.8, right: 2.6 },
    elements: {
      title: el({ fontCn: '方正小标宋简体', size: 22, align: 'center', indent: 0 }),
      recipient: el({ align: 'left', indent: 0 }),
      heading1: el({ fontCn: '黑体', align: 'left' }),
      heading2: el({ fontCn: '楷体_GB2312', align: 'left' }),
      heading3: el({ bold: true, align: 'left' }),
      heading4: el({ align: 'left' }),
      body: el(),
      signature: el({ align: 'right', indent: 0 }),
      date: el({ align: 'right', indent: 0 }),
      attachment: el(),
      closing: el({ align: 'left' }),
    },
    table: {
      fontCn: '仿宋_GB2312', fontEn: 'Times New Roman', size: 12,
      bold: false, lineSpacing: 22, firstLineIndent: 0,
      headerBold: true, smartAlign: false,
    },
    pageNumber: {
      enabled: true, font: '宋体', size: 12,
      style: 'dash', position: 'outside', offset: 0.7,
    },
    options: {
      spaceHandling: 'remove_all',
      firstLineBold: false,
      boldSerial: true,
    },
  },
  {
    id: 'academic',
    name: '学术论文',
    scope: 'builtin',
    page: { top: 2.5, bottom: 2.5, left: 2.5, right: 2.5 },
    elements: {
      title: el({ fontCn: '黑体', size: 18, bold: true, align: 'center', indent: 0 }),
      recipient: el({ align: 'left', indent: 0 }),
      heading1: el({ fontCn: '黑体', size: 14, bold: true, align: 'left', indent: 0, lineSpacing: 22 }),
      heading2: el({ fontCn: '楷体', size: 14, bold: true, align: 'left', indent: 0, lineSpacing: 22 }),
      heading3: el({ size: 14, bold: true, align: 'left', indent: 0, lineSpacing: 22 }),
      heading4: el({ size: 14, align: 'left', indent: 0, lineSpacing: 22 }),
      body: el({ fontCn: '宋体', size: 12, indent: 24, lineSpacing: 20 }),
      signature: el({ fontCn: '宋体', size: 12, align: 'right', indent: 0, lineSpacing: 20 }),
      date: el({ fontCn: '宋体', size: 12, align: 'right', indent: 0, lineSpacing: 20 }),
      attachment: el({ fontCn: '宋体', size: 12, indent: 24, lineSpacing: 20 }),
      closing: el({ fontCn: '宋体', size: 12, align: 'left', indent: 0, lineSpacing: 20 }),
    },
    table: {
      fontCn: '宋体', fontEn: 'Times New Roman', size: 10.5,
      bold: false, lineSpacing: 18, firstLineIndent: 0,
      headerBold: true, smartAlign: true,
    },
    pageNumber: {
      enabled: true, font: 'Times New Roman', size: 10.5,
      style: 'plain', position: 'center', offset: 0.7,
    },
    options: {
      spaceHandling: 'normalize',
      firstLineBold: false,
      boldSerial: false,
    },
  },
  {
    id: 'legal',
    name: '法律文书',
    scope: 'builtin',
    page: { top: 3.0, bottom: 2.5, left: 3.0, right: 2.5 },
    elements: {
      title: el({ fontCn: '宋体', size: 22, bold: true, align: 'center', indent: 0 }),
      recipient: el({ fontCn: '宋体', align: 'left', indent: 0 }),
      heading1: el({ fontCn: '黑体', size: 14, bold: true, align: 'left', indent: 28 }),
      heading2: el({ fontCn: '宋体', size: 14, bold: true, align: 'left', indent: 28 }),
      heading3: el({ fontCn: '宋体', size: 14, bold: true, align: 'left', indent: 28 }),
      heading4: el({ fontCn: '宋体', size: 14, align: 'left', indent: 28 }),
      body: el({ fontCn: '宋体', size: 14, indent: 28 }),
      signature: el({ fontCn: '宋体', size: 14, align: 'right', indent: 0 }),
      date: el({ fontCn: '宋体', size: 14, align: 'right', indent: 0 }),
      attachment: el({ fontCn: '宋体', size: 14, indent: 28 }),
      closing: el({ fontCn: '宋体', size: 14, align: 'left', indent: 0 }),
    },
    table: {
      fontCn: '宋体', fontEn: 'Times New Roman', size: 12,
      bold: false, lineSpacing: 22, firstLineIndent: 0,
      headerBold: true, smartAlign: true,
    },
    pageNumber: {
      enabled: true, font: '宋体', size: 12,
      style: 'plain', position: 'center', offset: 0.7,
    },
    options: {
      spaceHandling: 'normalize',
      firstLineBold: false,
      boldSerial: true,
    },
  },
];

export function createDefaultUserPreset(name: string): Preset {
  const official = BUILTIN_PRESETS[0];
  return {
    ...JSON.parse(JSON.stringify(official)),
    id: `user_${Date.now()}`,
    name,
    scope: 'user',
  };
}
