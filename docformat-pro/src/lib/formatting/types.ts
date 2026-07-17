export type ElementType =
  | 'title'
  | 'recipient'
  | 'heading1'
  | 'heading2'
  | 'heading3'
  | 'heading4'
  | 'body'
  | 'signature'
  | 'date'
  | 'attachment'
  | 'closing';

export type AlignMode = 'left' | 'center' | 'right' | 'justify';
export type PageNumberStyle = 'dash' | 'plain' | 'page_text' | 'page_total';
export type PageNumberPosition = 'outside' | 'center' | 'right' | 'left';
export type SpaceHandling = 'remove_all' | 'normalize' | 'keep';
export type ProcessMode = 'full' | 'diagnose' | 'punctuation' | 'ai_paste';

export interface ElementFormat {
  fontCn: string;
  fontEn: string;
  size: number;
  bold: boolean;
  align: AlignMode;
  indent: number;
  lineSpacing: number;
  spaceBefore: number;
  spaceAfter: number;
}

export interface TableFormat {
  fontCn: string;
  fontEn: string;
  size: number;
  bold: boolean;
  lineSpacing: number;
  firstLineIndent: number;
  headerBold: boolean;
  smartAlign: boolean;
}

export interface PageSettings {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

export interface PageNumberSettings {
  enabled: boolean;
  font: string;
  size: number;
  style: PageNumberStyle;
  position: PageNumberPosition;
  offset: number;
}

export interface AdvancedOptions {
  spaceHandling: SpaceHandling;
  firstLineBold: boolean;
  boldSerial: boolean;
}

export interface Preset {
  id: string;
  name: string;
  scope: 'builtin' | 'user';
  page: PageSettings;
  elements: Record<ElementType, ElementFormat>;
  table: TableFormat;
  pageNumber: PageNumberSettings;
  options: AdvancedOptions;
}

export const ELEMENT_LABELS: Record<ElementType, string> = {
  title: '标题',
  recipient: '主送机关',
  heading1: '一级标题',
  heading2: '二级标题',
  heading3: '三级标题',
  heading4: '四级标题',
  body: '正文',
  signature: '署名',
  date: '日期',
  attachment: '附件',
  closing: '结尾',
};

export const FONT_SIZES = [
  { label: '初号 (42pt)', value: 42 },
  { label: '小初 (36pt)', value: 36 },
  { label: '一号 (26pt)', value: 26 },
  { label: '小一 (24pt)', value: 24 },
  { label: '二号 (22pt)', value: 22 },
  { label: '小二 (18pt)', value: 18 },
  { label: '三号 (16pt)', value: 16 },
  { label: '小三 (15pt)', value: 15 },
  { label: '四号 (14pt)', value: 14 },
  { label: '小四 (12pt)', value: 12 },
  { label: '五号 (10.5pt)', value: 10.5 },
  { label: '小五 (9pt)', value: 9 },
];

export const CN_FONTS = [
  '方正小标宋简体', '方正仿宋_GBK', '仿宋_GB2312', '仿宋',
  '黑体', '楷体_GB2312', '楷体', '宋体', '华文中宋',
  '方正书宋_GBK', '方正楷体_GBK', '方正黑体_GBK',
];

export const EN_FONTS = [
  'Times New Roman', 'Arial', 'Calibri', 'Cambria',
  'Georgia', 'Garamond', 'Palatino Linotype',
];
