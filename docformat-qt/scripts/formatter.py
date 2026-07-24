# -*- coding: utf-8 -*-
"""
文档格式统一 — 向后兼容入口

本文件保留所有原有导出，实际逻辑已拆分到独立模块：
  data_model  - PRESETS, 字体适配, 预设加载
  font        - set_font, 修订标记
  page        - 页面设置, 文档网格, 页码
  paragraph   - 段落格式化, 深度清洗
  detector    - 段落类型检测, 识别规则
  table       - 表格格式化, 列宽, 边框
  signature   - 落款对位
  engine      - 主控编排

所有 from scripts.formatter import xxx 的调用无需修改。
"""

import logging

# ── 数据模型 & 预设 ──
from .data_model import (
    PRESETS,
    MACOS_FONT_FALLBACK,
    MACOS_FONT_ALIASES,
    load_custom_preset,
    get_missing_cn_fonts,
    _adapt_fonts_for_platform,
    _merge_preset_settings,
    _resolve_font_for_macos,
    _get_macos_installed_fonts,
)

# ── 字体 ──
from .font import (
    set_font,
    _force_normal_style,
    _next_rev_id,
    _rev_date,
    _add_ppr_change,
    _add_rpr_change,
    reset_revision_counter,
)

# ── 页面 ──
from .page import (
    _apply_page_grid,
    _set_normal_style_font,
    add_page_number,
    remove_background,
    _strip_autospacing_from_styles,
)

# ── 段落 ──
from .paragraph import (
    format_paragraph,
    _set_paragraph_spacing_points,
    _compact_empty_paragraph,
    _mark_structural_blank,
    _is_structural_blank,
    _format_structural_blank_paragraph,
    _format_empty_paragraphs,
    deep_clean_document,
    sanitize_document,
    paragraph_has_media,
    protect_media_paragraph,
)

# ── 检测器 ──
from .detector import (
    detect_para_type,
    _compile_rules,
    _build_text_context,
    _is_date_text,
    _normalize_date_text,
    _standardize_date_text,
    DEFAULT_DETECT_RULES,
)

# ── 表格 ──
from .table import (
    _iter_block_items,
    _set_table_borders,
    _set_table_cell_margins,
    _set_table_width_percent,
    _set_table_indent,
    _set_table_col_widths_by_content,
    _set_cell_borders,
    _is_numeric_text,
    _is_short_text,
    _is_table_title,
    _is_table_unit,
    _text_weight,
    _normalize_pcts,
    _insert_paragraph_after_table,
    _insert_paragraph_before_table,
    _insert_paragraph_after_paragraph,
    _insert_paragraph_before_paragraph,
    _split_heading_by_punct,
    _set_header_row_repeat,
    _set_cell_shading,
    _set_cell_vertical_alignment,
    _find_nested_tables,
    _BORDER_STYLES,
)

# ── 落款 ──
from .signature import _apply_gb_signature_layout

# ── 页眉页脚 ──
from .header_footer import (
    apply_header_footer,
    _get_default_hf_config,
    _format_header_footer_part,
    DEFAULT_HEADER_FOOTER_CONFIG,
)

# ── 图片处理 ──
from .image import (
    PictureConfig,
    DEFAULT_PICTURE_CONFIG,
    ImageCompressResult,
    compress_images,
    apply_image_size,
    apply_image_wrapping,
    apply_image_alignment,
    _find_image_parts,
    _replace_image_blob,
    _compress_single_image,
    _unit_to_emu,
    _calculate_image_size,
    _resize_drawing_element,
)

# ── 水印 ──
from .watermark import (
    WatermarkConfig,
    DEFAULT_WATERMARK_CONFIG,
    WATERMARK_PRESETS,
    apply_watermark,
    remove_watermark,
    get_watermark_preset,
)

# ── 引擎主控 ──
from .engine import format_document, _ensure_structural_blank_lines

logger = logging.getLogger('docformat.formatter')

# ── CLI 入口 ──
if __name__ == '__main__':
    import sys
    logging.basicConfig(level=logging.INFO, format='%(message)s')

    if len(sys.argv) < 3:
        print('Usage: python -m scripts.formatter input.docx output.docx [--preset official|official_gbk|academic|legal]')
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]

    preset = 'official'
    if '--preset' in sys.argv:
        idx = sys.argv.index('--preset')
        if idx + 1 < len(sys.argv):
            preset = sys.argv[idx + 1]

    format_document(input_file, output_file, preset_name=preset)
