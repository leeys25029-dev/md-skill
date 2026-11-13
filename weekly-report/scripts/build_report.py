#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_report.py — 周报 docx 生成器（周报管家专用）

根据 JSON 内容规格 + 周报模板（含命名样式），生成严格符合示范格式的周报 docx。

用法:
    python3 build_report.py <spec.json> --output <输出.docx> [--vault <Obsidian库路径>] [--template <模板.docx>]

spec.json 结构:
{
  "date": "2026-11-13",            # 可选，与 title 二选一
  "title": "2026-11-13 周报",       # 可选，优先于 date
  "blocks": [
    {"type": "h1",              "text": "主要内容"},
    {"type": "h2",              "text": "用10k个参数进行时序预测（FITS）"},
    {"type": "h3",              "text": "文章动机"},
    {"type": "para",            "text": "正文段落，首行缩进2字符"},
    {"type": "para",            "runs": ["文字", {"image": "xxx.png", "width": 1.9}]},
    {"type": "para_noindent",   "text": "1. 编号列表项，首行不缩进"},
    {"type": "web",             "text": "图片说明文字（Normal Web 样式）"},
    {"type": "image",           "src": "Screenshot xxx.png", "width": 5.0},
    {"type": "placeholder",     "text": "（灰色占位提示文字）"}
  ]
}

规则:
- h1 自动编号为 "N.标题"（如 1.主要内容），h2 自动编号为 "N.M 标题"（如 2.1 xxx），h3 不编号
- text 与 runs 二选一；runs 支持文字与内嵌图片混排
- image.src 支持绝对路径 / 相对路径 / 文件名（在 --vault 指定的 Obsidian 库中递归查找，优先 _images 目录）
- 图片 width 单位为英寸，缺省 4.5，上限 6.8；找不到的图片以红色【图片未找到】文字占位
"""

import sys
import os
import json
import argparse

# --- 自动安装 python-docx ---
try:
    import docx
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--user', 'python-docx'])
    import docx

from docx import Document
from docx.shared import Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.join(SCRIPT_DIR, '..', 'assets', 'weekly-report-template.docx')
DEFAULT_VAULT = os.path.expanduser(
    '~/Library/Mobile Documents/iCloud~md~obsidian/Documents/Tang\'s vault'
)
MAX_IMAGE_WIDTH_IN = 6.8
DEFAULT_IMAGE_WIDTH_IN = 4.5


class ImageResolver:
    """按 绝对路径 → 相对路径 → Obsidian 库文件名 的顺序解析图片"""

    def __init__(self, vault=None):
        self.vault = vault
        self._index = None  # 文件名 -> [完整路径]

    def _build_index(self):
        self._index = {}
        if not self.vault or not os.path.isdir(self.vault):
            return
        for root, _dirs, files in os.walk(self.vault):
            for f in files:
                if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp')):
                    self._index.setdefault(f, []).append(os.path.join(root, f))

    def resolve(self, src):
        if os.path.isabs(src) and os.path.isfile(src):
            return src
        if os.path.isfile(src):
            return os.path.abspath(src)
        if self._index is None:
            self._build_index()
        candidates = self._index.get(os.path.basename(src))
        if candidates:
            # 优先 _images / attachments 目录下的命中
            for c in candidates:
                if '_images' in c or 'attachments' in c.lower():
                    return c
            return candidates[0]
        return None


def set_no_indent(paragraph):
    """首行缩进与左缩进清零（用于编号列表项，严格对齐示范周报的做法）"""
    pPr = paragraph._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:left'), '0')
    ind.set(qn('w:leftChars'), '0')
    ind.set(qn('w:firstLine'), '0')
    ind.set(qn('w:firstLineChars'), '0')


def add_runs(paragraph, runs, resolver, warnings):
    """向段落写入 runs（字符串或图片对象）"""
    if isinstance(runs, str):
        runs = [runs]
    for item in runs:
        if isinstance(item, str):
            if item:
                paragraph.add_run(item)
        elif isinstance(item, dict) and 'image' in item:
            add_image_run(paragraph, item, resolver, warnings)
        else:
            raise ValueError(f'不支持的 run 类型: {item!r}')


def add_image_run(paragraph, spec, resolver, warnings):
    src = spec.get('image') or spec.get('src')
    width_in = spec.get('width')
    if width_in is None:
        width_in = DEFAULT_IMAGE_WIDTH_IN
    width_in = min(float(width_in), MAX_IMAGE_WIDTH_IN)
    path = resolver.resolve(src)
    if path is None:
        warnings.append(f'图片未找到: {src}')
        run = paragraph.add_run(f'【图片未找到：{src}】')
        run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        return
    run = paragraph.add_run()
    try:
        run.add_picture(path, width=Inches(width_in))
    except Exception as e:
        warnings.append(f'图片插入失败 {src}: {e}')
        run.text = f'【图片插入失败：{src}】'
        run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)


def build(spec_path, output_path, vault, template_path):
    with open(spec_path, 'r', encoding='utf-8') as f:
        spec = json.load(f)

    warnings = []
    resolver = ImageResolver(vault)
    doc = Document(template_path)

    # 大标题
    title = spec.get('title')
    if not title and spec.get('date'):
        title = f"{spec['date']} 周报"
    if not title:
        raise ValueError('spec.json 需提供 title 或 date 字段')
    doc.add_paragraph(title, style='周报大标题')

    h1_no = 0
    h2_no = 0
    for block in spec.get('blocks', []):
        btype = block.get('type')
        text = block.get('text')
        runs = block.get('runs')

        if btype == 'h1':
            h1_no += 1
            h2_no = 0
            doc.add_paragraph(f'{h1_no}.{text}', style='一级标题')
        elif btype == 'h2':
            h2_no += 1
            doc.add_paragraph(f'{h1_no}.{h2_no} {text}', style='二级标题')
        elif btype == 'h3':
            doc.add_paragraph(text, style='三级标题')
        elif btype == 'para':
            p = doc.add_paragraph(style='正文')
            add_runs(p, runs if runs is not None else text, resolver, warnings)
        elif btype == 'para_noindent':
            p = doc.add_paragraph(style='正文')
            add_runs(p, runs if runs is not None else text, resolver, warnings)
            set_no_indent(p)
        elif btype == 'web':
            p = doc.add_paragraph(style='Normal (Web)')
            add_runs(p, runs if runs is not None else text, resolver, warnings)
        elif btype == 'image':
            p = doc.add_paragraph(style='Normal (Web)')
            add_image_run(p, block, resolver, warnings)
        elif btype == 'placeholder':
            p = doc.add_paragraph(style='正文')
            run = p.add_run(text)
            run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
        elif btype == 'empty':
            doc.add_paragraph('', style='正文')
        else:
            raise ValueError(f'不支持的 block 类型: {btype!r}')

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)

    print(f'✅ 周报已生成: {output_path}')
    print(f'   大标题: {title}')
    print(f'   一级标题 {h1_no} 个，blocks {len(spec.get("blocks", []))} 个')
    if warnings:
        print(f'⚠️ {len(warnings)} 条警告:')
        for w in warnings:
            print(f'   - {w}')
    return warnings


def main():
    parser = argparse.ArgumentParser(description='周报 docx 生成器')
    parser.add_argument('spec', help='JSON 内容规格文件路径')
    parser.add_argument('--output', '-o', required=True, help='输出 docx 路径')
    parser.add_argument('--vault', default=DEFAULT_VAULT,
                        help=f'Obsidian 库根目录（用于按文件名查找图片），默认: {DEFAULT_VAULT}')
    parser.add_argument('--template', default=DEFAULT_TEMPLATE,
                        help='周报模板 docx 路径')
    args = parser.parse_args()
    build(args.spec, args.output, args.vault, args.template)


if __name__ == '__main__':
    main()
