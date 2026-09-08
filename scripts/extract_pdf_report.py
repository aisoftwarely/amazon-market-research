# -*- coding: utf-8 -*-
"""
extract_pdf_report.py — 卖家精灵市场分析报告PDF(图片型)转分段图片
生成 N 张分段PNG, 之后由 agent 用视觉读取(Read工具)逐段OCR提取数据。

用法:
  python extract_pdf_report.py <报告.pdf> [--outdir _analysis/segs] [--segments 8] [--zoom 2.2]
依赖: pip install pypdfium2 或 pip install PyMuPDF (任一即可)
"""
import argparse
import os

renderer = None
try:
    import pypdfium2 as pdfium
    renderer = 'pypdfium2'
except ImportError:
    try:
        import fitz  # PyMuPDF
        renderer = 'fitz'
    except ImportError:
        print('缺少PDF库: pip install pypdfium2 (或 pip install PyMuPDF)')
        raise SystemExit(1)


def render_pdf_pages(pdf_path, zoom):
    """返回 [(PIL.Image, width, height), ...]"""
    pages = []
    if renderer == 'pypdfium2':
        pdf = pdfium.PdfDocument(pdf_path)
        for i in range(len(pdf)):
            bitmap = pdf[i].render(scale=zoom)
            img = bitmap.to_pil()
            pages.append((img, img.width, img.height))
        pdf.close()
    else:
        doc = fitz.open(pdf_path)
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            img = __import__('PIL.Image', fromlist=['Image']).frombytes('RGB', (pix.width, pix.height), pix.samples)
            pages.append((img, pix.width, pix.height))
        doc.close()
    return pages


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pdf')
    ap.add_argument('--outdir', default='_analysis/segs')
    ap.add_argument('--segments', type=int, default=8)
    ap.add_argument('--zoom', type=float, default=2.2)
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)
    pages = render_pdf_pages(args.pdf, args.zoom)
    print(f'PDF: {args.pdf} | {len(pages)}页 | 渲染库: {renderer}')

    seg_paths = []
    for pno, (img, W, H) in enumerate(pages):
        step = H // args.segments
        for i in range(args.segments):
            top = i * step
            bottom = (i + 1) * step if i < args.segments - 1 else H
            path = os.path.join(args.outdir, f'p{pno + 1}_seg{i + 1}.png')
            img.crop((0, top, W, bottom)).save(path)
            seg_paths.append(path)
            print(f'saved {path} ({W}x{bottom - top})')

    print(f'\n共{len(seg_paths)}张分段图。请用 Read 工具逐张读取提取市场数据:')
    print('(容量/ASP/趋势/季节性/上架分布/价格分布等)')


if __name__ == '__main__':
    main()
