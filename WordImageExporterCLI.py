import os
import sys
import argparse
from io import BytesIO

from docx import Document
from PIL import Image


def iter_inline_image_rids(doc: Document):
    # 按文中顺序导出“插入型（inline）图片”
    for shape in doc.inline_shapes:
        blip = shape._inline.graphic.graphicData.pic.blipFill.blip
        yield blip.embed


def save_blob_as_png(blob: bytes, out_path: str, target_width: int = 500, upscale: bool = False):
    with Image.open(BytesIO(blob)) as im:
        im = im.convert("RGBA")
        w, h = im.size
        if w <= 0 or h <= 0:
            raise ValueError("Invalid image size")

        need_resize = (w != target_width) and (upscale or w > target_width)
        if need_resize:
            new_w = target_width
            new_h = int(round(h * (new_w / w)))
            im = im.resize((new_w, new_h), Image.LANCZOS)

        im.save(out_path, format="PNG")


def export_images(docx_path: str, out_dir: str, target_width: int = 500, upscale: bool = False):
    os.makedirs(out_dir, exist_ok=True)

    doc = Document(docx_path)
    rels = doc.part.related_parts

    count = 0
    seen = set()
    for rId in iter_inline_image_rids(doc):
        if rId in seen:
            continue
        seen.add(rId)

        part = rels.get(rId)
        if part is None:
            continue

        count += 1
        out_path = os.path.join(out_dir, f"{count}.png")
        save_blob_as_png(part.blob, out_path, target_width=target_width, upscale=upscale)

    return count


def default_out_dir(docx_path: str):
    doc_dir = os.path.dirname(os.path.abspath(docx_path))
    return os.path.join(doc_dir, "exported_images")


def parse_args(argv):
    p = argparse.ArgumentParser(
        description="Export images from DOCX to PNG (named by appearance order). Drag-and-drop supported."
    )
    p.add_argument("docx", nargs="?", help="Path to .docx (or drag .docx onto this exe)")
    p.add_argument("-o", "--out", default=None, help="Output directory (default: docx_dir\\exported_images)")
    p.add_argument("-w", "--width", type=int, default=500, help="Target width (px), default 500")
    p.add_argument("--upscale", action="store_true", help="Allow upscaling small images to target width")
    return p.parse_args(argv)


def main():
    args = parse_args(sys.argv[1:])

    # 支持“拖拽”：拖文件到 exe 上，sys.argv[1] 就是文件路径
    if not args.docx:
        print("Usage: WordImageExporterCLI.exe <file.docx>")
        print("Tip: You can drag a .docx file onto this exe.")
        input("Press Enter to exit...")
        return

    docx_path = args.docx.strip('"')
    if not os.path.isfile(docx_path) or not docx_path.lower().endswith(".docx"):
        print(f"Not a valid .docx file: {docx_path}")
        input("Press Enter to exit...")
        return

    out_dir = args.out or default_out_dir(docx_path)

    try:
        n = export_images(docx_path, out_dir, target_width=args.width, upscale=args.upscale)
        print(f"Done. Exported {n} image(s).")
        print(f"Output: {out_dir}")
    except Exception as e:
        print("Failed:", e)

    input("Press Enter to exit...")


if __name__ == "__main__":
    main()