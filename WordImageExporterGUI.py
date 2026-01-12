import os
from io import BytesIO
import tkinter as tk
from tkinter import filedialog, messagebox

from docx import Document
from PIL import Image


def iter_inline_image_rids(doc: Document):
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


def default_out_dir(docx_path: str) -> str:
    doc_dir = os.path.dirname(os.path.abspath(docx_path))
    return os.path.join(doc_dir, "exported_images")


def pick_docx():
    path = filedialog.askopenfilename(title="选择 Word 文件", filetypes=[("Word 文件", "*.docx")])
    if path:
        docx_var.set(path)
        out_var.set(default_out_dir(path))


def pick_outdir():
    path = filedialog.askdirectory(title="选择输出文件夹")
    if path:
        out_var.set(path)


def run_export():
    docx_path = docx_var.get().strip().strip('"')
    out_dir = out_var.get().strip().strip('"')

    if not docx_path or (not os.path.isfile(docx_path)) or (not docx_path.lower().endswith(".docx")):
        messagebox.showerror("错误", "请选择有效的 .docx 文件")
        return

    try:
        width = int(width_var.get().strip())
        if width <= 0:
            raise ValueError
    except Exception:
        messagebox.showerror("错误", "宽度必须是正整数（例如 500）")
        return

    if not out_dir:
        out_dir = default_out_dir(docx_path)
        out_var.set(out_dir)

    try:
        n = export_images(docx_path, out_dir, target_width=width, upscale=upscale_var.get())
        messagebox.showinfo("完成", f"导出完成：共导出 {n} 张图片\n输出目录：{out_dir}")
    except Exception as e:
        messagebox.showerror("失败", f"导出失败：{e}")


# ---------------- GUI ----------------
root = tk.Tk()
root.title("Word 图片导出器（按顺序导出 PNG）")
root.resizable(False, False)

docx_var = tk.StringVar(value="")
out_var = tk.StringVar(value="")
width_var = tk.StringVar(value="500")
upscale_var = tk.BooleanVar(value=False)

padx, pady = 10, 8
root.columnconfigure(1, weight=1)

tk.Label(root, text="Word 文件 (.docx)：").grid(row=0, column=0, sticky="w", padx=padx, pady=pady)
tk.Entry(root, textvariable=docx_var, width=60).grid(row=0, column=1, padx=padx, pady=pady)
tk.Button(root, text="选择...", command=pick_docx).grid(row=0, column=2, padx=padx, pady=pady)

tk.Label(root, text="输出文件夹：").grid(row=1, column=0, sticky="w", padx=padx, pady=pady)
tk.Entry(root, textvariable=out_var, width=60).grid(row=1, column=1, padx=padx, pady=pady)
tk.Button(root, text="选择...", command=pick_outdir).grid(row=1, column=2, padx=padx, pady=pady)

tk.Label(root, text="导出宽度(px)：").grid(row=2, column=0, sticky="w", padx=padx, pady=pady)
tk.Entry(root, textvariable=width_var, width=10).grid(row=2, column=1, sticky="w", padx=padx, pady=pady)

tk.Checkbutton(root, text="允许把小图放大到目标宽度", variable=upscale_var).grid(row=3, column=1, sticky="w", padx=padx, pady=pady)

tk.Button(root, text="开始导出", command=run_export, height=2, width=18).grid(row=4, column=1, pady=16)

# 自动适配窗口大小，避免显示不全
root.update_idletasks()
root.geometry(f"{root.winfo_reqwidth()}x{root.winfo_reqheight()}")

root.mainloop()