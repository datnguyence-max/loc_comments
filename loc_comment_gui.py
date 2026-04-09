import pdfplumber, re, threading
from pathlib import Path
from collections import OrderedDict
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

SUB_FROM = "Các giá trị này phải do ME cung cấp"
SUB_TO   = "Các giá trị tải trọng này phải do ME cung cấp"
SKIP_SUBJ = {"rectangle", "arrow", "placed image", "stamp"}
META = re.compile(r"Type:\s*(\S+)\s+Author:\s*(\S+)\s+Subject:\s*(.+?)\s+Date:")

def process(pdf_path, out_path, author, log):
    try:
        lines = []
        with pdfplumber.open(pdf_path) as pdf:
            total = len(pdf.pages)
            for idx, page in enumerate(pdf.pages, 1):
                txt = page.extract_text() or ""
                lines.extend(txt.splitlines())
                log(f"Đang đọc trang {idx}/{total}...")

        map_ = OrderedDict()
        i = 0
        while i < len(lines):
            m_page = re.match(r"^Page:\s*(\d+)$", lines[i].strip())
            if not m_page:
                i += 1
                continue
            page_num = int(m_page.group(1))
            i += 1

            meta_str = ""
            while i < len(lines) and not re.match(r"^Page:\s*\d+$", lines[i].strip()):
                meta_str += " " + lines[i].strip()
                i += 1
                if "Status:" in meta_str:
                    break

            m_meta = META.search(meta_str)
            if not m_meta:
                continue

            ann_type   = m_meta.group(1).strip().lower()
            ann_author = m_meta.group(2).strip()
            ann_subj   = m_meta.group(3).strip().lower()

            if ann_author != author:
                continue
            if ann_subj in SKIP_SUBJ or ann_type in SKIP_SUBJ:
                continue
            if ann_type != "freetext":
                continue

            content_lines = []
            while i < len(lines):
                ln = lines[i].strip()
                if re.match(r"^Page:\s*\d+$", ln):
                    break
                if ln and ln != "<None>":
                    content_lines.append(ln)
                i += 1

            content = " ".join(content_lines).strip()
            content = content.replace(SUB_FROM, SUB_TO)
            if not content:
                continue

            if content not in map_:
                map_[content] = []
            if page_num not in map_[content]:
                map_[content].append(page_num)

        out_lines = []
        for content, pages in map_.items():
            pg = ", ".join(str(p) for p in sorted(pages))
            out_lines.append(f"CV\t{pg}\t{content}")

        Path(out_path).write_text("\n".join(out_lines), encoding="utf-8-sig")
        log(f"\n✅ Hoàn tất! {len(out_lines)} comment → {out_path}")
        messagebox.showinfo("Xong", f"Xuất {len(out_lines)} comment\n{out_path}")

    except Exception as e:
        log(f"\n❌ Lỗi: {e}")
        messagebox.showerror("Lỗi", str(e))

# ── GUI ──────────────────────────────────────────────────────────────
root = tk.Tk()
root.title("Lọc Comment PDF")
root.resizable(False, False)
root.configure(padx=20, pady=16)

pdf_var    = tk.StringVar()
out_var    = tk.StringVar()
author_var = tk.StringVar(value="dat.nguyen7")

def browse_pdf():
    p = filedialog.askopenfilename(
        title="Chọn file PDF Summary of Comments",
        filetypes=[("PDF files", "*.pdf")]
    )
    if p:
        pdf_var.set(p)
        out_var.set(str(Path(p).with_suffix(".txt")))

def browse_out():
    p = filedialog.asksaveasfilename(
        title="Lưu file output",
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt")]
    )
    if p:
        out_var.set(p)

def run():
    pdf    = pdf_var.get().strip()
    out    = out_var.get().strip()
    author = author_var.get().strip()
    if not pdf:
        messagebox.showwarning("Thiếu", "Chưa chọn file PDF!"); return
    if not out:
        messagebox.showwarning("Thiếu", "Chưa chọn nơi lưu file!"); return
    if not author:
        messagebox.showwarning("Thiếu", "Chưa nhập Author!"); return

    log_box.config(state="normal")
    log_box.delete("1.0", tk.END)
    btn_run.config(state="disabled")

    def log(msg):
        log_box.after(0, lambda: (
            log_box.insert(tk.END, msg + "\n"),
            log_box.see(tk.END)
        ))

    def task():
        process(pdf, out, author, log)
        btn_run.after(0, lambda: btn_run.config(state="normal"))

    threading.Thread(target=task, daemon=True).start()

# ── Layout ───────────────────────────────────────────────────────────
def label(text, row, top_pad=10):
    tk.Label(root, text=text, anchor="w").grid(
        row=row, column=0, columnspan=3, sticky="w", pady=(top_pad, 2))

# Author
label("Author:", 0, top_pad=0)
tk.Entry(root, textvariable=author_var, width=30,
         font=("Consolas", 10)).grid(row=1, column=0, columnspan=3, sticky="w")

# PDF input
label("File PDF input:", 2)
tk.Entry(root, textvariable=pdf_var, width=52).grid(
    row=3, column=0, columnspan=2, sticky="ew", padx=(0, 6))
tk.Button(root, text="Mở...", width=8, command=browse_pdf).grid(row=3, column=2)

# TXT output
label("File TXT output:", 4)
tk.Entry(root, textvariable=out_var, width=52).grid(
    row=5, column=0, columnspan=2, sticky="ew", padx=(0, 6))
tk.Button(root, text="Lưu...", width=8, command=browse_out).grid(row=5, column=2)

# Run button
btn_run = tk.Button(root, text="▶  Xuất comment", command=run,
                    bg="#0078d4", fg="white", padx=14, pady=6,
                    font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2")
btn_run.grid(row=6, column=0, columnspan=3, pady=(16, 8), sticky="ew")

# Log
tk.Label(root, text="Log:", anchor="w").grid(row=7, column=0, columnspan=3, sticky="w")
log_box = scrolledtext.ScrolledText(root, width=62, height=10,
                                     font=("Consolas", 9), bg="#f5f5f5")
log_box.grid(row=8, column=0, columnspan=3, pady=(2, 0))

root.mainloop()
