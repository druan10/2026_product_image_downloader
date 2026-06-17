"""
Image URL Formatter to Dropbox Share URL Converter
Reads a CSV/Excel with id + image_1, image_2, ... columns,
downloads each image, converts to JPEG, uploads to Dropbox,
and exports an XLSX with id and Dropbox share URLs formatted as
["url1","url2"]

GUI mode:  python app.py
CLI mode:  python app.py --input data.xlsx --output result.xlsx
"""

import sys
import os
import argparse
import json
import requests
from datetime import datetime
from pathlib import Path
from time import sleep

import pandas as pd
from PIL import Image
import dropbox
from dotenv import load_dotenv

_base = Path(sys._MEIPASS) if getattr(sys, "frozen", False) else Path(__file__).parent
load_dotenv(_base / ".env")

LOG_DIR = _base / "logs"
TEMP_RETENTION_DAYS = int(os.getenv("TEMP_RETENTION_DAYS", "7"))


def _cleanup_old_files(directory: Path, pattern: str = "**/*"):
    if not directory.exists():
        return
    cutoff = datetime.now().timestamp() - TEMP_RETENTION_DAYS * 86400
    for f in directory.glob(pattern):
        if f.is_file() and f.stat().st_mtime < cutoff:
            f.unlink(missing_ok=True)
    for d in sorted(directory.glob("**/"), reverse=True):
        if d != directory and d.is_dir():
            try:
                d.rmdir()
            except OSError:
                pass


def _init_log_file() -> tuple[Path, object]:
    LOG_DIR.mkdir(exist_ok=True)
    _cleanup_old_files(LOG_DIR, "run_*.log")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = LOG_DIR / f"run_{ts}.log"
    return log_path, log_path.open("w", encoding="utf-8")

# ---------------------------------------------------------------------------
# Dropbox helpers
# ---------------------------------------------------------------------------

_dbx = dropbox.Dropbox(os.getenv("DROPBOX_ACCESS_TOKEN"))

def _dbx_upload(local_path: str, dbx_path: str, overwrite=True):
    with open(local_path, "rb") as f:
        data = f.read()
    mode = dropbox.files.WriteMode.overwrite if overwrite else dropbox.files.WriteMode.add
    try:
        res = _dbx.files_upload(data, dbx_path, mode=mode)
    except dropbox.exceptions.ApiError as err:
        print("*** API error", err)
        return None
    print("uploaded as", res.name.encode("utf8"))
    return res


def _dbx_get_share_link(fpath: str) -> str | None:
    try:
        res = _dbx.sharing_create_shared_link_with_settings(fpath)
        url = res.url
    except dropbox.exceptions.ApiError as err:
        if type(err.error) == dropbox.sharing.CreateSharedLinkWithSettingsError:
            res = _dbx.sharing_list_shared_links(fpath)
            url = res.links[0].url
        else:
            print("*** API error", err)
            return None
    return str(url).replace("www.dropbox.com", "www.dl.dropboxusercontent.com")

DBX_IMAGE_ROOT = os.getenv("DROPBOX_IMAGE_FOLDER", "/dbx_script_uploads/")
IMAGES_DIR = Path(os.getenv("TEMP", os.getenv("TMP", "/tmp"))) / "image_downloader"


# ---------------------------------------------------------------------------
# Core logic
# ---------------------------------------------------------------------------

class ProcessingAborted(Exception):
    pass

def load_input_file(filepath: str) -> pd.DataFrame:
    ext = Path(filepath).suffix.lower()
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(filepath, dtype=str)
    elif ext == ".csv":
        return pd.read_csv(filepath, dtype=str)
    else:
        raise ValueError(f"Unsupported file type '{ext}'. Use .csv, .xls, or .xlsx.")


def default_output_path(input_path: str) -> str:
    p = Path(input_path)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return str(p.parent / f"{p.stem}_formatted_{ts}.xlsx")


def download_image(url: str, dest_path: Path) -> Path:
    """Download image from URL. Returns the saved file path."""
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        content_type = r.headers.get("Content-Type", "image/jpeg")
        ext = content_type.split("/")[-1].split(";")[0].strip()
        if ext in ("jpeg", "jpg"):
            ext = "jpg"
        fpath = dest_path.with_suffix(f".{ext}")
        fpath.write_bytes(r.content)
        return fpath
    except Exception:
        import wget
        saved = wget.download(url, out=str(dest_path))
        return Path(saved)


MAX_MEGAPIXELS = 4

def convert_image(src: Path, dest: Path) -> Path:
    """Convert image to white-background JPEG. Returns dest path."""
    im = Image.open(src).convert("RGBA")
    w, h = im.size
    if w * h > MAX_MEGAPIXELS * 1_000_000:
        scale = (MAX_MEGAPIXELS * 1_000_000 / (w * h)) ** 0.5
        im = im.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    background = Image.new("RGBA", im.size, "WHITE")
    background.paste(im, (0, 0), im)
    dest.parent.mkdir(parents=True, exist_ok=True)
    background.convert("RGB").save(str(dest))
    return dest


def upload_and_share(local_path: Path, dbx_path: str, log=print) -> str:
    """Upload file to Dropbox and return a direct share link."""
    log(f"  Uploading to Dropbox: {dbx_path}")
    res = _dbx_upload(str(local_path), dbx_path, overwrite=True)
    if res is None:
        raise RuntimeError(f"Dropbox upload failed for {local_path}")
    sleep(0.2)
    link = _dbx_get_share_link(res.path_display)
    if link is None:
        raise RuntimeError(f"Failed to get share link for {res.path_display}")
    return link


def process_row(row_id: str, image_urls: list[str], log=print, on_converted=None, on_error=None) -> list[str]:
    """
    For one row: download, convert, upload each image URL.
    Returns list of Dropbox share links.
    """
    orig_dir = IMAGES_DIR / "original" / row_id
    conv_dir = IMAGES_DIR / "converted" / row_id
    orig_dir.mkdir(parents=True, exist_ok=True)
    conv_dir.mkdir(parents=True, exist_ok=True)

    share_links = []

    for idx, url in enumerate(image_urls, start=1):
        label = f"{row_id}_{idx}"
        log(f"[{label}] Downloading {url}")

        try:
            orig_path = download_image(url, orig_dir / label)
            log(f"[{label}] Downloaded → {orig_path.name}")

            conv_path = conv_dir / f"{label}.jpg"
            convert_image(orig_path, conv_path)
            log(f"[{label}] Converted → {conv_path.name}")
            if on_converted:
                on_converted(label, conv_path)

            dbx_path = f"{DBX_IMAGE_ROOT}{label}.jpg"
            link = upload_and_share(conv_path, dbx_path, log=log)
            log(f"[{label}] Share link: {link}")
            share_links.append(link)

        except Exception as e:
            log(f"[{label}] ERROR: {e}")
            if on_error and not on_error(label, e):
                raise ProcessingAborted()

    return share_links


def run_process(input_path: str, output_path: str, log=print, on_converted=None, on_error=None) -> str:
    log_path, log_file = _init_log_file()
    _cleanup_old_files(IMAGES_DIR)
    _orig_log = log
    def log(msg: str):
        _orig_log(msg)
        log_file.write(msg + "\n")
        log_file.flush()

    log(f"Log: {log_path}")
    log(f"Loading: {input_path}")
    df = load_input_file(input_path)
    log(f"Found {len(df)} row(s), columns: {list(df.columns)}")

    if "id" not in df.columns:
        raise ValueError("Input file must contain an 'id' column.")

    image_cols = sorted(
        [c for c in df.columns if c.lower().startswith("image_")],
        key=lambda c: int(c.split("_")[-1]) if c.split("_")[-1].isdigit() else 0,
    )
    log(f"Image columns found: {image_cols}")

    results = []
    total = len(df)

    for i, (_, row) in enumerate(df.iterrows(), start=1):
        row_id = str(row["id"]).strip()
        log(f"\n--- Row {i}/{total}: {row_id} ---")

        urls = []
        for col in image_cols:
            val = row.get(col, "")
            if pd.notna(val):
                val = str(val).strip()
            else:
                val = ""
            if val and val.lower() != "nan":
                urls.append(val)

        if not urls:
            log(f"  No image URLs found, skipping.")
            results.append({"id": row_id, "images": "[]"})
            continue

        try:
            share_links = process_row(row_id, urls, log=log, on_converted=on_converted, on_error=on_error)
        except ProcessingAborted:
            log("\nProcessing stopped by user.")
            results.append({"id": row_id, "images": "[]"})
            break
        formatted = json.dumps(share_links, separators=(",", ":"))
        results.append({"id": row_id, "images": formatted})

    result_df = pd.DataFrame(results)
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    result_df.to_excel(output_path, index=False)
    log(f"\nSaved {len(results)} row(s) → {output_path}")
    log_file.close()
    return output_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def cli_main(args):
    output = args.output or default_output_path(args.input)
    run_process(args.input, output)


# ---------------------------------------------------------------------------
# GUI (tkinter)
# ---------------------------------------------------------------------------

def gui_main():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import threading

    root = tk.Tk()
    root.title("Image URL to Dropbox Share URL Converter")
    root.resizable(False, False)

    pad = {"padx": 10, "pady": 5}

    # --- Input file row ---
    tk.Label(root, text="Input file:").grid(row=0, column=0, sticky="e", **pad)
    input_var = tk.StringVar()
    tk.Entry(root, textvariable=input_var, width=55).grid(row=0, column=1, **pad)

    def browse_input():
        path = filedialog.askopenfilename(
            title="Select input file",
            filetypes=[("Spreadsheets", "*.csv *.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            input_var.set(path)
            if not output_var.get():
                output_var.set(default_output_path(path))

    tk.Button(root, text="Browse…", command=browse_input).grid(row=0, column=2, **pad)

    # --- Output file row ---
    tk.Label(root, text="Output XLSX:").grid(row=1, column=0, sticky="e", **pad)
    output_var = tk.StringVar()
    tk.Entry(root, textvariable=output_var, width=55).grid(row=1, column=1, **pad)

    def browse_output():
        path = filedialog.asksaveasfilename(
            title="Save output as",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=Path(default_output_path(input_var.get())).name if input_var.get() else "output.xlsx",
        )
        if path:
            output_var.set(path)

    tk.Button(root, text="Browse…", command=browse_output).grid(row=1, column=2, **pad)

    # --- Log box ---
    log_frame = tk.Frame(root)
    log_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=5, sticky="nsew")

    log_box = tk.Text(
        log_frame, height=18, width=80, state="disabled",
        bg="#1e1e1e", fg="#d4d4d4", font=("Consolas", 9), wrap="word",
    )
    scrollbar = ttk.Scrollbar(log_frame, command=log_box.yview)
    log_box.configure(yscrollcommand=scrollbar.set)
    log_box.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def log(msg: str):
        log_box.configure(state="normal")
        log_box.insert("end", msg + "\n")
        log_box.see("end")
        log_box.configure(state="disabled")
        root.update_idletasks()

    # --- Preview strip ---
    from PIL import ImageTk

    strip_frame = tk.LabelFrame(root, text="Converted Images")
    strip_frame.grid(row=3, column=0, columnspan=3, padx=10, pady=(0, 5), sticky="ew")

    strip_canvas = tk.Canvas(strip_frame, height=115, bg="#2d2d2d", highlightthickness=0)
    strip_scroll = ttk.Scrollbar(strip_frame, orient="horizontal", command=strip_canvas.xview)
    strip_canvas.configure(xscrollcommand=strip_scroll.set)
    strip_scroll.pack(side="bottom", fill="x")
    strip_canvas.pack(side="top", fill="x")

    strip_inner = tk.Frame(strip_canvas, bg="#2d2d2d")
    strip_canvas.create_window((0, 0), window=strip_inner, anchor="nw")

    _thumb_refs = []

    def add_thumbnail(label: str, conv_path: Path):
        img = Image.open(conv_path)
        img.thumbnail((90, 90))
        photo = ImageTk.PhotoImage(img)
        _thumb_refs.append(photo)

        cell = tk.Frame(strip_inner, bg="#2d2d2d", padx=4, pady=4)
        cell.pack(side="left", anchor="n")

        path_str = str(conv_path)
        btn = tk.Label(cell, image=photo, bg="#2d2d2d", cursor="hand2")
        btn.pack()
        btn.bind("<Button-1>", lambda _, p=path_str: os.startfile(p))

        tk.Label(cell, text=label, bg="#2d2d2d", fg="#aaaaaa",
                 font=("Consolas", 7), wraplength=90).pack()

        strip_inner.update_idletasks()
        strip_canvas.configure(scrollregion=strip_canvas.bbox("all"))
        strip_canvas.xview_moveto(1.0)

    # --- Progress bar ---
    progress = ttk.Progressbar(root, mode="indeterminate", length=400)
    progress.grid(row=4, column=0, columnspan=3, padx=10, pady=(0, 5))

    # --- Run button ---
    run_btn = tk.Button(
        root, text="Download, Convert & Upload", command=None,
        bg="#0066cc", fg="white", font=("", 10, "bold"), padx=20, pady=6,
    )
    run_btn.grid(row=5, column=0, columnspan=3, pady=10)

    def run():
        inp = input_var.get().strip()
        out = output_var.get().strip()
        if not inp:
            messagebox.showwarning("Missing input", "Please select an input file.")
            return
        if not out:
            out = default_output_path(inp)
            output_var.set(out)

        run_btn.configure(state="disabled")
        progress.start(10)

        def on_converted(label, conv_path):
            root.after(0, lambda lbl=label, cp=conv_path: add_thumbnail(lbl, cp))

        def on_error(label, error):
            event = threading.Event()
            result = [True]

            def ask():
                answer = messagebox.askyesno(
                    "Image Failed",
                    f"Failed to process image: {label}\n\n{error}\n\nContinue with remaining images?",
                    icon="warning",
                )
                result[0] = answer
                event.set()

            root.after(0, ask)
            event.wait()
            return result[0]

        def worker():
            try:
                run_process(inp, out, log=log, on_converted=on_converted, on_error=on_error)
                root.after(0, lambda: messagebox.showinfo("Complete", f"Output saved to:\n{out}"))
            except Exception as e:
                root.after(0, lambda: messagebox.showerror("Error", str(e)))
                log(f"ERROR: {e}")
            finally:
                root.after(0, lambda: progress.stop())
                root.after(0, lambda: run_btn.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    run_btn.configure(command=run)
    root.mainloop()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Download, convert, upload images to Dropbox")
    parser.add_argument("--input", "-i", help="Input CSV or Excel file")
    parser.add_argument("--output", "-o", help="Output XLSX path (default: auto-generated)")

    args, _ = parser.parse_known_args()

    if args.input:
        cli_main(args)
    else:
        gui_main()
