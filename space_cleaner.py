"""Space Cleaner - scan a drive/folder, list files by size, delete the ones you pick.

Run:  python space_cleaner.py
Deletes go to the Recycle Bin by default; "Delete permanently" is opt-in.
"""
import ctypes
import os
import queue
import subprocess
import threading
import time
import tkinter as tk
from ctypes import wintypes
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

UNITS = {"KB": 1024, "MB": 1024**2, "GB": 1024**3}


def fmt_size(n):
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024 or unit == "TB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024


# ---------- Recycle Bin ----------
class SHFILEOPSTRUCTW(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("wFunc", wintypes.UINT),
        ("pFrom", wintypes.LPCWSTR),
        ("pTo", wintypes.LPCWSTR),
        ("fFlags", ctypes.c_ushort),
        ("fAnyOperationsAborted", wintypes.BOOL),
        ("hNameMappings", ctypes.c_void_p),
        ("lpszProgressTitle", wintypes.LPCWSTR),
    ]


def recycle(path):
    """Move a file to the Recycle Bin. Returns True on success."""
    FO_DELETE, FOF_ALLOWUNDO, FOF_NOCONFIRMATION, FOF_SILENT, FOF_NOERRORUI = 3, 0x40, 0x10, 0x4, 0x400
    op = SHFILEOPSTRUCTW()
    op.wFunc = FO_DELETE
    op.pFrom = os.path.abspath(path) + "\0\0"
    op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI
    rc = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    return rc == 0 and not op.fAnyOperationsAborted


def drives():
    mask = ctypes.windll.kernel32.GetLogicalDrives()
    return [f"{chr(65 + i)}:\\" for i in range(26) if mask & (1 << i)]


# ---------- Scanner ----------
def scan(root, min_bytes, out_q, stop):
    """Walk root iteratively; push (path, size, mtime) for files >= min_bytes."""
    stack = [root]
    seen = 0
    last = time.time()
    while stack and not stop.is_set():
        d = stack.pop()
        try:
            with os.scandir(d) as it:
                for e in it:
                    try:
                        if e.is_symlink():
                            continue
                        if e.is_dir(follow_symlinks=False):
                            # skip junctions/reparse points to avoid loops
                            if getattr(e.stat(follow_symlinks=False), "st_file_attributes", 0) & 0x400:
                                continue
                            stack.append(e.path)
                        elif e.is_file(follow_symlinks=False):
                            st = e.stat(follow_symlinks=False)
                            seen += 1
                            if st.st_size >= min_bytes:
                                out_q.put(("file", e.path, st.st_size, st.st_mtime))
                    except OSError:
                        pass
        except OSError:
            pass
        if time.time() - last > 0.15:
            out_q.put(("progress", seen, d))
            last = time.time()
    out_q.put(("done", seen, None))


# ---------- GUI ----------
class App(tk.Tk):
    COLS = ("name", "size", "modified", "folder")

    def __init__(self):
        super().__init__()
        self.title("Space Cleaner")
        self.geometry("1100x680")
        self.minsize(800, 480)

        self.files = {}  # path -> (size, mtime)
        self.checked = set()
        self.q = queue.Queue()
        self.stop = threading.Event()
        self.scanning = False
        self.sort_col, self.sort_rev = "size", True

        self._build()
        self.after(100, self._pump)

    def _build(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="Scan:").pack(side="left")
        self.path_var = tk.StringVar(value=os.path.expanduser("~"))
        self.path_box = ttk.Combobox(top, textvariable=self.path_var, values=drives(), width=48)
        self.path_box.pack(side="left", padx=4)
        ttk.Button(top, text="Browse…", command=self.browse).pack(side="left")
        ttk.Label(top, text="Min size:").pack(side="left", padx=(12, 2))
        self.min_var = tk.StringVar(value="10")
        ttk.Entry(top, textvariable=self.min_var, width=6).pack(side="left")
        self.unit_var = tk.StringVar(value="MB")
        ttk.Combobox(top, textvariable=self.unit_var, values=list(UNITS), width=4, state="readonly").pack(side="left", padx=2)
        self.scan_btn = ttk.Button(top, text="Scan", command=self.toggle_scan)
        self.scan_btn.pack(side="left", padx=8)

        flt = ttk.Frame(self, padding=(8, 0))
        flt.pack(fill="x")
        ttk.Label(flt, text="Filter:").pack(side="left")
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(flt, textvariable=self.filter_var, width=30).pack(side="left", padx=4)
        ttk.Label(flt, text="(name, extension or folder)", foreground="gray").pack(side="left")

        mid = ttk.Frame(self, padding=8)
        mid.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(mid, columns=self.COLS, show="tree headings", selectmode="extended")
        self.tree.heading("#0", text="☐", command=self.toggle_all)
        self.tree.column("#0", width=44, stretch=False, anchor="center")
        heads = {"name": ("Name", 280), "size": ("Size", 90), "modified": ("Modified", 130), "folder": ("Folder", 520)}
        for c, (t, w) in heads.items():
            self.tree.heading(c, text=t, command=lambda c=c: self.sort_by(c))
            self.tree.column(c, width=w, anchor="e" if c == "size" else "w")
        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Button-1>", self.on_click)
        self.tree.bind("<space>", self.on_space)
        self.tree.bind("<Double-1>", self.on_double)
        self.tree.bind("<Button-3>", self.on_right)

        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Show in Explorer", command=self.show_in_explorer)
        self.menu.add_command(label="Open file", command=self.open_file)
        self.menu.add_command(label="Check selected", command=lambda: self.set_selected(True))
        self.menu.add_command(label="Uncheck selected", command=lambda: self.set_selected(False))

        bot = ttk.Frame(self, padding=8)
        bot.pack(fill="x")
        self.perm_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(bot, text="Delete permanently (skip Recycle Bin)", variable=self.perm_var).pack(side="left")
        self.del_btn = ttk.Button(bot, text="Delete checked", command=self.delete_checked)
        self.del_btn.pack(side="right")
        self.sel_lbl = ttk.Label(bot, text="0 checked")
        self.sel_lbl.pack(side="right", padx=12)

        self.status = tk.StringVar(value="Pick a folder or drive and click Scan.")
        ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w", padding=4).pack(fill="x", side="bottom")

    # ----- scanning -----
    def browse(self):
        d = filedialog.askdirectory(initialdir=self.path_var.get())
        if d:
            self.path_var.set(os.path.normpath(d))

    def toggle_scan(self):
        if self.scanning:
            self.stop.set()
            return
        root = self.path_var.get().strip()
        if not os.path.isdir(root):
            messagebox.showerror("Space Cleaner", "That folder doesn't exist.")
            return
        try:
            min_bytes = int(float(self.min_var.get()) * UNITS[self.unit_var.get()])
        except ValueError:
            messagebox.showerror("Space Cleaner", "Min size must be a number.")
            return
        self.files.clear()
        self.checked.clear()
        self.refresh()
        self.stop = threading.Event()
        self.scanning = True
        self.scan_btn.config(text="Stop")
        threading.Thread(target=scan, args=(root, min_bytes, self.q, self.stop), daemon=True).start()

    def _pump(self):
        added = False
        try:
            while True:
                msg = self.q.get_nowait()
                if msg[0] == "file":
                    self.files[msg[1]] = (msg[2], msg[3])
                    added = True
                elif msg[0] == "progress":
                    self.status.set(f"Scanning… {msg[1]:,} files seen, {len(self.files):,} matches — {msg[2]}")
                elif msg[0] == "done":
                    self.scanning = False
                    self.scan_btn.config(text="Scan")
                    total = sum(s for s, _ in self.files.values())
                    self.status.set(f"Done. {msg[1]:,} files scanned; {len(self.files):,} listed, {fmt_size(total)} total.")
                    added = True
        except queue.Empty:
            pass
        if added:
            self.refresh()
        self.after(250 if self.scanning else 100, self._pump)

    # ----- list display -----
    def visible(self):
        f = self.filter_var.get().strip().lower()
        items = [(p, s, m) for p, (s, m) in self.files.items() if not f or f in p.lower()]
        keys = {
            "name": lambda t: os.path.basename(t[0]).lower(),
            "size": lambda t: t[1],
            "modified": lambda t: t[2],
            "folder": lambda t: os.path.dirname(t[0]).lower(),
        }
        items.sort(key=keys[self.sort_col], reverse=self.sort_rev)
        return items

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for p, s, m in self.visible()[:20000]:
            self.tree.insert("", "end", iid=p, text="☑" if p in self.checked else "☐",
                             values=(os.path.basename(p), fmt_size(s),
                                     datetime.fromtimestamp(m).strftime("%Y-%m-%d %H:%M"),
                                     os.path.dirname(p)))
        self.update_sel()

    def sort_by(self, col):
        self.sort_rev = not self.sort_rev if col == self.sort_col else (col in ("size", "modified"))
        self.sort_col = col
        self.refresh()

    # ----- selection -----
    def update_sel(self):
        total = sum(self.files[p][0] for p in self.checked if p in self.files)
        self.sel_lbl.config(text=f"{len(self.checked)} checked — {fmt_size(total)}")

    def set_check(self, p, on):
        if on:
            self.checked.add(p)
        else:
            self.checked.discard(p)
        if self.tree.exists(p):
            self.tree.item(p, text="☑" if on else "☐")

    def on_click(self, ev):
        if self.tree.identify_region(ev.x, ev.y) == "tree" and self.tree.identify_column(ev.x) == "#0":
            row = self.tree.identify_row(ev.y)
            if row:
                self.set_check(row, row not in self.checked)
                self.update_sel()
                return "break"

    def on_space(self, _):
        for p in self.tree.selection():
            self.set_check(p, p not in self.checked)
        self.update_sel()
        return "break"

    def set_selected(self, on):
        for p in self.tree.selection():
            self.set_check(p, on)
        self.update_sel()

    def toggle_all(self):
        shown = self.tree.get_children()
        on = not all(p in self.checked for p in shown)
        for p in shown:
            self.set_check(p, on)
        self.update_sel()

    def on_right(self, ev):
        row = self.tree.identify_row(ev.y)
        if row:
            if row not in self.tree.selection():
                self.tree.selection_set(row)
            self.menu.tk_popup(ev.x_root, ev.y_root)

    def on_double(self, ev):
        if self.tree.identify_column(ev.x) != "#0":
            self.show_in_explorer()

    def show_in_explorer(self):
        sel = self.tree.selection()
        if sel:
            subprocess.Popen(["explorer", "/select,", os.path.normpath(sel[0])])

    def open_file(self):
        sel = self.tree.selection()
        if sel:
            os.startfile(sel[0])

    # ----- deleting -----
    def delete_checked(self):
        paths = [p for p in self.checked if p in self.files]
        if not paths:
            messagebox.showinfo("Space Cleaner", "Nothing is checked.")
            return
        total = sum(self.files[p][0] for p in paths)
        perm = self.perm_var.get()
        how = "PERMANENTLY delete (cannot be undone)" if perm else "move to the Recycle Bin"
        preview = "\n".join(os.path.basename(p) for p in paths[:8]) + ("\n…" if len(paths) > 8 else "")
        if not messagebox.askyesno("Confirm", f"{how.capitalize()} {len(paths)} file(s), {fmt_size(total)}?\n\n{preview}",
                                   icon="warning" if perm else "question"):
            return
        freed, failed = 0, []
        for p in paths:
            try:
                if perm:
                    os.remove(p)
                elif not recycle(p):
                    raise OSError("Recycle Bin refused")
                freed += self.files[p][0]
                del self.files[p]
                self.checked.discard(p)
            except OSError as e:
                failed.append(f"{os.path.basename(p)}: {e}")
        self.refresh()
        self.status.set(f"Removed {len(paths) - len(failed)} file(s), freed {fmt_size(freed)}."
                        + (f" {len(failed)} failed." if failed else ""))
        if failed:
            messagebox.showwarning("Some files couldn't be deleted", "\n".join(failed[:15]))


if __name__ == "__main__":
    App().mainloop()
