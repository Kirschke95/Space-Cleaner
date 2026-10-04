"""Space Cleaner - scan a drive/folder, list files by size, delete the ones you pick.

Run:  python space_cleaner.py
Deletes go to the Recycle Bin by default; "Delete permanently" is opt-in.

Two views:
  Files   - flat list of every file over the minimum size.
  Folders - tree of the folders those files live in, with the full size of each
            folder, so you can spot (and remove) e.g. a whole game folder.
"""
import ctypes
import os
import queue
import shutil
import subprocess
import threading
import time
import tkinter as tk
from collections import defaultdict
from ctypes import wintypes
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

UNITS = {"KB": 1024, "MB": 1024**2, "GB": 1024**3}

THEMES = {
    "dark": dict(bg="#1e1f22", panel="#2b2d31", field="#2b2d31", heading="#25272b", hover="#3a3d44",
                 fg="#e6e6e6", muted="#9a9ca3", border="#3f4248", box="#6b6f78", select="#2f5d9e",
                 selfg="#ffffff", accent="#4c8dff", implied="#2c4a7a", folder="#8ab4ff"),
    "light": dict(bg="#f3f3f3", panel="#ffffff", field="#ffffff", heading="#e8e8e8", hover="#dcdcdc",
                  fg="#1f1f1f", muted="#6b6b6b", border="#c8c8c8", box="#8a8a8a", select="#cce0ff",
                  selfg="#000000", accent="#2f6fdb", implied="#a9c6f5", folder="#1f56b8"),
}


def fmt_size(n):
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024 or unit == "TB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024


def bar(frac, width=14):
    frac = max(0.0, min(1.0, frac))
    pct = f"{frac * 100:.0f}%".rjust(4, " ")  # figure space = digit width, so bars line up
    return pct + "   " + "█" * round(frac * width)


def under(p, d):
    """True if path p is strictly inside folder d."""
    return p.startswith(d.rstrip("\\/") + os.sep)


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
    """Move a file or folder to the Recycle Bin. Returns True on success."""
    FO_DELETE, FOF_ALLOWUNDO, FOF_NOCONFIRMATION, FOF_SILENT, FOF_NOERRORUI = 3, 0x40, 0x10, 0x4, 0x400
    FOF_WANTNUKEWARNING = 0x4000  # ask before permanently deleting something too big for the Recycle Bin
    op = SHFILEOPSTRUCTW()
    op.wFunc = FO_DELETE
    op.pFrom = os.path.abspath(path) + "\0\0"
    op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI | FOF_WANTNUKEWARNING
    rc = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    return rc == 0 and not op.fAnyOperationsAborted


def drives():
    mask = ctypes.windll.kernel32.GetLogicalDrives()
    return [f"{chr(65 + i)}:\\" for i in range(26) if mask & (1 << i)]


# ---------- Workers ----------
def rollup(own, root):
    """Turn per-folder direct sizes into totals that include all subfolders."""
    total = defaultdict(int)
    for d, s in own.items():
        while True:
            total[d] += s
            parent = os.path.dirname(d)
            if d == root or parent == d:
                break
            d = parent
    return dict(total)


def scan(root, min_bytes, out_q, stop):
    """Walk root iteratively; push (path, size, mtime) for files >= min_bytes, then folder totals."""
    stack = [root]
    own = defaultdict(int)  # folder -> bytes of the files directly inside it
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
                            own[d] += st.st_size
                            if st.st_size >= min_bytes:
                                out_q.put(("file", e.path, st.st_size, st.st_mtime))
                    except OSError:
                        pass
        except OSError:
            pass
        if time.time() - last > 0.15:
            out_q.put(("progress", seen, d))
            last = time.time()
    out_q.put(("dirs", rollup(own, root)))
    out_q.put(("done", seen))


def delete_worker(jobs, permanent, out_q):
    """jobs: [(path, size, is_dir)]. Reports each result back through out_q."""
    ctypes.windll.ole32.CoInitialize(None)
    for path, size, is_dir in jobs:
        out_q.put(("deleting", path))
        try:
            if not permanent:
                if not recycle(path):
                    raise OSError("Recycle Bin refused")
            elif is_dir:
                shutil.rmtree(path)
            else:
                os.remove(path)
            out_q.put(("deleted", path, size, is_dir))
        except OSError as e:
            out_q.put(("delfail", f"{os.path.basename(path) or path}: {e}"))
    out_q.put(("deldone",))


# ---------- GUI ----------
class App(tk.Tk):
    FILE_COLS = ("name", "size", "modified", "folder")
    DIR_COLS = ("size", "share", "count")

    def __init__(self):
        super().__init__()
        self.title("Space Cleaner")
        self.geometry("1100x680")
        self.minsize(800, 480)

        self.files = {}            # path -> (size, mtime) for files >= min size
        self.dir_sizes = {}        # folder -> total bytes of everything inside (arrives when a scan ends)
        self.fm = None             # folder index of the visible files, see build_model()
        self.scan_root = None
        self.checked = set()       # checked files
        self.checked_dirs = set()  # checked folders (deleted with everything inside)
        self.q = queue.Queue()
        self.stop = threading.Event()
        self.scanning = self.deleting = False
        self.del_stats = None
        self.stale = set()         # views to rebuild when next shown
        self.pending, self.last_refresh = False, 0.0
        self.sort_col, self.sort_rev = "size", True
        self.dark = True

        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.img_off, self.img_on, self.img_implied = (tk.PhotoImage(width=16, height=16) for _ in range(3))
        self._build()
        self.apply_theme()
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
        unit_box = ttk.Combobox(top, textvariable=self.unit_var, values=list(UNITS), width=4, state="readonly")
        unit_box.pack(side="left", padx=2)
        self.combos = (self.path_box, unit_box)
        self.scan_btn = ttk.Button(top, text="Scan", command=self.toggle_scan)
        self.scan_btn.pack(side="left", padx=8)
        self.theme_btn = ttk.Button(top, style="Icon.TButton", width=3, command=self.toggle_theme)
        self.theme_btn.pack(side="right")

        flt = ttk.Frame(self, padding=(8, 0))
        flt.pack(fill="x")
        ttk.Label(flt, text="Filter:").pack(side="left")
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(flt, textvariable=self.filter_var, width=30).pack(side="left", padx=4)
        self.hint = ttk.Label(flt, style="Muted.TLabel")
        self.hint.pack(side="left")
        self.view_var = tk.StringVar(value="files")
        for val, txt in (("dirs", "Folders"), ("files", "Files")):  # packed right-to-left
            ttk.Radiobutton(flt, text=txt, value=val, variable=self.view_var, style="Toolbutton",
                            command=self.switch_view).pack(side="right")
        ttk.Label(flt, text="View:").pack(side="right", padx=(0, 6))

        self.mid = ttk.Frame(self, padding=8)
        self.mid.pack(fill="both", expand=True)

        self.ftree = self._make_tree(self.FILE_COLS)
        self.ftree.heading("#0", text="☐", command=self.toggle_all)
        self.ftree.column("#0", width=48, stretch=False, anchor="center")
        heads = {"name": ("Name", 280), "size": ("Size", 90), "modified": ("Modified", 130), "folder": ("Folder", 520)}
        for c, (t, w) in heads.items():
            self.ftree.heading(c, text=t, command=lambda c=c: self.sort_by(c))
            self.ftree.column(c, width=w, anchor="e" if c == "size" else "w")

        self.dtree = self._make_tree(self.DIR_COLS)
        self.dtree.heading("#0", text="Folder / file", anchor="w")
        self.dtree.column("#0", width=560)
        heads = {"size": ("Size", 100), "share": ("Share of parent folder", 210), "count": ("Large files", 90)}
        for c, (t, w) in heads.items():
            self.dtree.heading(c, text=t, anchor="w" if c == "share" else "e")
            self.dtree.column(c, width=w, stretch=False, anchor="w" if c == "share" else "e")
        self.dtree.bind("<<TreeviewOpen>>", lambda _: self._populate(self.dtree.focus()))
        self.ftree.master.pack(fill="both", expand=True)

        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Show in Explorer", command=self.show_in_explorer)
        self.menu.add_command(label="Open", command=self.open_file)
        self.menu.add_command(label="Reveal in Folders view", command=self.reveal)
        self.menu.add_separator()
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
        ttk.Label(self, textvariable=self.status, style="Status.TLabel", anchor="w").pack(fill="x", side="bottom")
        self.switch_view()

    def _make_tree(self, cols):
        frame = ttk.Frame(self.mid)
        tree = ttk.Treeview(frame, columns=cols, show="tree headings", selectmode="extended")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        tree.bind("<Button-1>", self.on_click)
        tree.bind("<space>", self.on_space)
        tree.bind("<Double-1>", self.on_double)
        tree.bind("<Button-3>", self.on_right)
        return tree

    # ----- theme -----
    def toggle_theme(self):
        self.dark = not self.dark
        self.apply_theme()

    def apply_theme(self):
        c = THEMES["dark" if self.dark else "light"]
        st = self.style
        self.configure(background=c["bg"])
        st.configure(".", background=c["bg"], foreground=c["fg"], fieldbackground=c["field"],
                     bordercolor=c["border"], lightcolor=c["border"], darkcolor=c["border"],
                     troughcolor=c["bg"], selectbackground=c["select"], selectforeground=c["selfg"],
                     insertcolor=c["fg"], focuscolor=c["accent"], arrowcolor=c["fg"])
        st.map(".", foreground=[("disabled", c["muted"])])
        st.configure("TButton", background=c["panel"], padding=(10, 3))
        st.map("TButton", background=[("pressed", c["select"]), ("active", c["hover"])])
        st.configure("Icon.TButton", font=("Segoe UI Symbol", 13), padding=(2, 0), background=c["bg"],
                     bordercolor=c["bg"], lightcolor=c["bg"], darkcolor=c["bg"])
        st.map("Icon.TButton", background=[("active", c["hover"])])
        st.configure("Toolbutton", background=c["panel"], padding=(12, 3), anchor="center")
        st.map("Toolbutton", background=[("selected", c["accent"]), ("active", c["hover"])],
               foreground=[("selected", "#ffffff")])
        st.configure("TCombobox", background=c["panel"], foreground=c["fg"])
        st.map("TCombobox", fieldbackground=[("readonly", c["field"])], foreground=[("readonly", c["fg"])],
               selectbackground=[("readonly", c["field"])], selectforeground=[("readonly", c["fg"])],
               background=[("active", c["hover"])])
        st.configure("TCheckbutton", indicatorbackground=c["field"], indicatorforeground=c["fg"])
        st.map("TCheckbutton", background=[("active", c["bg"])],
               indicatorbackground=[("pressed", c["hover"]), ("selected", c["field"])])
        st.configure("Treeview", background=c["panel"], fieldbackground=c["panel"], foreground=c["fg"],
                     rowheight=22, borderwidth=0)
        st.map("Treeview", background=[("selected", c["select"])], foreground=[("selected", c["selfg"])])
        st.configure("Treeview.Heading", background=c["heading"], foreground=c["fg"], relief="flat", padding=4)
        st.map("Treeview.Heading", background=[("active", c["hover"])])
        st.configure("TScrollbar", background=c["panel"], troughcolor=c["bg"], bordercolor=c["bg"])
        st.map("TScrollbar", background=[("active", c["hover"])])
        st.configure("Muted.TLabel", foreground=c["muted"])
        st.configure("Status.TLabel", background=c["heading"], foreground=c["muted"], padding=(8, 4))

        # combobox dropdown lists are plain Tk listboxes outside the ttk style system
        for cb in self.combos:
            try:
                pd = self.tk.call("ttk::combobox::PopdownWindow", cb)
                self.tk.call(f"{pd}.f.l", "configure", "-background", c["panel"], "-foreground", c["fg"],
                             "-selectbackground", c["select"], "-selectforeground", c["selfg"])
            except tk.TclError:
                pass
        self.menu.configure(background=c["panel"], foreground=c["fg"], activebackground=c["select"],
                            activeforeground=c["selfg"], borderwidth=0)
        self.dtree.tag_configure("dir", foreground=c["folder"])
        self.theme_btn.config(text="☀" if self.dark else "☾")

        self._draw_box(self.img_off, c["box"], c["field"])
        self._draw_box(self.img_on, c["accent"], c["accent"], "#ffffff")
        self._draw_box(self.img_implied, c["implied"], c["implied"], "#ffffff")
        self._titlebar()

    @staticmethod
    def _draw_box(img, border, fill, tick=None):
        img.blank()
        img.put(border, to=(1, 1, 15, 15))
        img.put(fill, to=(2, 2, 14, 14))
        if tick:
            for x, y in ((4, 7), (5, 8), (6, 9), (7, 8), (8, 7), (9, 6), (10, 5), (11, 4)):
                img.put(tick, to=(x, y, x + 2, y + 2))

    def _titlebar(self):
        """Ask Windows to draw the title bar dark/light to match (Windows 10 20H1+ / 11)."""
        try:
            self.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            val = ctypes.c_int(int(self.dark))
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(val), ctypes.sizeof(val))
            ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x27)  # NOSIZE|NOMOVE|NOZORDER|FRAMECHANGED
        except Exception:
            pass

    # ----- scanning -----
    def browse(self):
        d = filedialog.askdirectory(initialdir=self.path_var.get())
        if d:
            self.path_var.set(os.path.normpath(d))

    def toggle_scan(self):
        if self.scanning:
            self.stop.set()
            return
        if self.deleting:
            return
        root = os.path.normpath(os.path.abspath(self.path_var.get().strip()))
        if not os.path.isdir(root):
            messagebox.showerror("Space Cleaner", "That folder doesn't exist.")
            return
        try:
            min_bytes = int(float(self.min_var.get()) * UNITS[self.unit_var.get()])
        except ValueError:
            messagebox.showerror("Space Cleaner", "Min size must be a number.")
            return
        self.scan_root = root
        self.files.clear()
        self.dir_sizes = {}
        self.checked.clear()
        self.checked_dirs.clear()
        self.refresh()
        self.stop = threading.Event()
        self.scanning = True
        self.scan_btn.config(text="Stop")
        threading.Thread(target=scan, args=(root, min_bytes, self.q, self.stop), daemon=True).start()

    def _pump(self):
        try:
            while True:
                msg = self.q.get_nowait()
                kind = msg[0]
                if kind == "file":
                    self.files[msg[1]] = (msg[2], msg[3])
                    self.pending = True
                elif kind == "progress":
                    self.status.set(f"Scanning… {msg[1]:,} files seen, {len(self.files):,} matches — {msg[2]}")
                elif kind == "dirs":
                    self.dir_sizes = msg[1]
                elif kind == "done":
                    self.scanning = False
                    self.scan_btn.config(text="Scan")
                    total = sum(s for s, _ in self.files.values())
                    partial = " (stopped early — folder sizes are partial)" if self.stop.is_set() else ""
                    self.status.set(f"Done. {msg[1]:,} files scanned; {len(self.files):,} listed, "
                                    f"{fmt_size(total)} total.{partial}")
                    self.pending = True
                elif kind == "deleting":
                    self.status.set(f"Deleting {msg[1]}…")
                elif kind == "deleted":
                    self._forget(msg[1], msg[2], msg[3])
                    self.del_stats[0] += 1
                    self.del_stats[1] += msg[2]
                elif kind == "delfail":
                    self.del_stats[2].append(msg[1])
                elif kind == "deldone":
                    self._finish_delete()
        except queue.Empty:
            pass
        # rebuilding the folder tree is heavier, so throttle it while a scan streams in
        if self.pending and (not self.scanning or self.view_var.get() == "files"
                             or time.time() - self.last_refresh > 1.5):
            self.pending = False
            self.refresh()
        self.after(250 if self.scanning else 100, self._pump)

    # ----- list display -----
    def active(self):
        return self.dtree if self.view_var.get() == "dirs" else self.ftree

    def switch_view(self):
        dirs = self.view_var.get() == "dirs"
        (self.ftree if dirs else self.dtree).master.pack_forget()
        self.active().master.pack(fill="both", expand=True)
        self.hint.config(text="(sizes include everything inside each folder — check a folder to remove it whole)"
                         if dirs else "(name, extension or folder — right-click a file to reveal its folder)")
        if self.view_var.get() in self.stale:
            self.refresh()

    def matching(self):
        f = self.filter_var.get().strip().lower()
        return [(p, s, m) for p, (s, m) in self.files.items() if not f or f in p.lower()]

    def visible(self):
        items = self.matching()
        keys = {
            "name": lambda t: os.path.basename(t[0]).lower(),
            "size": lambda t: t[1],
            "modified": lambda t: t[2],
            "folder": lambda t: os.path.dirname(t[0]).lower(),
        }
        items.sort(key=keys[self.sort_col], reverse=self.sort_rev)
        return items

    def refresh(self):
        view = self.view_var.get()
        if view == "dirs":
            self.refresh_dirs()
        else:
            self.refresh_files()
        self.stale = {"files", "dirs"} - {view}
        self.last_refresh = time.time()
        self.update_sel()

    def refresh_files(self):
        sel = self.ftree.selection()
        self.ftree.delete(*self.ftree.get_children())
        for p, s, m in self.visible()[:20000]:
            self.ftree.insert("", "end", iid=p, image=self.img_for(p),
                              values=(os.path.basename(p), fmt_size(s),
                                      datetime.fromtimestamp(m).strftime("%Y-%m-%d %H:%M"),
                                      os.path.dirname(p)))
        keep = [i for i in sel if self.ftree.exists(i)]
        if keep:
            self.ftree.selection_set(keep)

    def sort_by(self, col):
        self.sort_rev = not self.sort_rev if col == self.sort_col else (col in ("size", "modified"))
        self.sort_col = col
        self.refresh()

    # ----- folder tree -----
    def build_model(self):
        """Index the visible files by folder: subfolders, direct files, and large-file totals."""
        kids, direct = defaultdict(set), defaultdict(list)
        msum, mcount = defaultdict(int), defaultdict(int)
        for p, s, _ in self.matching():
            d = os.path.dirname(p)
            direct[d].append(p)
            while True:
                msum[d] += s
                mcount[d] += 1
                parent = os.path.dirname(d)
                if d == self.scan_root or parent == d:
                    break
                kids[parent].add(d)
                d = parent
        self.fm = {"kids": kids, "direct": direct, "msum": msum, "mcount": mcount}

    def dsize(self, d):
        """Full folder size once the scan has finished; until then, the large files seen so far."""
        return self.dir_sizes.get(d) or self.fm["msum"].get(d, 0)

    def folder_size(self, d):
        return self.dir_sizes.get(d) or sum(s for p, (s, _) in self.files.items() if under(p, d))

    def _collapse(self, d):
        """Skip down through pass-through folders (one subfolder, nothing else) to show e.g. 'Steam\\steamapps'."""
        fm = self.fm
        while not fm["direct"].get(d) and len(fm["kids"].get(d, ())) == 1:
            (c,) = fm["kids"][d]
            if self.dsize(c) != self.dsize(d):  # parent holds other stuff too; keep it visible
                break
            d = c
        return d

    def _insert(self, parent, p, name, parent_size):
        if p in self.files:
            size, tags, count = self.files[p][0], (), ""
        else:
            size, tags, count = self.dsize(p), ("dir",), f"{self.fm['mcount'].get(p, 0):,}"
        share = bar(size / parent_size) if parent_size else ""
        self.dtree.insert(parent, "end", iid=p, text=" " + name, image=self.img_for(p), tags=tags,
                          values=(fmt_size(size), share, count))
        if tags:
            self.dtree.insert(p, "end", iid="?" + p)  # placeholder so it can be expanded; filled lazily

    def _populate(self, d):
        t = self.dtree
        if not d or d.startswith("?"):
            return
        kids = t.get_children(d)
        if len(kids) != 1 or not kids[0].startswith("?"):
            return
        t.delete(kids[0])
        rows = [(self.dsize(c), c) for c in map(self._collapse, self.fm["kids"].get(d, ()))]
        rows += [(self.files[p][0], p) for p in self.fm["direct"].get(d, ())]
        psize = self.dsize(d)
        for _, p in sorted(rows, reverse=True):
            self._insert(d, p, os.path.relpath(p, d), psize)

    def _walk(self, t, item=""):
        for c in t.get_children(item):
            if not c.startswith("?"):
                yield c
                yield from self._walk(t, c)

    def refresh_dirs(self):
        t = self.dtree
        opened = [i for i in self._walk(t) if t.item(i, "open")]
        sel, yview = t.selection(), t.yview()[0]
        t.delete(*t.get_children())
        if not self.scan_root:
            return
        self.build_model()
        if not self.fm["mcount"]:
            return
        top = self._collapse(self.scan_root)
        self._insert("", top, top, 0)
        for i in [top] + opened:
            if t.exists(i):
                self._populate(i)
                t.item(i, open=True)
        keep = [i for i in sel if t.exists(i)]
        if keep:
            t.selection_set(keep)
        t.yview_moveto(yview)

    def reveal(self):
        sel = self.ftree.selection()
        if not sel:
            return
        p = sel[0]
        self.view_var.set("dirs")
        self.switch_view()
        t = self.dtree
        cur = next(iter(t.get_children()), None)
        while cur and cur != p:
            self._populate(cur)
            t.item(cur, open=True)
            cur = next((c for c in t.get_children(cur) if c == p or under(p, c)), None)
        if cur:
            t.selection_set(cur)
            t.focus(cur)
            t.see(cur)
            t.focus_set()

    # ----- selection -----
    def img_for(self, p):
        if p in self.checked or p in self.checked_dirs:
            return self.img_on
        if any(under(p, d) for d in self.checked_dirs):
            return self.img_implied
        return self.img_off

    def _paint(self, p):
        for t in (self.ftree, self.dtree):
            if t.exists(p):
                t.item(p, image=self.img_for(p))

    def effective(self):
        """What a delete would actually remove: top-most checked folders, plus checked files outside them."""
        dirs = [d for d in self.checked_dirs if not any(under(d, o) for o in self.checked_dirs)]
        files = [p for p in self.checked if p in self.files and not any(under(p, d) for d in dirs)]
        return dirs, files

    def update_sel(self):
        dirs, files = self.effective()
        total = sum(map(self.folder_size, dirs)) + sum(self.files[p][0] for p in files)
        what = f"{len(dirs)} folder(s) + {len(files)} file(s)" if dirs else f"{len(files)}"
        self.sel_lbl.config(text=f"{what} checked — {fmt_size(total)}")

    def set_check(self, p, on):
        if p.startswith("?"):
            return
        if p in self.files:
            if on and any(under(p, d) for d in self.checked_dirs):
                self.status.set("That file is already included — its folder is checked.")
                return
            (self.checked.add if on else self.checked.discard)(p)
            self._paint(p)
            return
        if on and (p == self.scan_root or os.path.dirname(p) == p):
            self.status.set("The scanned folder itself can't be checked — pick a folder inside it.")
            return
        (self.checked_dirs.add if on else self.checked_dirs.discard)(p)
        self._paint(p)
        for i in self.ftree.get_children():
            if under(i, p):
                self.ftree.item(i, image=self.img_for(i))
        if self.dtree.exists(p):
            for i in self._walk(self.dtree, p):
                self.dtree.item(i, image=self.img_for(i))

    def toggle(self, p):
        self.set_check(p, p not in self.checked and p not in self.checked_dirs)

    def _on_checkbox(self, t, ev):
        if t.identify_region(ev.x, ev.y) != "tree":
            return False
        if t is self.ftree:
            return t.identify_column(ev.x) == "#0"
        return "image" in t.identify_element(ev.x, ev.y)

    def on_click(self, ev):
        t = ev.widget
        row = t.identify_row(ev.y)
        if row and self._on_checkbox(t, ev):
            self.toggle(row)
            self.update_sel()
            return "break"

    def on_space(self, ev):
        for p in ev.widget.selection():
            self.toggle(p)
        self.update_sel()
        return "break"

    def set_selected(self, on):
        for p in self.active().selection():
            self.set_check(p, on)
        self.update_sel()

    def toggle_all(self):
        shown = self.ftree.get_children()
        on = not all(p in self.checked for p in shown)
        for p in shown:
            self.set_check(p, on)
        self.update_sel()

    def on_right(self, ev):
        t = ev.widget
        row = t.identify_row(ev.y)
        if row and not row.startswith("?"):
            if row not in t.selection():
                t.selection_set(row)
            self.menu.entryconfig("Reveal in Folders view", state="normal" if t is self.ftree else "disabled")
            self.menu.tk_popup(ev.x_root, ev.y_root)

    def on_double(self, ev):
        t = ev.widget
        # double-clicking a folder in the tree just expands it (default behaviour)
        if t.identify_row(ev.y) in self.files and not self._on_checkbox(t, ev):
            self.show_in_explorer()

    def show_in_explorer(self):
        sel = self.active().selection()
        if sel:
            subprocess.Popen(["explorer", "/select,", os.path.normpath(sel[0])])

    def open_file(self):
        sel = self.active().selection()
        if sel:
            os.startfile(sel[0])

    # ----- deleting -----
    def delete_checked(self):
        if self.deleting or self.scanning:
            return
        dirs, files = self.effective()
        if not dirs and not files:
            messagebox.showinfo("Space Cleaner", "Nothing is checked.")
            return
        jobs = [(d, self.folder_size(d), True) for d in dirs] + [(p, self.files[p][0], False) for p in files]
        total = sum(s for _, s, _ in jobs)
        perm = self.perm_var.get()
        how = "PERMANENTLY delete (cannot be undone)" if perm else "move to the Recycle Bin"
        what = " and ".join(x for x in (f"{len(dirs)} folder(s)" if dirs else "",
                                        f"{len(files)} file(s)" if files else "") if x)
        lines = [f"[folder] {d}" for d in dirs[:8]] + [os.path.basename(p) for p in files[:max(0, 8 - len(dirs))]]
        more = len(jobs) - len(lines)
        preview = "\n".join(lines) + (f"\n… and {more} more" if more > 0 else "")
        note = "\n\nChecked folders are removed with everything inside them." if dirs else ""
        if not messagebox.askyesno("Confirm", f"{how.capitalize()} {what}, {fmt_size(total)}?{note}\n\n{preview}",
                                   icon="warning" if perm or dirs else "question"):
            return
        self.deleting = True
        self.del_btn.config(state="disabled")
        self.scan_btn.config(state="disabled")
        self.del_stats = [0, 0, []]  # removed, bytes freed, failures
        threading.Thread(target=delete_worker, args=(jobs, perm, self.q), daemon=True).start()

    def _forget(self, p, size, is_dir):
        """Drop a deleted file/folder from all bookkeeping."""
        if is_dir:
            self.files = {f: v for f, v in self.files.items() if not under(f, p)}
            self.checked = {f for f in self.checked if not under(f, p)}
            self.checked_dirs = {d for d in self.checked_dirs if d != p and not under(d, p)}
            self.dir_sizes = {d: v for d, v in self.dir_sizes.items() if d != p and not under(d, p)}
        else:
            self.files.pop(p, None)
            self.checked.discard(p)
        d = os.path.dirname(p)
        while d in self.dir_sizes:
            self.dir_sizes[d] -= size
            if d == self.scan_root or os.path.dirname(d) == d:
                break
            d = os.path.dirname(d)

    def _finish_delete(self):
        removed, freed, failed = self.del_stats
        self.deleting = False
        self.del_btn.config(state="normal")
        self.scan_btn.config(state="normal")
        self.refresh()
        self.status.set(f"Removed {removed} item(s), freed {fmt_size(freed)}."
                        + (f" {len(failed)} failed." if failed else ""))
        if failed:
            messagebox.showwarning("Some items couldn't be deleted", "\n".join(failed[:15]))


if __name__ == "__main__":
    App().mainloop()
