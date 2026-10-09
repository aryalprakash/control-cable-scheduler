"""Control Cable Scheduler - standalone Windows desktop app (Python + Tkinter).

Flow: Start (New / Open project) -> Substation + voltage levels + bays ->
select a bay -> upload its drawings -> extract (sheet title + external TB / ferrule)
-> review -> generate cable schedule -> export to Excel -> pick another bay.
"""
import json
import os
import re
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

APP = "Control Cable Scheduler"
VOLTAGES = ["220", "132", "33"]
FILE_TYPES = [("Cable Scheduler project", "*.ccs"), ("JSON", "*.json")]

# --------------------------------------------------------------------------
# Extraction settings - tune these to your drawing standard
# --------------------------------------------------------------------------
# TB tag, optionally with terminal number: TB1, TB-1, X1, XT2, TB1:5, TB1-5
TB_RE = re.compile(r"^((?:TB|XT|TS|X)[-_]?\d+[A-Z]?)(?:[:./-](\d+[A-Z]?))?$", re.I)
# Ferrule / wire number: 101, 101A, R101, YA12, 1R-12
FERRULE_RE = re.compile(r"^(?:\d{2,4}[A-Z]?|[A-Z]{1,3}[-]?\d{1,4}[A-Z]?|\d{1,2}[A-Z]{1,2}[-]?\d{1,3})$", re.I)
# Words on a terminal row which indicate wiring leaves the panel
EXTERNAL_WORDS = {"TO", "EXT", "EXTERNAL", "FIELD", "MB", "MK", "CT", "VT", "CVT", "PT",
                  "RTU", "SCADA", "BCU", "PANEL", "CABLE", "MARSHALLING", "MARSH", "YARD"}
# Words which indicate internal wiring (never external)
INTERNAL_WORDS = {"INT", "INTERNAL", "JUMPER", "LINK", "SPARE", "BRIDGE"}
TITLE_LABELS = re.compile(r"(DRAWING\s*TITLE|DRG\.?\s*TITLE|SHEET\s*TITLE|TITLE)", re.I)
STOP_WORDS = re.compile(r"\b(SCALE|DRG|DWG|SHEET|REV|DATE|CHECKED|DRAWN|APPROVED|NO\.?|SIGN)\b", re.I)


def load_pages(path):
    """Return [(page_no, width, height, [(x0, y0, x1, y1, text), ...])]."""
    ext = os.path.splitext(path)[1].lower()
    pages = []
    if ext == ".pdf":
        import fitz  # PyMuPDF
        for i, p in enumerate(fitz.open(path)):
            words = [(w[0], w[1], w[2], w[3], w[4]) for w in p.get_text("words")]
            pages.append((i + 1, p.rect.width, p.rect.height, words))
    elif ext == ".dxf":
        import ezdxf
        msp = ezdxf.readfile(path).modelspace()
        words = []
        for e in msp.query("TEXT MTEXT"):
            text = e.dxf.text if e.dxftype() == "TEXT" else e.plain_text()
            x, y = e.dxf.insert.x, -e.dxf.insert.y  # flip so "down" is +y like PDF
            h = getattr(e.dxf, "height", 2.5) or 2.5
            for part in text.replace("\\P", " ").split():
                words.append((x, y, x + h * 0.6 * len(part), y + h, part))
                x += h * 0.6 * (len(part) + 1)
        if words:
            xs = [w[0] for w in words] + [w[2] for w in words]
            ys = [w[1] for w in words] + [w[3] for w in words]
            minx, miny = min(xs), min(ys)
            words = [(a - minx, b - miny, c - minx, d - miny, t) for a, b, c, d, t in words]
            pages.append((1, max(xs) - minx, max(ys) - miny, words))
    else:
        raise ValueError(f"Unsupported file type: {ext}")
    return pages


def group_lines(words, tol):
    """Cluster words into text lines (top-to-bottom, left-to-right)."""
    lines = []
    for w in sorted(words, key=lambda w: ((w[1] + w[3]) / 2, w[0])):
        yc = (w[1] + w[3]) / 2
        if lines and abs(lines[-1]["y"] - yc) <= tol:
            lines[-1]["w"].append(w)
        else:
            lines.append({"y": yc, "w": [w]})
    for ln in lines:
        ln["w"].sort(key=lambda w: w[0])
        ln["tokens"] = [w[4].strip() for w in ln["w"] if w[4].strip()]
        ln["text"] = " ".join(ln["tokens"])
    return lines


def extract_title(words, width, height, tol):
    """Sheet title from the title block (bottom-right region of the sheet)."""
    region = [w for w in words if (w[0] + w[2]) / 2 > width * 0.45 and (w[1] + w[3]) / 2 > height * 0.70]
    lines = group_lines(region, tol)
    for i, ln in enumerate(lines):
        m = TITLE_LABELS.search(ln["text"])
        if m:
            parts = []
            rest = re.sub(r"^[\s:.\-]+", "", ln["text"][m.end():])
            if len(rest) > 3 and not STOP_WORDS.search(rest):
                parts.append(rest)
            for nxt in lines[i + 1:i + 4]:
                if STOP_WORDS.search(nxt["text"]) or len(nxt["text"]) < 4:
                    break
                parts.append(nxt["text"])
                if len(parts) >= 2:
                    break
            if parts:
                return " ".join(parts).strip()
    # fallback: longest mostly-uppercase line in the region without admin words
    cand = [ln["text"] for ln in lines if not STOP_WORDS.search(ln["text"]) and len(ln["text"]) > 8]
    return max(cand, key=len) if cand else "UNTITLED"


def extract_terminals(words, tol):
    """Find TB No / Ferrule No rows; flag the ones that look externally wired."""
    out, seen, tb_ctx = [], set(), None
    for ln in group_lines(words, tol):
        toks = ln["tokens"]
        upper = {t.upper().strip(",.:;") for t in toks}
        tb, term = None, None
        for t in toks:
            m = TB_RE.match(t)
            if m:
                tb, term = m.group(1).upper(), m.group(2)
                tb_ctx = tb
                break
        if tb is None and tb_ctx is None:
            continue
        tb = tb or tb_ctx
        ferrules = [t for t in toks if FERRULE_RE.match(t) and not TB_RE.match(t) and t != term]
        if term is None and not ferrules:
            continue
        ferrule = "/".join(dict.fromkeys(ferrules))
        key = (tb, term, ferrule)
        if key in seen:
            continue
        seen.add(key)
        ext = bool(upper & EXTERNAL_WORDS) and not (upper & INTERNAL_WORDS)
        out.append({"tb": tb, "terminal": term or "", "ferrule": ferrule,
                    "desc": ln["text"][:80], "ext": ext})
    return out


def extract_drawing(path):
    result = []
    for page_no, w, h, words in load_pages(path):
        tol = max(h * 0.004, 1.0)
        title = extract_title(words, w, h, tol)
        for t in extract_terminals(words, tol):
            t.update({"file": os.path.basename(path), "page": page_no, "title": title})
            result.append(t)
        if not any(r["page"] == page_no and r["file"] == os.path.basename(path) for r in result):
            result.append({"file": os.path.basename(path), "page": page_no, "title": title,
                           "tb": "", "terminal": "", "ferrule": "", "desc": "(no TB found)", "ext": False})
    return result


# --------------------------------------------------------------------------
# Cable schedule generation
# --------------------------------------------------------------------------
def generate_cables(kv, bay, terminals):
    groups = {}
    for t in terminals:
        if t.get("ext") and t["tb"]:
            groups.setdefault((t["title"], t["tb"]), []).append(t)
    code = re.sub(r"[^A-Z0-9]", "", bay.upper())[:8] or "BAY"
    cables = []
    for n, ((title, tb), items) in enumerate(groups.items(), 1):
        ferr = [f"{i['terminal']}:{i['ferrule']}".strip(":") for i in items]
        cables.append({"cable_no": f"{kv}-{code}-{n:02d}", "title": title, "tb": tb,
                       "ferrules": ", ".join(ferr), "cores": len(items),
                       "to": "", "type": "", "size": ""})
    return cables


def export_excel(path, project, bays):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    wb = Workbook()
    wb.remove(wb.active)
    thin = Side(style="thin", color="999999")
    head = ["S.No", "Cable No", "Sheet Title", "TB No", "Terminal : Ferrule No",
            "No. of Cores", "To (Destination)", "Cable Type", "Size (sq.mm)"]
    for kv, bay, data in bays:
        ws = wb.create_sheet(re.sub(r"[\\/*?:\[\]]", "", f"{kv}kV {bay}")[:31])
        ws["A1"] = f"CONTROL CABLE SCHEDULE - {project.get('substation', '')}"
        ws["A2"] = f"{kv} kV  |  Bay: {bay}  |  Project: {project['name']}"
        ws["A1"].font = Font(bold=True, size=13)
        ws.append([])
        ws.append(head)
        for c in ws[4]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="DDE6F0")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for i, c in enumerate(data["cables"], 1):
            ws.append([i, c["cable_no"], c["title"], c["tb"], c["ferrules"], c["cores"],
                       c["to"], c["type"], c["size"]])
        for row in ws.iter_rows(min_row=4, max_row=ws.max_row):
            for c in row:
                c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
                if c.row > 4:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
        for col, w in zip("ABCDEFGHI", (6, 18, 40, 10, 34, 10, 24, 16, 12)):
            ws.column_dimensions[col].width = w
    if not wb.sheetnames:
        wb.create_sheet("Empty")
    wb.save(path)


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------
def new_bay():
    return {"drawings": [], "terminals": [], "cables": [], "status": "Pending"}


class StartDialog(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title(APP)
        self.result = None
        self.resizable(False, False)
        f = ttk.Frame(self, padding=24)
        f.pack()
        ttk.Label(f, text=APP, font=("Segoe UI", 16, "bold")).pack(pady=(0, 4))
        ttk.Label(f, text="What would you like to do?").pack(pady=(0, 14))
        ttk.Button(f, text="Create a new project", width=30, command=self.new).pack(pady=4, ipady=4)
        ttk.Button(f, text="Work on an existing project", width=30, command=self.open).pack(pady=4, ipady=4)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        # transient() to a withdrawn root hides this dialog on macOS
        if master.winfo_viewable():
            self.transient(master)
        self.update_idletasks()
        self.deiconify()
        self.lift()
        self.focus_force()
        self.attributes("-topmost", True)
        self.after(300, lambda: self.attributes("-topmost", False))
        self.grab_set()

    def new(self):
        name = simpledialog.askstring("New project", "Project name:", parent=self)
        if name and name.strip():
            self.result = ("new", name.strip())
            self.destroy()

    def open(self):
        p = filedialog.askopenfilename(parent=self, title="Open project", filetypes=FILE_TYPES)
        if p:
            self.result = ("open", p)
            self.destroy()


class VoltageDialog(tk.Toplevel):
    """Add a voltage level with number of bays and bay names."""
    def __init__(self, master, existing):
        super().__init__(master)
        self.title("Add voltage level")
        self.result = None
        self.entries = []
        top = ttk.Frame(self, padding=12)
        top.pack(fill="x")
        ttk.Label(top, text="Voltage level (kV):").grid(row=0, column=0, sticky="w")
        self.kv = ttk.Combobox(top, values=[v for v in VOLTAGES if v not in existing], width=8)
        self.kv.grid(row=0, column=1, padx=6)
        ttk.Label(top, text="No. of bays:").grid(row=0, column=2, padx=(12, 0))
        self.n = tk.Spinbox(top, from_=1, to=40, width=5)
        self.n.grid(row=0, column=3, padx=6)
        ttk.Button(top, text="Enter bay names", command=self.build).grid(row=0, column=4, padx=6)
        self.body = ttk.Frame(self, padding=(12, 0))
        self.body.pack(fill="both")
        ttk.Button(self, text="OK", command=self.ok).pack(pady=10)
        self.transient(master)
        self.grab_set()

    def build(self):
        for w in self.body.winfo_children():
            w.destroy()
        try:
            n = int(self.n.get())
        except ValueError:
            return
        self.entries = []
        for i in range(n):
            ttk.Label(self.body, text=f"Bay {i + 1}:").grid(row=i, column=0, sticky="w", pady=2)
            e = ttk.Entry(self.body, width=36)
            e.grid(row=i, column=1, pady=2, padx=6)
            self.entries.append(e)
        if self.entries:
            self.entries[0].focus()

    def ok(self):
        kv = self.kv.get().strip()
        if not kv.isdigit():
            messagebox.showerror(APP, "Select / enter a voltage level in kV.", parent=self)
            return
        names = [e.get().strip() for e in self.entries if e.get().strip()]
        if not names:
            messagebox.showerror(APP, "Enter the number of bays and click 'Enter bay names'.", parent=self)
            return
        self.result = (kv, names)
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP)
        self.geometry("1180x720")
        self.project = None
        self.path = None
        self.theme = tk.StringVar(value="system")
        self.apply_theme()
        # On macOS, dialogs attached to a withdrawn root often never appear.
        # Park the root off-screen instead of withdraw() until a project is chosen.
        if sys.platform == "darwin":
            self.geometry("1x1+-5000+-5000")
        else:
            self.withdraw()
        self.build_menu()
        self.build_ui()
        self.after(50, self.startup)

    # ---------- theme ----------
    def apply_theme(self):
        st = ttk.Style(self)
        if self.theme.get() == "system":
            st.theme_use("vista" if "vista" in st.theme_names() else "clam")
            st.configure("Treeview", rowheight=24)
            self.configure(bg=st.lookup("TFrame", "background") or "SystemButtonFace")
        else:  # soft light palette
            bg, fg, field = "#f3f4f6", "#202124", "#fafafa"
            st.theme_use("clam")
            st.configure(".", background=bg, foreground=fg, fieldbackground=field)
            st.configure("TButton", background="#e5e7eb", padding=5)
            st.map("TButton", background=[("active", "#d7dbe0")])
            st.configure("Treeview", background=field, fieldbackground=field, foreground=fg, rowheight=24)
            st.configure("Treeview.Heading", background="#e1e5ea", foreground=fg)
            st.map("Treeview", background=[("selected", "#c9ddf2")], foreground=[("selected", "#000000")])
            st.configure("TNotebook", background=bg)
            st.configure("TNotebook.Tab", background="#e5e7eb", padding=(10, 4))
            st.map("TNotebook.Tab", background=[("selected", field)])
            self.configure(bg=bg)

    # ---------- menu ----------
    def build_menu(self):
        m = tk.Menu(self)
        f = tk.Menu(m, tearoff=0)
        f.add_command(label="New project...", command=self.menu_new)
        f.add_command(label="Open project...", command=self.menu_open)
        f.add_command(label="Save", command=self.save, accelerator="Ctrl+S")
        f.add_command(label="Save as...", command=lambda: self.save(True))
        f.add_separator()
        f.add_command(label="Export all bays to Excel...", command=lambda: self.export(all_bays=True))
        f.add_separator()
        f.add_command(label="Exit", command=self.quit_app)
        m.add_cascade(label="File", menu=f)
        v = tk.Menu(m, tearoff=0)
        v.add_radiobutton(label="System default", variable=self.theme, value="system", command=self.apply_theme)
        v.add_radiobutton(label="Light", variable=self.theme, value="light", command=self.apply_theme)
        m.add_cascade(label="Appearance", menu=v)
        self.config(menu=m)
        self.bind("<Control-s>", lambda e: self.save())
        self.protocol("WM_DELETE_WINDOW", self.quit_app)

    def build_ui(self):
        pw = ttk.PanedWindow(self, orient="horizontal")
        pw.pack(fill="both", expand=True, padx=6, pady=6)
        left = ttk.Frame(pw, padding=4)
        pw.add(left, weight=1)
        self.lbl_sub = ttk.Label(left, text="", font=("Segoe UI", 10, "bold"), wraplength=260)
        self.lbl_sub.pack(anchor="w", pady=(0, 6))
        self.tree = ttk.Treeview(left, show="tree", selectmode="browse")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.on_select())
        bf = ttk.Frame(left)
        bf.pack(fill="x", pady=6)
        for i, (t, c) in enumerate([("Add voltage level", self.add_voltage), ("Add bay", self.add_bay),
                                    ("Delete selected", self.delete_selected), ("Rename substation", self.rename_sub)]):
            ttk.Button(bf, text=t, command=c).grid(row=i // 2, column=i % 2, sticky="ew", padx=2, pady=2)
        bf.columnconfigure((0, 1), weight=1)

        right = ttk.Frame(pw, padding=4)
        pw.add(right, weight=4)
        self.lbl_bay = ttk.Label(right, text="Select a bay on the left to begin.", font=("Segoe UI", 12, "bold"))
        self.lbl_bay.pack(anchor="w", pady=(0, 6))
        self.nb = ttk.Notebook(right)
        self.nb.pack(fill="both", expand=True)

        # tab 1: drawings
        t1 = ttk.Frame(self.nb, padding=8)
        self.nb.add(t1, text="1. Drawings")
        ttk.Label(t1, text="Upload only the drawings associated with the selected bay (PDF / DXF).").pack(anchor="w")
        self.lst = tk.Listbox(t1, height=10, relief="flat", highlightthickness=1)
        self.lst.pack(fill="both", expand=True, pady=6)
        r = ttk.Frame(t1)
        r.pack(fill="x")
        ttk.Button(r, text="Upload drawings...", command=self.upload).pack(side="left", padx=2)
        ttk.Button(r, text="Remove selected", command=self.remove_drawing).pack(side="left", padx=2)
        ttk.Button(r, text="Extract information", command=self.extract).pack(side="right", padx=2)

        # tab 2: extracted terminals
        t2 = ttk.Frame(self.nb, padding=8)
        self.nb.add(t2, text="2. Extracted TB / Ferrule")
        ttk.Label(t2, text="Tick 'External' only for terminals needing external wiring. "
                           "Double-click the External column to toggle; double-click other cells to edit.").pack(anchor="w")
        cols = ("title", "tb", "terminal", "ferrule", "ext", "desc")
        self.tv_t = ttk.Treeview(t2, columns=cols, show="headings", selectmode="extended")
        for c, h, w in zip(cols, ("Sheet Title", "TB No", "Terminal", "Ferrule No", "External", "Row text / source"),
                           (300, 70, 70, 110, 70, 300)):
            self.tv_t.heading(c, text=h)
            self.tv_t.column(c, width=w, anchor="w")
        self.tv_t.pack(fill="both", expand=True, pady=6)
        self.tv_t.bind("<Double-1>", lambda e: self.cell_edit(self.tv_t, e, "terminals", {"title": 0, "tb": 1, "terminal": 2, "ferrule": 3}, "ext", 4))
        r = ttk.Frame(t2)
        r.pack(fill="x")
        ttk.Button(r, text="Add row", command=self.add_terminal).pack(side="left", padx=2)
        ttk.Button(r, text="Delete selected rows", command=self.del_terminals).pack(side="left", padx=2)
        ttk.Button(r, text="Generate cable schedule", command=self.generate).pack(side="right", padx=2)

        # tab 3: cable schedule
        t3 = ttk.Frame(self.nb, padding=8)
        self.nb.add(t3, text="3. Cable Schedule")
        ttk.Label(t3, text="Double-click To / Cable Type / Size cells to fill them in.").pack(anchor="w")
        cols = ("sn", "cable_no", "title", "tb", "ferrules", "cores", "to", "type", "size")
        self.tv_c = ttk.Treeview(t3, columns=cols, show="headings")
        for c, h, w in zip(cols, ("S.No", "Cable No", "Sheet Title", "TB No", "Terminal : Ferrule", "Cores", "To", "Cable Type", "Size"),
                           (45, 120, 240, 60, 200, 50, 130, 100, 70)):
            self.tv_c.heading(c, text=h)
            self.tv_c.column(c, width=w, anchor="w")
        self.tv_c.pack(fill="both", expand=True, pady=6)
        self.tv_c.bind("<Double-1>", lambda e: self.cell_edit(self.tv_c, e, "cables", {"cable_no": 1, "to": 6, "type": 7, "size": 8}))
        r = ttk.Frame(t3)
        r.pack(fill="x")
        ttk.Button(r, text="Mark bay as completed", command=self.mark_done).pack(side="left", padx=2)
        ttk.Button(r, text="Export this bay to Excel...", command=self.export).pack(side="right", padx=2)

    # ---------- project handling ----------
    def startup(self):
        print(f"{APP}: waiting for start dialog…", flush=True)
        d = StartDialog(self)
        self.wait_window(d)
        if not d.result:
            self.destroy()
            return
        kind, val = d.result
        if kind == "new":
            self.new_project(val)
        elif not self.load(val):
            return self.startup()
        self.geometry("1180x720")
        self.deiconify()
        self.lift()
        self.focus_force()

    def new_project(self, name):
        sub = simpledialog.askstring("Substation", "Substation name:", parent=self) or ""
        self.project = {"name": name, "substation": sub.strip(), "voltages": {}}
        self.path = None
        self.refresh_tree()
        self.add_voltage()

    def menu_new(self):
        if self.confirm_discard():
            n = simpledialog.askstring("New project", "Project name:", parent=self)
            if n and n.strip():
                self.new_project(n.strip())

    def menu_open(self):
        if self.confirm_discard():
            p = filedialog.askopenfilename(title="Open project", filetypes=FILE_TYPES)
            if p:
                self.load(p)

    def confirm_discard(self):
        return messagebox.askyesno(APP, "Unsaved changes will be lost. Continue?") if self.project else True

    def load(self, path):
        try:
            with open(path, encoding="utf-8") as f:
                self.project = json.load(f)
            assert "voltages" in self.project
        except Exception as e:
            messagebox.showerror(APP, f"Could not open file:\n{e}")
            return False
        self.path = path
        self.refresh_tree()
        return True

    def save(self, as_new=False):
        if not self.project:
            return
        if as_new or not self.path:
            p = filedialog.asksaveasfilename(defaultextension=".ccs", filetypes=FILE_TYPES,
                                             initialfile=self.project["name"])
            if not p:
                return
            self.path = p
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self.project, f, indent=1)
        self.title(f"{APP} - {self.project['name']}")

    def quit_app(self):
        if self.project and messagebox.askyesnocancel(APP, "Save project before exiting?") is True:
            self.save()
        self.destroy()

    # ---------- tree: voltages / bays ----------
    def refresh_tree(self, select=None):
        self.tree.delete(*self.tree.get_children())
        self.lbl_sub.config(text=f"Project: {self.project['name']}\nSubstation: {self.project['substation'] or '-'}")
        self.title(f"{APP} - {self.project['name']}")
        for kv in sorted(self.project["voltages"], key=lambda k: -int(k)):
            self.tree.insert("", "end", iid=f"v|{kv}", text=f"{kv} kV", open=True)
            for bay, d in self.project["voltages"][kv]["bays"].items():
                mark = "  [done]" if d["status"] == "Completed" else ""
                self.tree.insert(f"v|{kv}", "end", iid=f"b|{kv}|{bay}", text=bay + mark)
        if select and self.tree.exists(select):
            self.tree.selection_set(select)
        else:
            self.on_select()

    def sel(self):
        s = self.tree.selection()
        if not s:
            return None
        p = s[0].split("|")
        return p[0], p[1], (p[2] if len(p) > 2 else None)

    def cur(self):
        s = self.sel()
        if s and s[0] == "b":
            return s[1], s[2], self.project["voltages"][s[1]]["bays"][s[2]]
        return None

    def add_voltage(self):
        d = VoltageDialog(self, self.project["voltages"].keys())
        self.wait_window(d)
        if d.result:
            kv, names = d.result
            bays = self.project["voltages"].setdefault(kv, {"bays": {}})["bays"]
            for n in names:
                bays.setdefault(n, new_bay())
            self.refresh_tree()

    def add_bay(self):
        s = self.sel()
        if not s:
            return messagebox.showinfo(APP, "Select a voltage level first.")
        kv = s[1]
        name = simpledialog.askstring("Add bay", f"New bay name under {kv} kV:", parent=self)
        if name and name.strip():
            self.project["voltages"][kv]["bays"].setdefault(name.strip(), new_bay())
            self.refresh_tree(f"b|{kv}|{name.strip()}")

    def delete_selected(self):
        s = self.sel()
        if not s:
            return
        kind, kv, bay = s
        if kind == "v":
            if messagebox.askyesno(APP, f"Delete the whole {kv} kV level and all its bays/schedules?"):
                del self.project["voltages"][kv]
        elif messagebox.askyesno(APP, f"Delete bay '{bay}' ({kv} kV) and its schedule?"):
            del self.project["voltages"][kv]["bays"][bay]
        self.refresh_tree()

    def rename_sub(self):
        n = simpledialog.askstring("Substation", "Substation name:", initialvalue=self.project["substation"], parent=self)
        if n is not None:
            self.project["substation"] = n.strip()
            self.refresh_tree()

    # ---------- per-bay work ----------
    def on_select(self):
        c = self.cur()
        self.lst.delete(0, "end")
        self.tv_t.delete(*self.tv_t.get_children())
        self.tv_c.delete(*self.tv_c.get_children())
        if not c:
            self.lbl_bay.config(text="Select a bay on the left to begin.")
            return
        kv, bay, d = c
        self.lbl_bay.config(text=f"{kv} kV  >  {bay}   ({d['status']})")
        for p in d["drawings"]:
            self.lst.insert("end", p)
        self.fill_terminals(d)
        self.fill_cables(d)

    def fill_terminals(self, d):
        self.tv_t.delete(*self.tv_t.get_children())
        for i, t in enumerate(d["terminals"]):
            self.tv_t.insert("", "end", iid=str(i), values=(t["title"], t["tb"], t["terminal"], t["ferrule"],
                                                            "Yes" if t["ext"] else "No", t["desc"]))

    def fill_cables(self, d):
        self.tv_c.delete(*self.tv_c.get_children())
        for i, c in enumerate(d["cables"]):
            self.tv_c.insert("", "end", iid=str(i), values=(i + 1, c["cable_no"], c["title"], c["tb"],
                                                            c["ferrules"], c["cores"], c["to"], c["type"], c["size"]))

    def need_bay(self):
        c = self.cur()
        if not c:
            messagebox.showinfo(APP, "Select a bay first.")
        return c

    def upload(self):
        c = self.need_bay()
        if not c:
            return
        files = filedialog.askopenfilenames(title=f"Drawings for {c[1]}",
                                            filetypes=[("Drawings", "*.pdf *.dxf"), ("All files", "*.*")])
        for f in files:
            if f not in c[2]["drawings"]:
                c[2]["drawings"].append(f)
        self.on_select()

    def remove_drawing(self):
        c = self.cur()
        if c and self.lst.curselection():
            c[2]["drawings"].pop(self.lst.curselection()[0])
            self.on_select()

    def extract(self):
        c = self.need_bay()
        if not c:
            return
        if not c[2]["drawings"]:
            return messagebox.showinfo(APP, "Upload drawings for this bay first.")
        rows, errs = [], []
        self.config(cursor="watch")
        self.update_idletasks()
        for p in c[2]["drawings"]:
            try:
                rows += extract_drawing(p)
            except Exception as e:
                errs.append(f"{os.path.basename(p)}: {e}")
        self.config(cursor="")
        c[2]["terminals"] = rows
        self.fill_terminals(c[2])
        self.nb.select(1)
        msg = f"Extracted {len(rows)} row(s). Review the External column, then generate the schedule."
        if errs:
            msg += "\n\nProblems:\n" + "\n".join(errs)
            if any("fitz" in e or "ezdxf" in e for e in errs):
                msg += "\n\nInstall dependencies: pip install pymupdf ezdxf openpyxl"
        messagebox.showinfo(APP, msg)

    def cell_edit(self, tv, event, key, editable, toggle=None, toggle_col=None):
        """Double-click: toggle External or edit a text cell in place."""
        c = self.cur()
        row, col = tv.identify_row(event.y), tv.identify_column(event.x)
        if not c or not row:
            return
        ci = int(col[1:]) - 1
        rec = c[2][key][int(row)]
        if toggle and ci == toggle_col:
            rec[toggle] = not rec[toggle]
            return self.fill_terminals(c[2])
        field = next((k for k, v in editable.items() if v == ci), None)
        if not field:
            return
        x, y, w, h = tv.bbox(row, col)
        e = ttk.Entry(tv)
        e.place(x=x, y=y, width=w, height=h)
        e.insert(0, str(rec[field]))
        e.focus()

        def done(_=None):
            rec[field] = e.get()
            e.destroy()
            (self.fill_terminals if key == "terminals" else self.fill_cables)(c[2])
        e.bind("<Return>", done)
        e.bind("<FocusOut>", done)
        e.bind("<Escape>", lambda _: e.destroy())

    def add_terminal(self):
        c = self.need_bay()
        if c:
            c[2]["terminals"].append({"title": "", "tb": "", "terminal": "", "ferrule": "",
                                      "desc": "(manual)", "ext": True, "file": "", "page": 0})
            self.fill_terminals(c[2])

    def del_terminals(self):
        c = self.cur()
        if c:
            for i in sorted((int(i) for i in self.tv_t.selection()), reverse=True):
                c[2]["terminals"].pop(i)
            self.fill_terminals(c[2])

    def generate(self):
        c = self.need_bay()
        if not c:
            return
        kv, bay, d = c
        d["cables"] = generate_cables(kv, bay, d["terminals"])
        self.fill_cables(d)
        self.nb.select(2)
        if not d["cables"]:
            messagebox.showwarning(APP, "No external terminals are ticked, so no cables were generated.")

    def mark_done(self):
        c = self.need_bay()
        if c:
            c[2]["status"] = "Completed"
            self.save() if self.path else None
            self.refresh_tree(f"b|{c[0]}|{c[1]}")
            messagebox.showinfo(APP, "Bay marked completed. Select another bay on the left to continue.")

    def export(self, all_bays=False):
        if not self.project:
            return
        if all_bays:
            bays = [(kv, b, d) for kv in self.project["voltages"] for b, d in self.project["voltages"][kv]["bays"].items() if d["cables"]]
        else:
            c = self.need_bay()
            if not c:
                return
            bays = [c]
        if not bays or not any(b[2]["cables"] for b in bays):
            return messagebox.showinfo(APP, "Nothing to export - generate a cable schedule first.")
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                         initialfile=f"Cable_Schedule_{self.project['name']}")
        if p:
            try:
                export_excel(p, self.project, bays)
                messagebox.showinfo(APP, f"Saved:\n{p}")
            except Exception as e:
                messagebox.showerror(APP, f"Export failed:\n{e}")


if __name__ == "__main__":
    print(f"Starting {APP}…", flush=True)
    App().mainloop()