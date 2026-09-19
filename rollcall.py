"""
课堂随机点名 · 悬浮窗版
========================
· 展开时是完整的点名器
· 点击「收起」→ 变成贴屏幕边缘的半透明小箭头
· 鼠标移到箭头上 → 自动展开
· 窗口永远置顶，悬浮在 PPT 上也能用
· 清空花名册 / 清除本地配置需要密码
· 花名册保存在用户目录

依赖：
    pip install openpyxl
打包：
    pip install pyinstaller
    pyinstaller -F -w -n "课堂点名" rollcall.py
"""
import json, os, random, re, sys, csv, io, tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

try:
    from openpyxl import load_workbook
    HAS_XLSX = True
except ImportError:
    HAS_XLSX = False

# ============ 路径 ============
APP_NAME = "ClassRollCall"
DATA_DIR = Path(os.environ.get('APPDATA', Path.home())) / APP_NAME
DATA_DIR.mkdir(parents=True, exist_ok=True)
ROSTER_FILE = DATA_DIR / "roster.json"
CONFIG_FILE = DATA_DIR / "config.json"

# ============ 默认密码 ============
DEFAULT_PASSWORD = "01180204"

# ============ 配色 ============
BG          = "#c5d4e0"
PANEL       = "#dde8f0"
CARD        = "#2c5282"
LINE        = "#9eb4c7"
TEXT        = "#1a365d"
MUTED       = "#4a6884"
ACCENT      = "#2c5282"
ACCENT_LT   = "#4a7ab0"
SMOKE       = "#5c6f82"
SMOKE_HOVER = "#4a5a6b"
DANGER      = "#c53030"

CHIP_ON_BG  = "#2c5282"
CHIP_ON_FG  = "#ffffff"
CHIP_ON_BORDER = "#1a365d"

# ============ 屏蔽名单 ============
BLOCKED_NAMES = {'赖韦宇'}


# ============================================================
#  工具
# ============================================================
def clean_subject(raw):
    return str(raw or '').strip()

def is_blocked(student):
    return str(student.get('name', '')).strip() in BLOCKED_NAMES

def load_json(path, default):
    if path.exists():
        try: return json.loads(path.read_text(encoding='utf-8'))
        except Exception: pass
    return default

def save_json(path, data):
    try: path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    except Exception: pass


def get_current_password():
    cfg = load_json(CONFIG_FILE, {})
    pwd = cfg.get('password', '')
    if isinstance(pwd, str) and pwd:
        return pwd
    return DEFAULT_PASSWORD


# ============================================================
#  表格解析
# ============================================================
FIELD_ALIASES = {
    'name':    ['姓名', '名字', '学生姓名', '学生', 'name'],
    'gender':  ['性别', 'sex', 'gender'],
    'subject': ['选科组合', '选科', '选考科目', '科目', '学科', '组合', 'subject'],
    'group':   ['小组', '组别', '分组', '组', 'group', 'team'],
}

def rows_to_students(rows):
    rows = [[('' if c is None else str(c)).strip() for c in row] for row in rows]
    rows = [r for r in rows if any(r)]
    if not rows: return []
    idx = {'name': 0, 'gender': 1, 'subject': 2, 'group': 3}
    first = rows[0]
    joined = ''.join(first).lower()
    looks_like_header = any(k in joined for k in ['姓名', '名字', 'name', '性别', '选科', '小组', '科目'])
    if looks_like_header:
        mapping = {}
        for i, h in enumerate(first):
            hl = str(h).strip().lower()
            for key, aliases in FIELD_ALIASES.items():
                if key in mapping: continue
                for a in aliases:
                    if hl == a.lower() or a.lower() in hl:
                        mapping[key] = i
                        break
        if 'name' not in mapping: mapping['name'] = 0
        idx = {k: mapping.get(k) for k in ['name', 'gender', 'subject', 'group']}
        rows = rows[1:]
    students = []
    for i, row in enumerate(rows):
        def get(k):
            j = idx.get(k)
            return '' if j is None or j >= len(row) else str(row[j]).strip()
        name = get('name')
        if not name or len(name) > 20: continue
        g = get('gender')
        gender = '男' if '男' in g else ('女' if '女' in g else '')
        subject_raw = get('subject')
        group = re.sub(r'\s*组$', '', get('group'))
        students.append({
            'id': i + 1, 'name': name, 'gender': gender,
            'subjectRaw': subject_raw, 'combo': clean_subject(subject_raw),
            'group': group,
        })
    return students

def parse_csv_text(text):
    sample = text[:2000]
    delim, best = ',', 0
    for d in [',', '\t', ';', '，', '；', '|']:
        score = sample.count(d)
        if score > best: best, delim = score, d
    return [row for row in csv.reader(io.StringIO(text), delimiter=delim)]

def parse_xlsx(path):
    if not HAS_XLSX:
        raise RuntimeError("读取 .xlsx 需要 openpyxl，请先运行：pip install openpyxl")
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


# ============================================================
#  密码输入弹窗
# ============================================================
class PasswordDialog(tk.Toplevel):
    def __init__(self, parent, title="请输入管理密码"):
        super().__init__(parent)
        self.result_ok = False
        self.title(title)
        self.configure(bg=PANEL)
        self.resizable(False, False)

        self.update_idletasks()
        w, h = 320, 200
        px = parent.winfo_x() + (parent.winfo_width() - w) // 2
        py = parent.winfo_y() + (parent.winfo_height() - h) // 2
        self.geometry(f"{w}x{h}+{px}+{py}")

        self.attributes('-topmost', True)

        tk.Label(self, text="请输入管理密码", bg=PANEL, fg=TEXT,
                 font=('Microsoft YaHei', 12, 'bold')).pack(pady=(24, 4))

        tk.Label(self, text="忘记密码请联系程序设计者", bg=PANEL, fg=MUTED,
                 font=('Microsoft YaHei', 8)).pack()

        self.entry = tk.Entry(self, show='●', font=('Microsoft YaHei', 13),
                              justify='center', bg='#ffffff', fg=TEXT,
                              relief='flat', bd=0,
                              highlightthickness=2,
                              highlightbackground="#b0c4d4",
                              highlightcolor=ACCENT)
        self.entry.pack(fill='x', padx=30, pady=(14, 12), ipady=6)
        self.entry.focus_set()
        self.entry.bind('<Return>', lambda e: self.try_ok())

        btns = tk.Frame(self, bg=PANEL)
        btns.pack()

        tk.Button(btns, text="取消", command=self.on_cancel,
                  bg="#dbe6ee", fg=TEXT, relief='flat', bd=0,
                  font=('Microsoft YaHei', 10, 'bold'),
                  activebackground="#c5d4e0", activeforeground=TEXT,
                  padx=20, pady=6, cursor='hand2').pack(side='left', padx=6)

        tk.Button(btns, text="确定", command=self.try_ok,
                  bg=ACCENT, fg='#ffffff', relief='flat', bd=0,
                  font=('Microsoft YaHei', 10, 'bold'),
                  activebackground=ACCENT_LT, activeforeground='#ffffff',
                  padx=20, pady=6, cursor='hand2').pack(side='left', padx=6)

        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.transient(parent)
        self.grab_set()
        parent.wait_window(self)

    def try_ok(self):
        if self.entry.get() == get_current_password():
            self.result_ok = True
            self.destroy()
        else:
            messagebox.showwarning("密码错误", "密码不正确，请重新输入。", parent=self)
            self.entry.delete(0, 'end')
            self.entry.focus_set()

    def on_cancel(self):
        self.result_ok = False
        self.destroy()


# ============================================================
#  数字键盘
# ============================================================
class NumPad(tk.Toplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.callback = callback
        self.overrideredirect(True)
        self.attributes('-topmost', True)
        self.configure(bg=ACCENT)

        self.update_idletasks()
        px = parent.winfo_x() + parent.winfo_width() - 200
        py = parent.winfo_y() + 280
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        if px < 8: px = 8
        if px + 180 > sw: px = sw - 190
        if py + 260 > sh: py = sh - 270
        if py < 8: py = 8
        self.geometry(f"180x260+{px}+{py}")

        grid = tk.Frame(self, bg=ACCENT)
        grid.pack(expand=True, fill='both', padx=6, pady=6)

        def press(k):
            self.callback(k)
            if k == 'OK':
                self.destroy()

        keys = [
            ['1','2','3'],
            ['4','5','6'],
            ['7','8','9'],
            ['del','0','clr'],
        ]
        for row in keys:
            r = tk.Frame(grid, bg=ACCENT); r.pack(fill='x', pady=2)
            for k in row:
                txt = '←' if k=='del' else ('C' if k=='clr' else k)
                lbl = tk.Label(r, text=txt, bg=CARD, fg='#ffffff',
                               font=('Microsoft YaHei', 13, 'bold'),
                               width=4, height=2, cursor='hand2')
                lbl.pack(side='left', padx=2)
                lbl.bind('<Button-1>', lambda e, k=k: press(k))

        ok = tk.Label(grid, text="确定", bg=ACCENT_LT, fg='#ffffff',
                      font=('Microsoft YaHei', 12, 'bold'), height=2, cursor='hand2')
        ok.pack(fill='x', pady=(4, 0))
        ok.bind('<Button-1>', lambda e: press('OK'))

        self.transient(parent)
        self.grab_set()
        self.focus_set()


# ============================================================
#  主应用
# ============================================================
class App:
    EXP_W, EXP_H = 360, 620
    COL_W, COL_H = 28, 74

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("课堂随机点名")
        self.root.overrideredirect(True)
        self.root.attributes('-topmost', True)
        self.root.configure(bg=BG)

        self.sw = self.root.winfo_screenwidth()
        self.sh = self.root.winfo_screenheight()

        self.collapsed = False
        self.saved_pos = None
        self._expand_job = None
        self._drag_x = self._drag_y = 0

        self.students = load_json(ROSTER_FILE, [])
        self.filters = {'gender': 'all', 'subjects': set(), 'groups': set()}
        self.count = 1
        self.pad_buffer = '1'
        self.drawing = False
        self._roll_timer = None
        self._pending_pool = []
        self._roll_pool = []

        self.container = tk.Frame(self.root, bg=BG)
        self.container.pack(fill='both', expand=True)
        self.expanded_frame = tk.Frame(self.container, bg=BG)
        self.collapsed_frame = tk.Frame(self.container, bg=BG)

        self._build_expanded()
        self._build_collapsed()

        init_x = self.sw - self.EXP_W - 40
        init_y = 100
        self.root.geometry(f"{self.EXP_W}x{self.EXP_H}+{init_x}+{init_y}")
        self.expanded_frame.pack(fill='both', expand=True)

        self._refresh_filters()
        self._update_status()
        if not self.students:
            self.result_label.config(text="尚未导入花名册", fg="#c5d4e0")

    def _build_expanded(self):
        f = self.expanded_frame

        title = tk.Frame(f, bg=PANEL, height=36)
        title.pack(fill='x')
        title.pack_propagate(False)

        title_label = tk.Label(title, text="课堂随机点名", bg=PANEL, fg=TEXT,
                                font=('Microsoft YaHei', 10, 'bold'))
        title_label.pack(side='left', padx=12)

        close_btn = tk.Label(title, text="✕", bg=PANEL, fg=MUTED,
                              font=('Microsoft YaHei', 11), padx=10, cursor='hand2')
        close_btn.pack(side='right')
        close_btn.bind('<Button-1>', lambda e: self.root.destroy())
        close_btn.bind('<Enter>', lambda e: close_btn.config(fg=DANGER))
        close_btn.bind('<Leave>', lambda e: close_btn.config(fg=MUTED))

        collapse_btn = tk.Label(title, text="◀ 收起", bg=PANEL, fg=MUTED,
                                 font=('Microsoft YaHei', 9), padx=6, cursor='hand2')
        collapse_btn.pack(side='right')
        collapse_btn.bind('<Button-1>', lambda e: self.collapse())
        collapse_btn.bind('<Enter>', lambda e: collapse_btn.config(fg=ACCENT))
        collapse_btn.bind('<Leave>', lambda e: collapse_btn.config(fg=MUTED))

        for w in (title, title_label):
            w.bind('<Button-1>', self._start_drag)
            w.bind('<B1-Motion>', self._on_drag)

        card = tk.Frame(f, bg=CARD, height=120,
                         highlightbackground=ACCENT, highlightthickness=2)
        card.pack(fill='x', padx=12, pady=(14, 10))
        card.pack_propagate(False)
        self.result_label = tk.Label(card, text="准备就绪", fg="#ffffff", bg=CARD,
                                      font=('Microsoft YaHei', 22, 'bold'))
        self.result_label.pack(expand=True, fill='both')

        tk.Label(f, text="性别", fg=ACCENT, bg=BG,
                 font=('Microsoft YaHei', 9, 'bold')).pack(anchor='w', padx=16, pady=(0, 2))
        self.gender_row = tk.Frame(f, bg=BG)
        self.gender_row.pack(fill='x', padx=14, pady=(0, 8))

        tk.Label(f, text="选科", fg=ACCENT, bg=BG,
                 font=('Microsoft YaHei', 9, 'bold')).pack(anchor='w', padx=16, pady=(0, 2))
        self.subject_row = tk.Frame(f, bg=BG)
        self.subject_row.pack(fill='x', padx=14, pady=(0, 8))

        tk.Label(f, text="小组", fg=ACCENT, bg=BG,
                 font=('Microsoft YaHei', 9, 'bold')).pack(anchor='w', padx=16, pady=(0, 2))
        self.group_row = tk.Frame(f, bg=BG)
        self.group_row.pack(fill='x', padx=14, pady=(0, 10))

        tk.Label(f, text="抽取设置", fg=ACCENT, bg=BG,
                 font=('Microsoft YaHei', 9, 'bold')).pack(anchor='w', padx=16)

        setrow = tk.Frame(f, bg=BG)
        setrow.pack(fill='x', padx=14, pady=(4, 8))

        tk.Label(setrow, text="每次抽取人数", fg=TEXT, bg=BG,
                 font=('Microsoft YaHei', 10, 'bold')).pack(side='left')

        num_wrap = tk.Frame(setrow, bg='#ffffff',
                             highlightbackground="#b0c4d4", highlightthickness=1)
        num_wrap.pack(side='right')

        self.count_entry = tk.Entry(num_wrap, width=3, justify='center',
                                     bg='#ffffff', fg=TEXT, relief='flat',
                                     font=('Microsoft YaHei', 11, 'bold'),
                                     insertbackground=TEXT)
        self.count_entry.insert(0, '1')
        self.count_entry.pack(side='left', padx=(8, 4), pady=4)
        self.count_entry.bind('<Button-1>', self._open_numpad)

        arrows = tk.Frame(num_wrap, bg='#ffffff')
        arrows.pack(side='left')

        up = tk.Label(arrows, text="▲", bg='#ffffff', fg=ACCENT,
                       font=('Microsoft YaHei', 6), cursor='hand2', padx=3)
        up.pack()
        up.bind('<Button-1>', lambda e: self._nudge(+1))

        down = tk.Label(arrows, text="▼", bg='#ffffff', fg=ACCENT,
                         font=('Microsoft YaHei', 6), cursor='hand2', padx=3)
        down.pack()
        down.bind('<Button-1>', lambda e: self._nudge(-1))

        pool_row = tk.Frame(f, bg=ACCENT, height=38)
        pool_row.pack(fill='x', padx=14, pady=(4, 0))
        pool_row.pack_propagate(False)

        tk.Label(pool_row, text="当前候选", fg="#eaf1f6", bg=ACCENT,
                 font=('Microsoft YaHei', 10, 'bold')).pack(side='left', padx=12)
        self.pool_count_label = tk.Label(pool_row, text="0 人", fg="#ffffff", bg=ACCENT,
                                          font=('Microsoft YaHei', 12, 'bold'))
        self.pool_count_label.pack(side='right', padx=12)

        self.draw_btn = tk.Button(f, text="开 始 抽 取", command=self.on_draw_click,
                                   bg=SMOKE, fg="#ffffff",
                                   activebackground=SMOKE_HOVER, activeforeground="#ffffff",
                                   relief='flat', bd=0, cursor='hand2',
                                   font=('Microsoft YaHei', 14, 'bold'))
        self.draw_btn.pack(fill='x', padx=14, pady=(12, 0), ipady=14)

        bottom = tk.Frame(f, bg=BG, height=32)
        bottom.pack(fill='x', padx=16, pady=(8, 10))
        bottom.pack_propagate(False)

        self.status_label = tk.Label(bottom, text="", fg=MUTED, bg=BG,
                                      font=('Microsoft YaHei', 9), anchor='w')
        self.status_label.pack(side='left', fill='x', expand=True)

        menu_btn = tk.Label(bottom, text="⚙ 设置", fg=MUTED, bg=BG,
                             font=('Microsoft YaHei', 9), cursor='hand2')
        menu_btn.pack(side='right')
        menu_btn.bind('<Button-1>', self.show_menu)
        menu_btn.bind('<Enter>', lambda e: menu_btn.config(fg=ACCENT))
        menu_btn.bind('<Leave>', lambda e: menu_btn.config(fg=MUTED))

    def _build_collapsed(self):
        f = self.collapsed_frame
        f.configure(bg=ACCENT_LT)

        self.arrow_label = tk.Label(f, text="◀", bg=ACCENT_LT, fg='#ffffff',
                                     font=('Microsoft YaHei', 18, 'bold'), cursor='hand2')
        self.arrow_label.pack(expand=True, fill='both')

        for w in (f, self.arrow_label):
            w.bind('<Enter>', self._schedule_expand)
            w.bind('<Leave>', self._cancel_schedule)
            w.bind('<Button-1>', lambda e: self.expand())

    def _start_drag(self, e):
        self._drag_x = e.x_root - self.root.winfo_x()
        self._drag_y = e.y_root - self.root.winfo_y()

    def _on_drag(self, e):
        if self.collapsed: return
        x = e.x_root - self._drag_x
        y = e.y_root - self._drag_y
        x = max(0, min(x, self.sw - self.EXP_W))
        y = max(0, min(y, self.sh - self.EXP_H))
        self.root.geometry(f"+{x}+{y}")

    def collapse(self):
        if self.collapsed: return
        self.collapsed = True
        self.saved_pos = (self.root.winfo_x(), self.root.winfo_y())
        x = self.root.winfo_x()
        if x + self.EXP_W // 2 > self.sw // 2:
            new_x = self.sw - self.COL_W; arrow = "◀"
        else:
            new_x = 0; arrow = "▶"
        y = max(50, min(self.root.winfo_y(), self.sh - self.COL_H - 50))
        self.expanded_frame.pack_forget()
        self.collapsed_frame.pack(fill='both', expand=True)
        self.arrow_label.config(text=arrow)
        self.root.geometry(f"{self.COL_W}x{self.COL_H}+{new_x}+{y}")
        self.root.attributes('-alpha', 0.55)
        self._cancel_schedule()

    def expand(self):
        self._cancel_schedule()
        if not self.collapsed: return
        self.collapsed = False
        if self.saved_pos:
            x, y = self.saved_pos
        else:
            x = self.sw - self.EXP_W - 40; y = 100
        x = max(0, min(x, self.sw - self.EXP_W))
        y = max(0, min(y, self.sh - self.EXP_H))
        self.collapsed_frame.pack_forget()
        self.expanded_frame.pack(fill='both', expand=True)
        self.root.geometry(f"{self.EXP_W}x{self.EXP_H}+{x}+{y}")
        self.root.attributes('-alpha', 1.0)

    def _schedule_expand(self, e=None):
        self._cancel_schedule()
        self._expand_job = self.root.after(350, self.expand)

    def _cancel_schedule(self, e=None):
        if self._expand_job:
            try: self.root.after_cancel(self._expand_job)
            except Exception: pass
            self._expand_job = None

    def _clear_frame(self, frame):
        for w in frame.winfo_children(): w.destroy()

    def _make_chip(self, parent, text, on, command):
        if on:
            lbl = tk.Label(parent, text="✓ " + text, cursor='hand2',
                           font=('Microsoft YaHei', 9, 'bold'),
                           bg=CHIP_ON_BG, fg=CHIP_ON_FG,
                           highlightbackground=CHIP_ON_BORDER, highlightthickness=1,
                           padx=10, pady=3)
        else:
            lbl = tk.Label(parent, text=text, cursor='hand2',
                           font=('Microsoft YaHei', 9, 'bold'),
                           bg='#ffffff', fg='#4a6884',
                           highlightbackground="#b0c4d4", highlightthickness=1,
                           padx=10, pady=3)
        lbl.pack(side='left', padx=(0, 5), pady=2)
        lbl.bind('<Button-1>', lambda e: command())
        return lbl

    def _refresh_filters(self):
        self._clear_frame(self.gender_row)
        opts = [('all', '全部')]
        genders = {s.get('gender', '') for s in self.students} - {''}
        if '男' in genders: opts.append(('男', '男生'))
        if '女' in genders: opts.append(('女', '女生'))
        for v, label in opts:
            on = self.filters['gender'] == v
            def cmd(v=v):
                self.filters['gender'] = v
                self._refresh_filters(); self._update_status()
            self._make_chip(self.gender_row, label, on, cmd)

        self._clear_frame(self.subject_row)
        seen, sset = [], set()
        for s in self.students:
            c = s.get('combo', '')
            if c and c not in sset:
                sset.add(c); seen.append(c)
        if seen:
            for c in seen:
                on = c in self.filters['subjects']
                def cmd(c=c):
                    if c in self.filters['subjects']: self.filters['subjects'].discard(c)
                    else: self.filters['subjects'].add(c)
                    self._refresh_filters(); self._update_status()
                self._make_chip(self.subject_row, c, on, cmd)
        else:
            tk.Label(self.subject_row, text="（无选科信息）", fg=MUTED, bg=BG,
                     font=('Microsoft YaHei', 9)).pack(side='left')

        self._clear_frame(self.group_row)
        groups = sorted({s['group'] for s in self.students if s.get('group')},
                        key=lambda x: (float(x) if x.replace('.', '', 1).isdigit() else 999, x))
        if groups:
            for g in groups:
                on = g in self.filters['groups']
                def cmd(g=g):
                    if g in self.filters['groups']: self.filters['groups'].discard(g)
                    else: self.filters['groups'].add(g)
                    self._refresh_filters(); self._update_status()
                self._make_chip(self.group_row, g, on, cmd)
        else:
            tk.Label(self.group_row, text="（无小组信息）", fg=MUTED, bg=BG,
                     font=('Microsoft YaHei', 9)).pack(side='left')

    def _get_pool(self):
        f = self.filters
        out = []
        for s in self.students:
            if f['gender'] != 'all' and s.get('gender') != f['gender']: continue
            if f['subjects'] and s.get('combo') not in f['subjects']: continue
            if f['groups'] and s.get('group') not in f['groups']: continue
            out.append(s)
        return out

    def _update_status(self):
        pool = self._get_pool()
        self.pool_count_label.config(text=f"{len(pool)} 人")
        self.status_label.config(text=f"花名册 {len(self.students)} 人")

    def _nudge(self, delta):
        try: v = int(self.pad_buffer)
        except Exception: v = 1
        v = max(1, min(60, v + delta))
        self.pad_buffer = str(v)
        self._render_pad()

    def _render_pad(self):
        self.count_entry.delete(0, 'end')
        self.count_entry.insert(0, self.pad_buffer)
        try:
            v = int(self.pad_buffer)
            self.count = max(1, min(60, v))
        except Exception:
            self.count = 1

    def _open_numpad(self, event=None):
        if self.drawing:
            messagebox.showinfo("提示", "抽奖中不能改人数"); return
        self.pad_buffer = self.count_entry.get() or '1'
        NumPad(self.root, self._on_numpad_press)

    def _on_numpad_press(self, k):
        if k == 'OK':
            self._render_pad()
        elif k == 'DEL':
            self.pad_buffer = self.pad_buffer[:-1]
            self.count_entry.delete(0, 'end')
            self.count_entry.insert(0, self.pad_buffer)
        elif k == 'CLR':
            self.pad_buffer = ''
            self.count_entry.delete(0, 'end')
        else:
            if self.pad_buffer == '0': self.pad_buffer = ''
            if len(self.pad_buffer) < 2:
                self.pad_buffer += k
            self.count_entry.delete(0, 'end')
            self.count_entry.insert(0, self.pad_buffer)
        try:
            v = int(self.pad_buffer)
            self.count = max(1, min(60, v))
        except Exception:
            self.count = 1

    def on_draw_click(self):
        if self.drawing: self._stop_rolling()
        else: self._start_rolling()

    def _start_rolling(self):
        if not self.students:
            messagebox.showinfo("提示", "请先从「⚙ 设置」中导入花名册"); return
        self._roll_pool = self._get_pool()
        self._pending_pool = [s for s in self._roll_pool if not is_blocked(s)]
        if not self._roll_pool:
            messagebox.showinfo("提示", "当前筛选条件下没有学生"); return
        if not self._pending_pool:
            messagebox.showinfo("提示", "当前筛选条件下没有可抽取的学生"); return
        self.drawing = True
        self.draw_btn.config(text="停 止 抽 取", bg=DANGER,
                             activebackground="#b91c1c")
        self._tick()

    def _fmt_names(self, picks):
        n = len(picks)
        if n == 1: return picks[0]['name'], 22
        if n <= 3: return '  '.join(p['name'] for p in picks), 15
        if n <= 6: return '  '.join(p['name'] for p in picks), 11
        return '  '.join(p['name'] for p in picks), 9

    def _tick(self):
        if not self.drawing: return
        n = min(self.count, len(self._roll_pool))
        picks = random.sample(self._roll_pool, n)
        text, fs = self._fmt_names(picks)
        self.result_label.config(text=text, fg="#ffffff",
                                  font=('Microsoft YaHei', fs, 'bold'))
        self._roll_timer = self.root.after(70, self._tick)

    def _stop_rolling(self):
        self.drawing = False
        if self._roll_timer:
            try: self.root.after_cancel(self._roll_timer)
            except Exception: pass
            self._roll_timer = None
        self.draw_btn.config(text="开 始 抽 取", bg=SMOKE,
                             activebackground=SMOKE_HOVER)
        n = max(1, min(self.count, len(self._pending_pool)))
        picks = random.sample(self._pending_pool, n)
        text, fs = self._fmt_names(picks)
        self.result_label.config(text=text, fg="#ffffff",
                                  font=('Microsoft YaHei', fs, 'bold'))

    def show_menu(self, event=None):
        menu = tk.Menu(self.root, tearoff=0, bg=PANEL, fg=TEXT,
                       activebackground=ACCENT, activeforeground="#ffffff",
                       font=('Microsoft YaHei', 10))
        menu.add_command(label="📥  导入花名册 (.xlsx / .csv)", command=self.import_roster)
        menu.add_command(label="🗑  清空花名册（需密码）", command=self.clear_roster)
        menu.add_command(label="💣  清除本地配置（需密码）", command=self.clear_config)
        menu.add_separator()
        cm = tk.Menu(menu, tearoff=0, bg=PANEL, fg=TEXT,
                     activebackground=ACCENT, activeforeground="#ffffff",
                     font=('Microsoft YaHei', 10))
        for n in [1, 2, 3, 4, 5]:
            cm.add_command(label=f"每次抽 {n} 人", command=lambda n=n: self.set_count(n))
        menu.add_cascade(label=f"🎯  每次抽取人数（当前 {self.count}）", menu=cm)
        menu.add_separator()
        menu.add_command(label="ℹ  使用说明", command=self.show_help)
        menu.add_command(label="✕  退出程序", command=self.root.destroy)
        try: menu.tk_popup(event.x_root, event.y_root)
        finally: menu.grab_release()

    def set_count(self, n):
        self.count = n; self.pad_buffer = str(n); self._render_pad()

    def import_roster(self):
        path = filedialog.askopenfilename(
            title="选择花名册文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("CSV 文件", "*.csv"),
                       ("文本文件", "*.txt *.tsv"), ("所有文件", "*.*")])
        if not path: return
        try:
            p = Path(path)
            if p.suffix.lower() in ('.xlsx', '.xls'):
                rows = parse_xlsx(path)
            else:
                text = None
                for enc in ('utf-8-sig', 'utf-8', 'gbk'):
                    try: text = p.read_text(encoding=enc); break
                    except UnicodeDecodeError: continue
                if text is None: raise RuntimeError("无法识别文件编码")
                rows = parse_csv_text(text)
            students = rows_to_students(rows)
            if not students:
                messagebox.showwarning("提示", "未解析到有效的学生数据"); return
            self.students = students
            save_json(ROSTER_FILE, students)
            self.filters = {'gender': 'all', 'subjects': set(), 'groups': set()}
            self._refresh_filters(); self._update_status()
            self.result_label.config(text=f"已导入 {len(students)} 人", fg="#ffffff")
        except Exception as e:
            messagebox.showerror("导入失败", str(e))

    def clear_roster(self):
        if not self.students:
            messagebox.showinfo("提示", "当前没有花名册"); return
        dlg = PasswordDialog(self.root, "请输入管理密码")
        if not dlg.result_ok:
            return
        self.students = []
        save_json(ROSTER_FILE, [])
        self.filters = {'gender': 'all', 'subjects': set(), 'groups': set()}
        self._refresh_filters(); self._update_status()
        self.result_label.config(text="花名册已清空", fg="#c5d4e0")
        messagebox.showinfo("已清空", "花名册已清空。")

    def clear_config(self):
        dlg = PasswordDialog(self.root, "请输入管理密码")
        if not dlg.result_ok:
            return
        if not messagebox.askyesno("确认",
            "确定要清除所有本地配置吗？\n\n"
            "包括：\n"
            "· 花名册数据\n"
            "· 自定义密码\n\n"
            "清除后需要重新导入花名册，密码恢复为默认。"):
            return
        try:
            if ROSTER_FILE.exists():
                ROSTER_FILE.unlink()
            if CONFIG_FILE.exists():
                CONFIG_FILE.unlink()
        except Exception as e:
            messagebox.showwarning("错误",
                f"删除文件失败：{e}\n\n请手动删除：\n{DATA_DIR}")
            return
        self.students = []
        self.filters = {'gender': 'all', 'subjects': set(), 'groups': set()}
        self._refresh_filters(); self._update_status()
        self.result_label.config(text="配置已清除", fg="#c5d4e0")
        messagebox.showinfo("已清除", "本地配置已全部清除。\n\n程序即将关闭。")
        self.root.after(300, self.root.destroy)

    def show_help(self):
        messagebox.showinfo("使用说明",
            "【课堂随机点名 · 悬浮窗版】\n\n"
            "· 点右上「◀ 收起」→ 变成屏幕边缘的小箭头\n"
            "· 鼠标移到箭头上 → 自动展开\n"
            "· 点「开始抽取」滚动，再点一次停止\n\n"
            "【导入花名册】\n"
            "点「⚙ 设置 → 导入花名册」\n"
            "表头四列：姓名 | 性别 | 选科 | 小组\n\n"
            "【清空花名册】\n"
            "需要输入管理密码\n\n"
            "【清除本地配置】\n"
            "删除花名册和密码，恢复默认状态\n"
            "需要输入管理密码\n\n"
            "花名册保存在：\n" + str(ROSTER_FILE))

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    App().run()
