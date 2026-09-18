"""
课堂随机点名 · 悬浮窗版
========================
· 展开时是完整的点名器
· 点击「收起」→ 变成贴屏幕边缘的半透明小箭头
· 鼠标移到箭头上 → 自动展开
· 窗口永远置顶，悬浮在 PPT 上也能用
· 花名册保存在用户目录，重装系统前不会丢

依赖：
    pip install openpyxl
打包：
    pip install pyinstaller
    pyinstaller -F -w -n 课堂点名 rollcall.py
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

# ============ 配色 ============
BG          = "#0f172a"
PANEL       = "#1e293b"
ACCENT      = "#38bdf8"
TEXT        = "#e2e8f0"
MUTED       = "#94a3b8"
BTN_BG      = "#334155"
GOLD        = "#fbbf24"
GOLD_HOVER  = "#f59e0b"
CHIP_ON_BG  = "#0ea5e9"
CHIP_ON_FG  = "#04212e"
DANGER      = "#ef4444"

SUBJECT_COMBOS = ['物化生', '物化地', '历政生', '历政地']

# ============ 工具 ============
def normalize_combo(raw):
    if not raw: return ''
    s = re.sub(r'\s+', '', str(raw))
    if not s: return ''
    hasW, hasH, hasS = '物' in s, '化' in s, '生' in s
    hasD, hasZ = '地' in s, '政' in s
    hasL = ('历' in s) or ('史' in s)
    if hasW and hasH and hasS: return '物化生'
    if hasW and hasH and hasD: return '物化地'
    if hasL and hasZ and hasS: return '历政生'
    if hasL and hasZ and hasD: return '历政地'
    return s

def load_json(path, default):
    if path.exists():
        try: return json.loads(path.read_text(encoding='utf-8'))
        except Exception: pass
    return default

def save_json(path, data):
    try: path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    except Exception: pass

# ============ 解析 ============
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
            'subjectRaw': subject_raw, 'combo': normalize_combo(subject_raw),
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

# ============ 开机自启 ============
def get_exe_path():
    return sys.executable if getattr(sys, 'frozen', False) else None

def is_autostart_enabled():
    if sys.platform != 'win32': return False
    try:
        import winreg
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
        try:
            winreg.QueryValueEx(k, APP_NAME); return True
        except FileNotFoundError: return False
        finally: winreg.CloseKey(k)
    except Exception: return False

def set_autostart(enabled):
    if sys.platform != 'win32': return False
    exe = get_exe_path()
    if not exe: return False
    try:
        import winreg
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
        if enabled:
            winreg.SetValueEx(k, APP_NAME, 0, winreg.REG_SZ, f'"{exe}"')
        else:
            try: winreg.DeleteValue(k, APP_NAME)
            except FileNotFoundError: pass
        winreg.CloseKey(k); return True
    except Exception: return False


# ================================================================
#  主应用
# ================================================================
class App:
    EXP_W, EXP_H = 340, 480      # 展开尺寸
    COL_W, COL_H = 28, 74        # 收起尺寸

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("课堂随机点名")
        self.root.overrideredirect(True)      # 无边框
        self.root.attributes('-topmost', True) # 永远置顶
        self.root.configure(bg=BG)

        self.sw = self.root.winfo_screenwidth()
        self.sh = self.root.winfo_screenheight()

        # 状态
        self.collapsed = False
        self.side = 'right'
        self.saved_pos = None
        self._expand_job = None
        self._drag_x = self._drag_y = 0

        # 数据
        self.students = load_json(ROSTER_FILE, [])
        self.filters = {'gender': 'all', 'subjects': set(), 'groups': set()}
        self.count = 1
        self.drawing = False
        self._final_picks = []
        self._roll_step = 0
        self._roll_total = 26
        self._roll_pool = []

        # 主容器
        self.container = tk.Frame(self.root, bg=BG)
        self.container.pack(fill='both', expand=True)
        self.expanded_frame = tk.Frame(self.container, bg=BG)
        self.collapsed_frame = tk.Frame(self.container, bg=BG)

        self._build_expanded()
        self._build_collapsed()

        # 初始位置（屏幕右侧偏上）
        init_x = self.sw - self.EXP_W - 40
        init_y = 120
        self.root.geometry(f"{self.EXP_W}x{self.EXP_H}+{init_x}+{init_y}")
        self.root.attributes('-alpha', 0.96)
        self.expanded_frame.pack(fill='both', expand=True)

        self._refresh_filters()
        self._update_status()
        if not self.students:
            self.result_label.config(text="尚未导入花名册", fg=MUTED)

    # ------------------------------------------------------------
    #  展开态 UI
    # ------------------------------------------------------------
    def _build_expanded(self):
        f = self.expanded_frame

        # 标题栏
        title = tk.Frame(f, bg=PANEL, height=32)
        title.pack(fill='x')
        title.pack_propagate(False)

        title_label = tk.Label(title, text="课堂随机点名", bg=PANEL, fg=TEXT,
                                font=('Microsoft YaHei', 10, 'bold'))
        title_label.pack(side='left', padx=10)

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

        # 结果卡片
        card = tk.Frame(f, bg=PANEL, height=110)
        card.pack(fill='x', padx=12, pady=(14, 10))
        card.pack_propagate(False)
        self.result_label = tk.Label(card, text="准备就绪", fg=ACCENT, bg=PANEL,
                                      font=('Microsoft YaHei', 20, 'bold'), wraplength=300)
        self.result_label.pack(expand=True)

        # 性别
        tk.Label(f, text="性别", fg=MUTED, bg=BG,
                 font=('Microsoft YaHei', 9)).pack(anchor='w', padx=14)
        self.gender_row = tk.Frame(f, bg=BG)
        self.gender_row.pack(fill='x', padx=12, pady=(2, 8))

        # 选科
        tk.Label(f, text="选科", fg=MUTED, bg=BG,
                 font=('Microsoft YaHei', 9)).pack(anchor='w', padx=14)
        self.subject_row = tk.Frame(f, bg=BG)
        self.subject_row.pack(fill='x', padx=12, pady=(2, 8))

        # 小组
        tk.Label(f, text="小组", fg=MUTED, bg=BG,
                 font=('Microsoft YaHei', 9)).pack(anchor='w', padx=14)
        self.group_row = tk.Frame(f, bg=BG)
        self.group_row.pack(fill='x', padx=12, pady=(2, 10))

        # 抽取按钮
        self.draw_btn = tk.Button(f, text="开 始 抽 取", command=self.start_draw,
                                   bg=GOLD, fg='#3b2500',
                                   activebackground=GOLD_HOVER, activeforeground='#3b2500',
                                   relief='flat', bd=0, cursor='hand2',
                                   font=('Microsoft YaHei', 14, 'bold'))
        self.draw_btn.pack(fill='x', padx=12, ipady=12)

        # 底部状态栏
        bottom = tk.Frame(f, bg=BG)
        bottom.pack(fill='x', padx=14, pady=(10, 12))
        self.status_label = tk.Label(bottom, text="", fg=MUTED, bg=BG,
                                      font=('Microsoft YaHei', 9), anchor='w')
        self.status_label.pack(side='left', fill='x', expand=True)

        menu_btn = tk.Label(bottom, text="⚙ 设置", fg=MUTED, bg=BG,
                             font=('Microsoft YaHei', 9), cursor='hand2')
        menu_btn.pack(side='right')
        menu_btn.bind('<Button-1>', self.show_menu)
        menu_btn.bind('<Enter>', lambda e: menu_btn.config(fg=ACCENT))
        menu_btn.bind('<Leave>', lambda e: menu_btn.config(fg=MUTED))

    # ------------------------------------------------------------
    #  收起态 UI
    # ------------------------------------------------------------
    def _build_collapsed(self):
        f = self.collapsed_frame
        f.configure(bg=ACCENT)

        self.arrow_label = tk.Label(f, text="◀", bg=ACCENT, fg="#04212e",
                                     font=('Microsoft YaHei', 18, 'bold'),
                                     cursor='hand2')
        self.arrow_label.pack(expand=True, fill='both')

        # 鼠标进入 → 悬停一小会儿后自动展开
        for w in (f, self.arrow_label):
            w.bind('<Enter>', self._schedule_expand)
            w.bind('<Leave>', self._cancel_schedule)
            w.bind('<Button-1>', lambda e: self.expand())

    # ------------------------------------------------------------
    #  拖动
    # ------------------------------------------------------------
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

    # ------------------------------------------------------------
    #  收起 / 展开
    # ------------------------------------------------------------
    def collapse(self):
        if self.collapsed: return
        self.collapsed = True
        self.saved_pos = (self.root.winfo_x(), self.root.winfo_y())

        x = self.root.winfo_x()
        cx = x + self.EXP_W // 2
        if cx > self.sw // 2:
            self.side = 'right'
            new_x = self.sw - self.COL_W
            arrow = "◀"
        else:
            self.side = 'left'
            new_x = 0
            arrow = "▶"

        y = self.root.winfo_y()
        y = max(50, min(y, self.sh - self.COL_H - 50))

        # 先切换内容再改尺寸，避免闪屏
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
            x = self.sw - self.EXP_W - 40
            y = 120

        x = max(0, min(x, self.sw - self.EXP_W))
        y = max(0, min(y, self.sh - self.EXP_H))

        self.collapsed_frame.pack_forget()
        self.expanded_frame.pack(fill='both', expand=True)
        self.root.geometry(f"{self.EXP_W}x{self.EXP_H}+{x}+{y}")
        self.root.attributes('-alpha', 0.96)

    # 悬停展开
    def _schedule_expand(self, e=None):
        self._cancel_schedule()
        self._expand_job = self.root.after(350, self.expand)

    def _cancel_schedule(self, e=None):
        if self._expand_job:
            try: self.root.after_cancel(self._expand_job)
            except Exception: pass
            self._expand_job = None

    # ------------------------------------------------------------
    #  筛选控件
    # ------------------------------------------------------------
    def _clear_frame(self, frame):
        for w in frame.winfo_children(): w.destroy()

    def _make_chip(self, parent, text, on, command):
        lbl = tk.Label(parent, text=text, cursor='hand2',
                       font=('Microsoft YaHei', 10),
                       bg=(CHIP_ON_BG if on else BTN_BG),
                       fg=(CHIP_ON_FG if on else TEXT),
                       padx=10, pady=4)
        lbl.pack(side='left', padx=(0, 5), pady=2)
        lbl.bind('<Button-1>', lambda e: command())
        return lbl

    def _refresh_filters(self):
        # 性别
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

        # 选科
        self._clear_frame(self.subject_row)
        all_combos = {s['combo'] for s in self.students if s.get('combo')}
        if all_combos:
            std = [c for c in SUBJECT_COMBOS if c in all_combos]
            extra = sorted(all_combos - set(SUBJECT_COMBOS))
            for c in std + extra:
                on = c in self.filters['subjects']
                def cmd(c=c):
                    if c in self.filters['subjects']: self.filters['subjects'].discard(c)
                    else: self.filters['subjects'].add(c)
                    self._refresh_filters(); self._update_status()
                self._make_chip(self.subject_row, c, on, cmd)
        else:
            tk.Label(self.subject_row, text="（无选科信息）", fg=MUTED, bg=BG,
                     font=('Microsoft YaHei', 9)).pack(side='left')

        # 小组
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
        self.status_label.config(text=f"花名册 {len(self.students)} 人 · 候选 {len(pool)} 人")

    # ------------------------------------------------------------
    #  抽取
    # ------------------------------------------------------------
    def start_draw(self):
        if self.drawing: return
        if not self.students:
            messagebox.showinfo("提示", "请先从「⚙ 设置」中导入花名册"); return
        pool = self._get_pool()
        if not pool:
            messagebox.showinfo("提示", "当前筛选条件下没有学生"); return
        n = min(self.count, len(pool))
        self._final_picks = random.sample(pool, n)
        self.drawing = True
        self.draw_btn.config(state='disabled', text='抽 取 中…')
        self._roll_step = 0
        self._roll_pool = pool
        self._roll_do()

    def _roll_do(self):
        if self._roll_step >= self._roll_total:
            names = '  '.join(p['name'] for p in self._final_picks)
            n = len(self._final_picks)
            fs = 22 if n == 1 else (15 if n <= 3 else 11)
            self.result_label.config(text=names, fg=GOLD,
                                     font=('Microsoft YaHei', fs, 'bold'))
            self.drawing = False
            self.draw_btn.config(state='normal', text='开 始 抽 取')
            return
        n = min(self.count, len(self._roll_pool))
        picks = random.sample(self._roll_pool, n)
        names = '  '.join(p['name'] for p in picks)
        fs = 22 if n == 1 else (15 if n <= 3 else 11)
        self.result_label.config(text=names, fg=ACCENT,
                                 font=('Microsoft YaHei', fs, 'bold'))
        self._roll_step += 1
        t = self._roll_step / self._roll_total
        delay = 45 + int(190 * (t ** 3))
        self.root.after(delay, self._roll_do)

    # ------------------------------------------------------------
    #  设置菜单
    # ------------------------------------------------------------
    def show_menu(self, event=None):
        menu = tk.Menu(self.root, tearoff=0, bg=PANEL, fg=TEXT,
                       activebackground=ACCENT, activeforeground=CHIP_ON_FG,
                       font=('Microsoft YaHei', 10))
        menu.add_command(label="📥  导入花名册 (.xlsx / .csv)", command=self.import_roster)
        menu.add_command(label="🗑  清空花名册", command=self.clear_roster)
        menu.add_separator()

        cm = tk.Menu(menu, tearoff=0, bg=PANEL, fg=TEXT,
                     activebackground=ACCENT, activeforeground=CHIP_ON_FG,
                     font=('Microsoft YaHei', 10))
        for n in [1, 2, 3, 4, 5]:
            cm.add_command(label=f"每次抽 {n} 人", command=lambda n=n: self.set_count(n))
        menu.add_cascade(label=f"🎯  每次抽取人数（当前 {self.count}）", menu=cm)

        menu.add_separator()
        if sys.platform == 'win32' and get_exe_path():
            enabled = is_autostart_enabled()
            menu.add_command(
                label=("✅  开机自启（已开启）" if enabled else "☐  开机自启"),
                command=lambda: self.toggle_autostart(not enabled))
        else:
            menu.add_command(label="（打包成 exe 后支持开机自启）", state='disabled')

        menu.add_separator()
        menu.add_command(label="ℹ  使用说明", command=self.show_help)
        menu.add_command(label="✕  退出程序", command=self.root.destroy)
        try: menu.tk_popup(event.x_root, event.y_root)
        finally: menu.grab_release()

    def set_count(self, n): self.count = n

    def toggle_autostart(self, enabled):
        if set_autostart(enabled):
            messagebox.showinfo("提示", "已开启开机自启" if enabled else "已关闭开机自启")
        else:
            messagebox.showwarning("提示", "操作失败")

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
            self.result_label.config(text=f"已导入 {len(students)} 人", fg=ACCENT)
        except Exception as e:
            messagebox.showerror("导入失败", str(e))

    def clear_roster(self):
        if not self.students: return
        if not messagebox.askyesno("确认", "确定要清空花名册吗？"): return
        self.students = []
        save_json(ROSTER_FILE, [])
        self.filters = {'gender': 'all', 'subjects': set(), 'groups': set()}
        self._refresh_filters(); self._update_status()
        self.result_label.config(text="花名册已清空", fg=MUTED)

    def show_help(self):
        messagebox.showinfo("使用说明",
            "【课堂随机点名 · 悬浮窗版】\n\n"
            "· 点右上「◀ 收起」→ 变成屏幕边缘的小箭头\n"
            "· 鼠标移到箭头上 → 自动展开\n"
            "· 箭头贴在屏幕左/右边缘，跟着窗口位置自动判断\n\n"
            "【导入花名册】\n"
            "点「⚙ 设置 → 导入花名册」\n"
            "表头四列：姓名 | 性别 | 选科 | 小组\n"
            "· 性别：男 / 女\n"
            "· 选科：物化生 / 物化地 / 历政生 / 历政地\n"
            "· 小组：1、2、3（可自定义）\n\n"
            "花名册保存在：\n" + str(ROSTER_FILE))

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    App().run()
