"""
课堂随机点名 · 悬浮窗版（PySide6 圆角真无边框）
================================================
· 真圆角 + 真透明
· 悬浮在 PPT 上也能用
· 收起后贴在屏幕边缘成半透明小箭头
· 清空花名册需要输入密码
· 支持开机自动启动
· 密码支持外部重置（配合 reset_password.py）
· 关闭时彻底释放文件，不会残留进程
"""
import json, os, random, re, sys, csv, io
from pathlib import Path

from PySide6.QtCore import Qt, QTimer, QPoint, QRectF, QEvent
from PySide6.QtGui import QFont, QColor, QPainter, QPainterPath, QBrush, QAction
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFrame, QHBoxLayout, QVBoxLayout,
    QGridLayout, QLineEdit, QDialog, QMenu, QFileDialog, QMessageBox
)

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
GREEN       = "#2f855a"
DANGER      = "#c53030"
CHIP_BG     = "#dbe6ee"
CHIP_BORDER = "#c5d4e0"
CHIP_TEXT   = "#8fa9c2"
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
    """优先读外部配置，没有就用默认密码"""
    cfg = load_json(CONFIG_FILE, {})
    pwd = cfg.get('password', '')
    if isinstance(pwd, str) and pwd:
        return pwd
    return DEFAULT_PASSWORD


# ============================================================
#  开机自启
# ============================================================
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
#  圆角容器
# ============================================================
class RoundedWidget(QWidget):
    def __init__(self, radius=18, bg=BG, parent=None):
        super().__init__(parent)
        self.radius = radius
        self.bg_color = QColor(bg)
        self.setAttribute(Qt.WA_TranslucentBackground, True)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        path = QPainterPath()
        rect = QRectF(0, 0, self.width(), self.height())
        path.addRoundedRect(rect, self.radius, self.radius)
        painter.fillPath(path, QBrush(self.bg_color))
        painter.end()


# ============================================================
#  密码对话框
# ============================================================
class PasswordDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setModal(True)
        self.setFixedSize(300, 200)
        self.result_ok = False

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        self.bg = RoundedWidget(radius=14, bg=PANEL, parent=self)
        outer.addWidget(self.bg)

        lay = QVBoxLayout(self.bg)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(10)

        title = QLabel("请输入管理密码")
        title.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        title.setStyleSheet(f"color:{TEXT}; background:transparent;")
        title.setAlignment(Qt.AlignCenter)
        lay.addWidget(title)

        hint = QLabel("忘记密码请联系程序设计者")
        hint.setFont(QFont("Microsoft YaHei", 8))
        hint.setStyleSheet(f"color:{MUTED}; background:transparent;")
        hint.setAlignment(Qt.AlignCenter)
        lay.addWidget(hint)

        self.edit = QLineEdit()
        self.edit.setEchoMode(QLineEdit.Password)
        self.edit.setFont(QFont("Microsoft YaHei", 12))
        self.edit.setStyleSheet(f"""
            QLineEdit {{
                background:#ffffff;
                color:{TEXT};
                border:2px solid {CHIP_BORDER};
                border-radius:8px;
                padding:8px 12px;
            }}
            QLineEdit:focus {{
                border-color:{ACCENT};
            }}
        """)
        self.edit.returnPressed.connect(self.try_ok)
        lay.addWidget(self.edit)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

        cancel = QPushButton("取消")
        cancel.setCursor(Qt.PointingHandCursor)
        cancel.setStyleSheet(f"""
            QPushButton {{
                background:{CHIP_BG}; color:{TEXT};
                border:none; border-radius:8px;
                padding:8px 18px;
                font-family:'Microsoft YaHei'; font-size:10pt; font-weight:bold;
            }}
            QPushButton:hover {{ background:#b0c4d4; }}
        """)
        cancel.clicked.connect(self.reject)

        ok = QPushButton("确定")
        ok.setCursor(Qt.PointingHandCursor)
        ok.setStyleSheet(f"""
            QPushButton {{
                background:{ACCENT}; color:#ffffff;
                border:none; border-radius:8px;
                padding:8px 18px;
                font-family:'Microsoft YaHei'; font-size:10pt; font-weight:bold;
            }}
            QPushButton:hover {{ background:{ACCENT_LT}; }}
        """)
        ok.clicked.connect(self.try_ok)

        btn_row.addStretch()
        btn_row.addWidget(cancel)
        btn_row.addWidget(ok)
        lay.addLayout(btn_row)

        self.edit.setFocus()

    def try_ok(self):
        if self.edit.text() == get_current_password():
            self.result_ok = True
            self.accept()
        else:
            QMessageBox.warning(self, "密码错误", "密码不正确，请重新输入。")
            self.edit.clear()
            self.edit.setFocus()

    def mousePressEvent(self, e):
        self._drag = e.globalPosition().toPoint()
        self._orig = self.pos()

    def mouseMoveEvent(self, e):
        if hasattr(self, "_drag"):
            delta = e.globalPosition().toPoint() - self._drag
            self.move(self._orig + delta)


# ============================================================
#  数字键盘
# ============================================================
class NumPad(QWidget):
    def __init__(self, callback, parent=None):
        super().__init__(parent)
        self.callback = callback
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Popup)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setFixedSize(180, 260)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        self.bg = RoundedWidget(radius=10, bg=ACCENT, parent=self)
        outer.addWidget(self.bg)

        grid = QGridLayout(self.bg)
        grid.setContentsMargins(6, 6, 6, 6)
        grid.setSpacing(4)

        keys = [
            ('1',0,0),('2',0,1),('3',0,2),
            ('4',1,0),('5',1,1),('6',1,2),
            ('7',2,0),('8',2,1),('9',2,2),
            ('←',3,0),('0',3,1),('C',3,2),
        ]
        for txt, r, c in keys:
            b = QPushButton(txt)
            b.setCursor(Qt.PointingHandCursor)
            b.setFixedHeight(46)
            b.setStyleSheet(f"""
                QPushButton {{
                    background:{CARD}; color:#ffffff;
                    border:none; border-radius:8px;
                    font-family:'Microsoft YaHei'; font-size:13pt; font-weight:bold;
                }}
                QPushButton:hover {{ background:#3a6fa8; }}
                QPushButton:pressed {{ background:{ACCENT_LT}; }}
            """)
            b.clicked.connect(lambda _, t=txt: self.on_press(t))
            grid.addWidget(b, r, c)

        ok = QPushButton("确定")
        ok.setCursor(Qt.PointingHandCursor)
        ok.setFixedHeight(40)
        ok.setStyleSheet(f"""
            QPushButton {{
                background:{ACCENT_LT}; color:#ffffff;
                border:none; border-radius:8px;
                font-family:'Microsoft YaHei'; font-size:11pt; font-weight:bold;
            }}
            QPushButton:hover {{ background:{ACCENT}; }}
        """)
        ok.clicked.connect(lambda: self.on_press('OK'))
        grid.addWidget(ok, 4, 0, 1, 3)

    def on_press(self, k):
        if k == 'OK':
            self.callback('OK'); self.close()
        elif k == '←':
            self.callback('DEL')
        elif k == 'C':
            self.callback('CLR')
        else:
            self.callback(k)


# ============================================================
#  主窗口
# ============================================================
class RollCallApp(QWidget):
    EXP_W, EXP_H = 360, 620
    COL_W, COL_H = 28, 74

    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setFixedSize(self.EXP_W, self.EXP_H)

        self.students = load_json(ROSTER_FILE, [])
        self.filters = {'gender': 'all', 'subjects': set(), 'groups': set()}
        self.count = 1
        self.pad_buffer = '1'
        self.drawing = False
        self._roll_timer = None
        self._pending_pool = []
        self._roll_pool = []
        self.collapsed = False
        self.saved_pos = None
        self._drag_pos = None
        self._numpad = None

        screen = QApplication.primaryScreen().availableGeometry()
        self.sw, self.sh = screen.width(), screen.height()
        self.move(self.sw - self.EXP_W - 40, 80)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        self.bg = RoundedWidget(radius=18, bg=BG, parent=self)
        outer.addWidget(self.bg)

        self._build_ui()

    # ---------------- UI ----------------
    def _build_ui(self):
        main = QVBoxLayout(self.bg)
        main.setContentsMargins(0, 0, 0, 0)
        main.setSpacing(0)

        # 标题栏
        title = QFrame()
        title.setFixedHeight(40)
        title.setStyleSheet(f"background:{PANEL}; border-top-left-radius:18px; border-top-right-radius:18px;")
        tlay = QHBoxLayout(title)
        tlay.setContentsMargins(16, 0, 10, 0)

        self.title_label = QLabel("课堂随机点名")
        self.title_label.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        self.title_label.setStyleSheet(f"color:{TEXT}; background:transparent;")
        tlay.addWidget(self.title_label)
        tlay.addStretch()

        collapse_btn = QPushButton("◀ 收起")
        collapse_btn.setCursor(Qt.PointingHandCursor)
        collapse_btn.setStyleSheet(f"""
            QPushButton {{
                color:{MUTED}; background:transparent;
                border:none; padding:4px 8px;
                font-family:'Microsoft YaHei'; font-size:9pt;
            }}
            QPushButton:hover {{ color:{ACCENT}; }}
        """)
        collapse_btn.clicked.connect(self.collapse)
        tlay.addWidget(collapse_btn)

        close_btn = QPushButton("✕")
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setFixedSize(28, 28)
        close_btn.setStyleSheet(f"""
            QPushButton {{
                color:{MUTED}; background:transparent;
                border:none;
                font-family:'Microsoft YaHei'; font-size:11pt;
            }}
            QPushButton:hover {{ color:{DANGER}; }}
        """)
        close_btn.clicked.connect(self.close)
        tlay.addWidget(close_btn)

        main.addWidget(title)

        # 内容区
        content = QWidget()
        content.setStyleSheet("background:transparent;")
        clay = QVBoxLayout(content)
        clay.setContentsMargins(14, 14, 14, 14)
        clay.setSpacing(10)

        # 结果卡片
        self.result_card = QFrame()
        self.result_card.setFixedHeight(120)
        self.result_card.setStyleSheet(f"""
            QFrame {{
                background:{CARD};
                border-radius:14px;
            }}
        """)
        rlay = QVBoxLayout(self.result_card)
        rlay.setContentsMargins(0, 0, 0, 0)
        self.result_label = QLabel("尚未导入花名册")
        self.result_label.setAlignment(Qt.AlignCenter)
        self.result_label.setFont(QFont("Microsoft YaHei", 22, QFont.Bold))
        self.result_label.setStyleSheet(f"color:#ffffff; background:transparent;")
        rlay.addWidget(self.result_label)
        clay.addWidget(self.result_card)

        # 性别
        clay.addWidget(self._section_label("性别"))
        self.gender_row = QHBoxLayout()
        self.gender_row.setSpacing(6)
        self.gender_row.setContentsMargins(0, 0, 0, 0)
        gw = QWidget(); gw.setStyleSheet("background:transparent;"); gw.setLayout(self.gender_row)
        clay.addWidget(gw)

        # 选科
        clay.addWidget(self._section_label("选科"))
        self.subject_row = QHBoxLayout()
        self.subject_row.setSpacing(6)
        self.subject_row.setContentsMargins(0, 0, 0, 0)
        sw = QWidget(); sw.setStyleSheet("background:transparent;"); sw.setLayout(self.subject_row)
        clay.addWidget(sw)

        # 小组
        clay.addWidget(self._section_label("小组"))
        self.group_row = QHBoxLayout()
        self.group_row.setSpacing(6)
        self.group_row.setContentsMargins(0, 0, 0, 0)
        grw = QWidget(); grw.setStyleSheet("background:transparent;"); grw.setLayout(self.group_row)
        clay.addWidget(grw)

        # 抽取设置
        clay.addWidget(self._section_label("抽取设置"))
        setrow = QHBoxLayout()
        setrow.setContentsMargins(0, 0, 0, 0)
        setrow.setSpacing(8)

        set_lbl = QLabel("每次抽取人数")
        set_lbl.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        set_lbl.setStyleSheet(f"color:{TEXT}; background:transparent;")
        setrow.addWidget(set_lbl)
        setrow.addStretch()

        num_box = QFrame()
        num_box.setFixedHeight(32)
        num_box.setStyleSheet(f"""
            QFrame {{
                background:{CHIP_BG};
                border:1px solid {CHIP_BORDER};
                border-radius:8px;
            }}
        """)
        nlay = QHBoxLayout(num_box)
        nlay.setContentsMargins(4, 0, 4, 0)
        nlay.setSpacing(2)

        self.count_input = QLineEdit("1")
        self.count_input.setFixedWidth(36)
        self.count_input.setAlignment(Qt.AlignCenter)
        self.count_input.setReadOnly(True)
        self.count_input.setCursor(Qt.PointingHandCursor)
        self.count_input.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        self.count_input.setStyleSheet(f"""
            QLineEdit {{
                background:transparent; color:{TEXT};
                border:none;
            }}
        """)
        self.count_input.mousePressEvent = self._on_count_click
        nlay.addWidget(self.count_input)

        arr_col = QVBoxLayout()
        arr_col.setContentsMargins(0, 0, 0, 0)
        arr_col.setSpacing(0)
        up = QPushButton("▲")
        up.setFixedSize(20, 15)
        up.setCursor(Qt.PointingHandCursor)
        up.setStyleSheet(f"""
            QPushButton {{
                background:transparent; color:{ACCENT};
                border:none; font-size:6pt;
            }}
            QPushButton:hover {{ color:{ACCENT_LT}; }}
        """)
        up.clicked.connect(lambda: self._nudge(+1))
        arr_col.addWidget(up)

        down = QPushButton("▼")
        down.setFixedSize(20, 15)
        down.setCursor(Qt.PointingHandCursor)
        down.setStyleSheet(f"""
            QPushButton {{
                background:transparent; color:{ACCENT};
                border:none; font-size:6pt;
            }}
            QPushButton:hover {{ color:{ACCENT_LT}; }}
        """)
        down.clicked.connect(lambda: self._nudge(-1))
        arr_col.addWidget(down)

        nlay.addLayout(arr_col)
        setrow.addWidget(num_box)
        clay.addLayout(setrow)

        # 候选人数
        pool_box = QFrame()
        pool_box.setFixedHeight(38)
        pool_box.setStyleSheet(f"""
            QFrame {{
                background:{ACCENT};
                border-radius:8px;
            }}
        """)
        play = QHBoxLayout(pool_box)
        play.setContentsMargins(14, 0, 14, 0)

        pl = QLabel("当前候选")
        pl.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        pl.setStyleSheet("color:#eaf1f6; background:transparent;")
        play.addWidget(pl)
        play.addStretch()

        self.pool_count_label = QLabel("0 人")
        self.pool_count_label.setFont(QFont("Microsoft YaHei", 12, QFont.Bold))
        self.pool_count_label.setStyleSheet("color:#ffffff; background:transparent;")
        play.addWidget(self.pool_count_label)
        clay.addWidget(pool_box)

        # 抽取按钮
        self.draw_btn = QPushButton("开 始 抽 取")
        self.draw_btn.setCursor(Qt.PointingHandCursor)
        self.draw_btn.setFixedHeight(52)
        self.draw_btn.setFont(QFont("Microsoft YaHei", 14, QFont.Bold))
        self.draw_btn.setStyleSheet(f"""
            QPushButton {{
                background:{SMOKE};
                color:#ffffff;
                border:none; border-radius:12px;
                letter-spacing:4px;
            }}
            QPushButton:hover {{ background:{SMOKE_HOVER}; }}
        """)
        self.draw_btn.clicked.connect(self.on_draw_click)
        clay.addWidget(self.draw_btn)

        # 底部
        bottom = QHBoxLayout()
        bottom.setContentsMargins(2, 0, 2, 0)
        self.status_label = QLabel("")
        self.status_label.setFont(QFont("Microsoft YaHei", 9))
        self.status_label.setStyleSheet(f"color:{MUTED}; background:transparent;")
        bottom.addWidget(self.status_label)
        bottom.addStretch()

        menu_btn = QPushButton("⚙ 设置")
        menu_btn.setCursor(Qt.PointingHandCursor)
        menu_btn.setStyleSheet(f"""
            QPushButton {{
                color:{MUTED}; background:transparent;
                border:none; padding:4px 8px;
                font-family:'Microsoft YaHei'; font-size:9pt;
            }}
            QPushButton:hover {{ color:{ACCENT}; }}
        """)
        menu_btn.clicked.connect(self.show_menu)
        bottom.addWidget(menu_btn)
        clay.addLayout(bottom)

        main.addWidget(content, 1)

        self._refresh_filters()
        self._update_status()

    def _section_label(self, text):
        lbl = QLabel(text)
        lbl.setFont(QFont("Microsoft YaHei", 9, QFont.Bold))
        lbl.setStyleSheet(f"color:{ACCENT}; background:transparent;")
        return lbl

    def _clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()

    # ★★★ 筛选按钮：选中态明显不同 ★★★
    def _make_chip(self, text, on, command):
        btn = QPushButton(("✓ " if on else "") + text)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setFixedHeight(28)
        if on:
            # 选中：深海蓝底 + 纯白字 + 深蓝边框
            btn.setStyleSheet(f"""
                QPushButton {{
                    background:{CHIP_ON_BG}; color:{CHIP_ON_FG};
                    border:2px solid {CHIP_ON_BORDER};
                    border-radius:14px;
                    padding:0 12px;
                    font-family:'Microsoft YaHei'; font-size:9pt; font-weight:bold;
                }}
                QPushButton:hover {{ background:{ACCENT_LT}; }}
            """)
        else:
            # 未选中：极浅灰蓝底 + 淡蓝字
            btn.setStyleSheet(f"""
                QPushButton {{
                    background:{CHIP_BG}; color:{CHIP_TEXT};
                    border:1px solid {CHIP_BORDER};
                    border-radius:14px;
                    padding:0 14px;
                    font-family:'Microsoft YaHei'; font-size:9pt; font-weight:bold;
                }}
                QPushButton:hover {{ background:#cddde8; color:#5a7d9e; }}
            """)
        btn.clicked.connect(command)
        return btn

    def _refresh_filters(self):
        self._clear_layout(self.gender_row)
        opts = [('all', '全部')]
        genders = {s.get('gender', '') for s in self.students} - {''}
        if '男' in genders: opts.append(('男', '男生'))
        if '女' in genders: opts.append(('女', '女生'))
        for v, label in opts:
            on = self.filters['gender'] == v
            def cmd(v=v):
                self.filters['gender'] = v
                self._refresh_filters(); self._update_status()
            self.gender_row.addWidget(self._make_chip(label, on, cmd))
        self.gender_row.addStretch()

        self._clear_layout(self.subject_row)
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
                self.subject_row.addWidget(self._make_chip(c, on, cmd))
        else:
            lbl = QLabel("（无选科信息）")
            lbl.setStyleSheet(f"color:{MUTED}; background:transparent;")
            self.subject_row.addWidget(lbl)
        self.subject_row.addStretch()

        self._clear_layout(self.group_row)
        groups = sorted({s['group'] for s in self.students if s.get('group')},
                        key=lambda x: (float(x) if x.replace('.', '', 1).isdigit() else 999, x))
        if groups:
            for g in groups:
                on = g in self.filters['groups']
                def cmd(g=g):
                    if g in self.filters['groups']: self.filters['groups'].discard(g)
                    else: self.filters['groups'].add(g)
                    self._refresh_filters(); self._update_status()
                self.group_row.addWidget(self._make_chip(g, on, cmd))
        else:
            lbl = QLabel("（无小组信息）")
            lbl.setStyleSheet(f"color:{MUTED}; background:transparent;")
            self.group_row.addWidget(lbl)
        self.group_row.addStretch()

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
        self.pool_count_label.setText(f"{len(pool)} 人")
        self.status_label.setText(f"花名册 {len(self.students)} 人")

    # ---------------- 数字框 ----------------
    def _nudge(self, delta):
        try: v = int(self.pad_buffer)
        except Exception: v = 1
        v = max(1, min(60, v + delta))
        self.pad_buffer = str(v)
        self._render_pad()

    def _render_pad(self):
        self.count_input.setText(self.pad_buffer)
        try:
            v = int(self.pad_buffer)
            self.count = max(1, min(60, v))
        except Exception:
            self.count = 1

    def _on_count_click(self, e):
        if self.drawing:
            QMessageBox.information(self, "提示", "抽奖中不能改人数")
            return
        self.pad_buffer = self.count_input.text() or '1'
        self._numpad = NumPad(self._on_numpad_press, self)
        pos = self.mapToGlobal(self.count_input.pos())
        self._numpad.move(pos.x() - 130, pos.y() + 40)
        self._numpad.show()

    def _on_numpad_press(self, k):
        if k == 'OK':
            self._render_pad()
        elif k == 'DEL':
            self.pad_buffer = self.pad_buffer[:-1]
            self.count_input.setText(self.pad_buffer)
        elif k == 'CLR':
            self.pad_buffer = ''
            self.count_input.setText('')
        else:
            if self.pad_buffer == '0': self.pad_buffer = ''
            if len(self.pad_buffer) < 2:
                self.pad_buffer += k
            self.count_input.setText(self.pad_buffer)
        try:
            v = int(self.pad_buffer)
            self.count = max(1, min(60, v))
        except Exception:
            self.count = 1

    # ---------------- 抽取 ----------------
    def on_draw_click(self):
        if self.drawing: self._stop_rolling()
        else: self._start_rolling()

    def _start_rolling(self):
        if not self.students:
            QMessageBox.information(self, "提示", "请先从「⚙ 设置」中导入花名册")
            return
        self._roll_pool = self._get_pool()
        self._pending_pool = [s for s in self._roll_pool if not is_blocked(s)]
        if not self._roll_pool:
            QMessageBox.information(self, "提示", "当前筛选条件下没有学生"); return
        if not self._pending_pool:
            QMessageBox.information(self, "提示", "当前筛选条件下没有可抽取的学生"); return
        self.drawing = True
        self.draw_btn.setText("停 止 抽 取")
        self.draw_btn.setStyleSheet(f"""
            QPushButton {{
                background:{DANGER};
                color:#ffffff;
                border:none; border-radius:12px;
                letter-spacing:4px;
            }}
            QPushButton:hover {{ background:#a82020; }}
        """)
        # 启动计时器
        if self._roll_timer is None:
            self._roll_timer = QTimer(self)
            self._roll_timer.timeout.connect(self._tick)
        self._roll_timer.start(70)

    def _fmt_names(self, picks):
        n = len(picks)
        if n == 1: return picks[0]['name'], 22
        if n <= 3: return '  '.join(p['name'] for p in picks), 15
        if n <= 6: return '  '.join(p['name'] for p in picks), 11
        return '  '.join(p['name'] for p in picks), 9

    def _tick(self):
        if not self.drawing:
            if self._roll_timer:
                self._roll_timer.stop()
            return
        n = min(self.count, len(self._roll_pool))
        picks = random.sample(self._roll_pool, n)
        text, fs = self._fmt_names(picks)
        self.result_label.setText(text)
        self.result_label.setFont(QFont("Microsoft YaHei", fs, QFont.Bold))

    def _stop_rolling(self):
        self.drawing = False
        if self._roll_timer:
            self._roll_timer.stop()
        self.draw_btn.setText("开 始 抽 取")
        self.draw_btn.setStyleSheet(f"""
            QPushButton {{
                background:{SMOKE};
                color:#ffffff;
                border:none; border-radius:12px;
                letter-spacing:4px;
            }}
            QPushButton:hover {{ background:{SMOKE_HOVER}; }}
        """)
        n = max(1, min(self.count, len(self._pending_pool)))
        picks = random.sample(self._pending_pool, n)
        text, fs = self._fmt_names(picks)
        self.result_label.setText(text)
        self.result_label.setFont(QFont("Microsoft YaHei", fs, QFont.Bold))

    # ---------------- 菜单 ----------------
    def show_menu(self):
        menu = QMenu(self)
        menu.setStyleSheet(f"""
            QMenu {{
                background:{PANEL};
                color:{TEXT};
                border:1px solid {CHIP_BORDER};
                border-radius:8px;
                padding:6px;
                font-family:'Microsoft YaHei'; font-size:10pt;
            }}
            QMenu::item {{
                padding:8px 24px 8px 12px;
                border-radius:6px;
            }}
            QMenu::item:selected {{
                background:{ACCENT}; color:#ffffff;
            }}
            QMenu::separator {{
                height:1px; background:{CHIP_BORDER}; margin:4px 8px;
            }}
        """)

        menu.addAction("📥  导入花名册 (.xlsx / .csv)", self.import_roster)
        menu.addAction("🗑  清空花名册（需密码）", self.clear_roster)
        menu.addSeparator()

        sub = menu.addMenu(f"🎯  每次抽取人数（当前 {self.count}）")
        for n in range(1, 6):
            a = sub.addAction(f"每次抽 {n} 人")
            a.triggered.connect(lambda _, nn=n: self.set_count(nn))

        menu.addSeparator()

        auto_action = QAction("开机自动启动", self)
        auto_action.setCheckable(True)
        if sys.platform == 'win32' and get_exe_path():
            auto_action.setChecked(is_autostart_enabled())
            auto_action.setEnabled(True)
            auto_action.triggered.connect(self.toggle_autostart_action)
        else:
            auto_action.setChecked(False)
            auto_action.setEnabled(False)
            auto_action.setText("开机自动启动（仅 exe 版可用）")
        menu.addAction(auto_action)

        menu.addSeparator()
        menu.addAction("ℹ  使用说明", self.show_help)
        menu.addAction("✕  退出程序", self.close)

        menu.exec(self.mapToGlobal(self.bg.pos() + QPoint(self.EXP_W - 60, self.EXP_H - 30)))

    def toggle_autostart_action(self):
        current = is_autostart_enabled()
        new_state = not current
        ok = set_autostart(new_state)
        if ok:
            QMessageBox.information(self, "提示",
                "已开启开机自启" if new_state else "已关闭开机自启")
        else:
            QMessageBox.warning(self, "提示", "设置失败，请确认用的是 exe 版本。")

    def set_count(self, n):
        self.count = n; self.pad_buffer = str(n); self._render_pad()

    # ---------------- 导入 / 清空 ----------------
    def import_roster(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择花名册文件", "",
            "表格文件 (*.xlsx *.xls *.csv *.txt *.tsv);;所有文件 (*.*)")
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
                QMessageBox.warning(self, "提示", "未解析到有效的学生数据"); return
            self.students = students
            save_json(ROSTER_FILE, students)
            self.filters = {'gender': 'all', 'subjects': set(), 'groups': set()}
            self._refresh_filters(); self._update_status()
            self.result_label.setText(f"已导入 {len(students)} 人")
        except Exception as e:
            QMessageBox.critical(self, "导入失败", str(e))

    def clear_roster(self):
        if not self.students:
            QMessageBox.information(self, "提示", "当前没有花名册"); return

        dlg = PasswordDialog(self)
        dlg.move(self.geometry().center() - QPoint(150, 100))
        if not dlg.exec() or not dlg.result_ok:
            return

        self.students = []
        save_json(ROSTER_FILE, [])
        self.filters = {'gender': 'all', 'subjects': set(), 'groups': set()}
        self._refresh_filters(); self._update_status()
        self.result_label.setText("花名册已清空")
        QMessageBox.information(self, "已清空", "花名册已清空。")

    def show_help(self):
        QMessageBox.information(self, "使用说明",
            "【课堂随机点名 · 悬浮窗版】\n\n"
            "· 点右上「◀ 收起」→ 变成屏幕边缘的小箭头\n"
            "· 鼠标移到箭头上 → 自动展开\n"
            "· 点「开始抽取」滚动，再点一次停止\n\n"
            "【导入花名册】\n"
            "点「⚙ 设置 → 导入花名册」\n"
            "表头四列：姓名 | 性别 | 选科 | 小组\n\n"
            "【清空花名册】\n"
            "需要输入管理密码\n\n"
            "花名册保存在：\n" + str(ROSTER_FILE))

    # ---------------- 拖动 / 收起 ----------------
    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton and e.position().y() < 40:
            self._drag_pos = e.globalPosition().toPoint() - self.frameGeometry().topLeft()
            e.accept()

    def mouseMoveEvent(self, e):
        if self._drag_pos and e.buttons() & Qt.LeftButton:
            self.move(e.globalPosition().toPoint() - self._drag_pos)
            e.accept()

    def mouseReleaseEvent(self, e):
        self._drag_pos = None

    # ★★★ 关闭时彻底释放 ★★★
    def closeEvent(self, event):
        self.drawing = False
        if self._roll_timer:
            try:
                self._roll_timer.stop()
            except Exception:
                pass
            self._roll_timer = None

        if self._numpad:
            try:
                self._numpad.close()
                self._numpad.deleteLater()
            except Exception:
                pass
            self._numpad = None

        self.hide()
        QApplication.processEvents()
        event.accept()
        QTimer.singleShot(100, QApplication.quit)

    def collapse(self):
        if self.collapsed: return
        self.collapsed = True
        self.saved_pos = self.pos()
        x = self.x()
        if x + self.EXP_W // 2 > self.sw // 2:
            new_x = self.sw - self.COL_W
        else:
            new_x = 0
        y = max(50, min(self.y(), self.sh - self.COL_H - 50))
        self.setFixedSize(self.COL_W, self.COL_H)
        self.setWindowOpacity(0.55)
        self.move(new_x, y)
        self._show_arrow(True)

    def _show_arrow(self, show):
        if show:
            old = self.bg.layout()
            if old:
                while old.count():
                    it = old.takeAt(0)
                    w = it.widget()
                    if w: w.deleteLater()
            lay = QVBoxLayout(self.bg)
            lay.setContentsMargins(0, 0, 0, 0)
            arrow = QLabel("◀" if self.x() > self.sw // 2 else "▶")
            arrow.setAlignment(Qt.AlignCenter)
            arrow.setFont(QFont("Microsoft YaHei", 16, QFont.Bold))
            arrow.setStyleSheet("color:#ffffff; background:transparent;")
            lay.addWidget(arrow)
            self.bg.bg_color = QColor(ACCENT_LT)
            self.bg.radius = 6
            self.bg.installEventFilter(self)
            arrow.installEventFilter(self)
            self.bg.update()
        else:
            self.bg.bg_color = QColor(BG)
            self.bg.radius = 18
            old = self.bg.layout()
            if old:
                while old.count():
                    it = old.takeAt(0)
                    w = it.widget()
                    if w: w.deleteLater()
            self._build_ui()

    def eventFilter(self, obj, e):
        if self.collapsed:
            if e.type() == QEvent.Enter:
                QTimer.singleShot(350, self.expand)
            elif e.type() == QEvent.MouseButtonPress:
                self.expand()
                return True
        return super().eventFilter(obj, e)

    def expand(self):
        if not self.collapsed: return
        self.collapsed = False
        self.setWindowOpacity(1.0)
        self.setFixedSize(self.EXP_W, self.EXP_H)
        if self.saved_pos:
            x, y = self.saved_pos.x(), self.saved_pos.y()
        else:
            x = self.sw - self.EXP_W - 40; y = 80
        x = max(0, min(x, self.sw - self.EXP_W))
        y = max(0, min(y, self.sh - self.EXP_H))
        self.move(x, y)
        self._show_arrow(False)


# ============================================================
#  入口
# ============================================================
def main():
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei", 10))
    app.setQuitOnLastWindowClosed(True)
    win = RollCallApp()
    win.show()
    code = app.exec()
    sys.exit(code)


if __name__ == '__main__':
    main()
