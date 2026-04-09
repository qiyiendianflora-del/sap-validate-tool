#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAP凭证校验工具 v2.1 - PyQt5 GUI（橙色主题）
打包命令：pyinstaller --onefile --windowed --name SAP凭证校验工具 validate_gui.py
"""
import sys, os, json
from datetime import datetime
from decimal import Decimal, InvalidOperation
from collections import defaultdict

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem,
    QHeaderView, QProgressBar, QMessageBox, QGroupBox, QAbstractItemView,
    QTabBar, QFrame, QGraphicsDropShadowEffect
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QDragEnterEvent, QDropEvent

import openpyxl
from openpyxl.styles import PatternFill, Font as XlFont, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 常量 & 配置
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
APP_NAME = "SAP凭证校验工具"
VERSION  = "v2.1"
DATA_ROW = 4

CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".sap_validate_config.json")
CACHE_FILE  = os.path.join(os.path.expanduser("~"), ".sap_validate_cache.json")

DEBIT_CODES  = {'01', '21', '40'}
CREDIT_CODES = {'11', '31', '50'}
ALL_CODES    = DEBIT_CODES | CREDIT_CODES
RECON_CODES  = {'D': {'01','11'}, 'K': {'21','31'}, None: {'40','50'}, '': {'40','50'}}
RECON_LABELS = {'D': '客户统驭', 'K': '供应商统驭', None: '普通科目', '': '普通科目'}

HARDCODED_RULES = [(1002010001, 1002019999, [('AD','现金流量码'),('X','利润中心')])]

FORBIDDEN_COLS_NON6 = [
    ('AF','客户'),('AG','供应商'),('AH','生产（物料）'),('AI','流量类型'),
    ('AJ','渠道'),('AK','一级费用'),('AL','二级费用'),('AM','城市'),
    ('AN','项目'),('AO','销售部门'),('AP','产品'),('AQ','提报人'),
    ('AR','提报部门'),('AS','报销人'),('AT','销售大区'),('AU','销售员'),
]

A_IDX  = column_index_from_string('A') - 1
B_IDX  = column_index_from_string('B') - 1
J_IDX  = column_index_from_string('J') - 1
M_IDX  = column_index_from_string('M') - 1
O_IDX  = column_index_from_string('O') - 1
AK_IDX = column_index_from_string('AK') - 1
AL_IDX = column_index_from_string('AL') - 1

# 主题色
C_PRIMARY='#E65100'; C_PRIMARY_L='#FF6D00'; C_PRIMARY_XL='#FFF3E0'
C_ACCENT='#FF9100'; C_BG='#FAFAFA'; C_CARD='#FFFFFF'
C_TEXT='#212121'; C_TEXT_SEC='#757575'
C_SUCCESS='#2E7D32'; C_ERROR='#C62828'; C_WARN='#E65100'; C_BORDER='#E0E0E0'

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 配置持久化
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE,'r',encoding='utf-8') as f: return json.load(f)
        except: pass
    return {}

def save_config(cfg):
    try:
        with open(CONFIG_FILE,'w',encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except: pass

def load_mapping_cache():
    """从本地缓存加载费用mapping"""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE,'r',encoding='utf-8') as f: return json.load(f)
        except: pass
    return None

def save_mapping_cache(data):
    """缓存费用mapping到本地JSON"""
    try:
        with open(CACHE_FILE,'w',encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except: pass

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 校验逻辑
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def col_letter_to_idx(letter):
    return column_index_from_string(letter) - 1

def get_hardcoded_required(acct_str):
    try: acct_int = int(acct_str)
    except: return None
    for s,e,cols in HARDCODED_RULES:
        if s <= acct_int <= e: return cols
    return None

def load_rule_table(rule_file):
    wb = openpyxl.load_workbook(rule_file); ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    field_names, col_ids = rows[0], rows[1]
    rule_map, recon_map = {}, {}
    for row in rows[2:]:
        acct = row[0]
        if acct is None: continue
        acct_str = str(int(acct)) if isinstance(acct, float) else str(acct).strip()
        recon = row[2] if len(row)>2 else None
        if recon: recon = str(recon).strip()
        recon_map[acct_str] = recon if recon else None
        req = []
        for fn, cid, val in zip(field_names, col_ids, row):
            if cid and cid != 'M' and val == '√': req.append((cid, fn))
        rule_map[acct_str] = req
    return rule_map, recon_map

def load_mapping_table(mapping_file):
    """加载费用mapping规则表，返回 {科目: set((一级,二级), ...)}"""
    wb = openpyxl.load_workbook(mapping_file); ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    mapping = defaultdict(set)
    for row in rows[1:]:  # 跳过表头
        if row[0] is None: continue
        acct = str(int(row[0])) if isinstance(row[0], float) else str(row[0]).strip()
        fy1 = str(int(row[1])) if isinstance(row[1], float) else str(row[1]).strip() if row[1] else ''
        fy2 = str(int(row[2])) if isinstance(row[2], float) else str(row[2]).strip() if row[2] else ''
        mapping[acct].add((fy1, fy2))
    # 转成可JSON序列化格式
    return {k: list(v) for k, v in mapping.items()}

def safe_decimal(val):
    if val is None: return None
    try: return Decimal(str(val).strip())
    except: return None

def check_decimal_places(val):
    d = safe_decimal(val)
    if d is None: return True
    s, digits, exp = d.as_tuple()
    return abs(exp) <= 2 if exp < 0 else True

def _norm(val):
    """标准化为字符串，数字去小数"""
    if val is None: return ''
    if isinstance(val, float): return str(int(val))
    return str(val).strip()

def run_validate(b_file, rule_file, mapping_data=None, progress_cb=None):
    rule_map, recon_map = load_rule_table(rule_file)
    wb_b = openpyxl.load_workbook(b_file); ws_b = wb_b.active
    max_col = ws_b.max_column
    total = ws_b.max_row - DATA_ROW + 1
    if total <= 0:
        return {}, [], openpyxl.Workbook(), ws_b, rule_map, 0

    # 费用mapping → set 格式
    mapping_sets = {}
    if mapping_data:
        for acct, combos in mapping_data.items():
            mapping_sets[acct] = set(tuple(c) for c in combos)

    all_rows = {}
    for r in range(DATA_ROW, ws_b.max_row+1):
        all_rows[r] = [ws_b.cell(row=r, column=c+1).value for c in range(max_col)]

    errors = defaultdict(list)
    warnings = []

    for idx, (excel_row, rv) in enumerate(all_rows.items()):
        if progress_cb and idx % 50 == 0:
            progress_cb(int(idx/total*55))

        acct_val = rv[M_IDX] if M_IDX < len(rv) else None
        if acct_val is None: continue
        acct_str = _norm(acct_val)

        # ── 校验一：必输项 ──
        hc = get_hardcoded_required(acct_str)
        is_hc = hc is not None
        if is_hc:
            req_cols = hc
        elif acct_str in rule_map:
            req_cols = rule_map[acct_str]
        else:
            warnings.append((excel_row, acct_str, '科目不在规则表'))
            req_cols = []

        for cl, fn in req_cols:
            ci = col_letter_to_idx(cl)
            v = rv[ci] if ci < len(rv) else None
            if v is None or str(v).strip() == '':
                errors[excel_row].append(('必输项为空', [cl], f'缺失必输项：{cl}({fn})'))

        # ── 校验二a：O列金额 ──
        o_val = rv[O_IDX] if O_IDX < len(rv) else None
        if o_val is None or str(o_val).strip() == '':
            errors[excel_row].append(('金额为空', ['O'], '凭证货币金额(O列)为空'))
        elif not check_decimal_places(o_val):
            errors[excel_row].append(('金额格式错误', ['O'], f'金额小数位超过2位：{o_val}'))

        # ── 校验三：统驭科目vs记账码 ──
        j_val = rv[J_IDX] if J_IDX < len(rv) else None
        j_str = _norm(j_val)
        if len(j_str) == 1: j_str = '0' + j_str

        if not is_hc and acct_str in recon_map:
            rt = recon_map[acct_str]
            rk = ''
            if rt:
                ru = str(rt).strip().upper()
                if '客户' in str(rt) or ru == 'D': rk = 'D'
                elif '供应' in str(rt) or ru == 'K': rk = 'K'
            allowed = RECON_CODES.get(rk, set())
            label = RECON_LABELS.get(rk, '普通科目')
            if j_str and j_str not in allowed:
                errors[excel_row].append(('记账码不匹配', ['J'],
                    f'统驭科目为{label}({rk or "空"})，应为{"/".join(sorted(allowed))}，实际{j_str}'))

        if j_str and j_str not in ALL_CODES:
            warnings.append((excel_row, j_str, '记账码不在标准范围'))

        # ── 校验四：非6开头禁填AF~AU ──
        if not acct_str.startswith('6'):
            for cl, fn in FORBIDDEN_COLS_NON6:
                ci = col_letter_to_idx(cl)
                v = rv[ci] if ci < len(rv) else None
                if v is not None and str(v).strip() != '':
                    errors[excel_row].append(('禁填字段有值', [cl],
                        f'非6开头科目，{cl}({fn})不允许有值'))

        # ── 校验五：费用类别组合校验 ──
        if mapping_sets and acct_str in mapping_sets:
            ak_val = _norm(rv[AK_IDX] if AK_IDX < len(rv) else None)
            al_val = _norm(rv[AL_IDX] if AL_IDX < len(rv) else None)
            # 只有AK/AL都有值时才校验组合
            if ak_val and al_val:
                if (ak_val, al_val) not in mapping_sets[acct_str]:
                    errors[excel_row].append(('费用类别不匹配', ['AK','AL'],
                        f'科目{acct_str}下，一级费用{ak_val}+二级费用{al_val}组合不在配置表中'))
            elif ak_val and not al_val:
                # AK有值但AL为空，检查是否必输（查规则表）
                # 如果规则表要求AL必输，校验一已经报了，这里不重复
                pass
            elif not ak_val and al_val:
                # AL有值但AK为空
                pass

    if progress_cb: progress_cb(65)

    # ── 校验二d：借贷平衡 ──
    groups = defaultdict(list)
    for r, rv in all_rows.items():
        a = rv[A_IDX] if A_IDX < len(rv) else None
        if a is None: continue
        a_s = _norm(a)
        b_s = _norm(rv[B_IDX] if B_IDX < len(rv) else None)
        groups[(a_s, b_s)].append(r)

    if progress_cb: progress_cb(75)

    for (lid, bk), rlist in groups.items():
        d_sum, c_sum = Decimal('0'), Decimal('0')
        for r in rlist:
            rv = all_rows[r]
            ov = rv[O_IDX] if O_IDX < len(rv) else None
            jv = rv[J_IDX] if J_IDX < len(rv) else None
            js = _norm(jv)
            if len(js)==1: js='0'+js
            amt = safe_decimal(ov)
            if amt is None: continue
            if js in DEBIT_CODES: d_sum += amt
            elif js in CREDIT_CODES: c_sum += amt
        diff = d_sum - c_sum
        if diff != 0:
            msg = f'借贷不平衡：凭证{lid}/公司{bk}，借方{d_sum} 贷方{c_sum} 差额{diff}'
            for r in rlist:
                errors[r].append(('借贷不平衡', [], msg))

    if progress_cb: progress_cb(85)
    out_wb = build_report(errors, warnings, wb_b, ws_b, rule_map, total)
    if progress_cb: progress_cb(100)
    return errors, warnings, out_wb, ws_b, rule_map, total


def build_report(errors, warnings, wb_b, ws_b, rule_map, total):
    RED=PatternFill(start_color="FF4444",end_color="FF4444",fill_type="solid")
    YELLOW=PatternFill(start_color="FFF2CC",end_color="FFF2CC",fill_type="solid")
    ORANGE=PatternFill(start_color="FFE0B2",end_color="FFE0B2",fill_type="solid")
    HDR=PatternFill(start_color="E65100",end_color="E65100",fill_type="solid")
    GRAY=PatternFill(start_color="F2F2F2",end_color="F2F2F2",fill_type="solid")
    WHITE=PatternFill(start_color="FFFFFF",end_color="FFFFFF",fill_type="solid")
    thin=Side(style="thin",color="DDDDDD")
    bdr=Border(left=thin,right=thin,top=thin,bottom=thin)

    out_wb=openpyxl.Workbook(); ws_out=out_wb.active; ws_out.title="校验结果"

    error_cells=set()
    for r,el in errors.items():
        for _,cols,_ in el:
            for cl in cols: error_cells.add((r, cl))
    warn_acct = {w[0]:w[1] for w in warnings if w[2]=='科目不在规则表'}

    mc=ws_b.max_column; ec=mc+1
    for r in range(1, ws_b.max_row+1):
        for c in range(1, mc+1):
            v=ws_b.cell(row=r,column=c).value
            d=ws_out.cell(row=r,column=c,value=v)
            d.border=bdr; d.alignment=Alignment(vertical="center")
            cl=get_column_letter(c)
            if r<=3:
                d.fill=HDR if r==1 else GRAY
                d.font=XlFont(name="微软雅黑",bold=(r==1),color="FFFFFF" if r==1 else "333333",size=10)
            elif r in errors:
                if (r,cl) in error_cells:
                    d.fill=RED; d.font=XlFont(name="微软雅黑",color="FFFFFF",size=10)
                else:
                    d.fill=YELLOW; d.font=XlFont(name="微软雅黑",size=10)
            elif r in warn_acct:
                d.fill=ORANGE; d.font=XlFont(name="微软雅黑",size=10)
            else:
                d.fill=WHITE; d.font=XlFont(name="微软雅黑",size=10)

        if r==1:
            c=ws_out.cell(row=1,column=ec,value="错误说明")
            c.fill=HDR;c.font=XlFont(name="微软雅黑",bold=True,color="FFFFFF",size=10);c.border=bdr
        elif r==2:
            c=ws_out.cell(row=2,column=ec,value="VALIDATE_MSG")
            c.fill=GRAY;c.font=XlFont(name="微软雅黑",size=10);c.border=bdr
        if r in errors:
            msgs=[e[2] for e in errors[r]]
            c=ws_out.cell(row=r,column=ec,value=" | ".join(msgs))
            c.fill=YELLOW;c.font=XlFont(name="微软雅黑",color="CC0000",size=10);c.border=bdr
        elif r in warn_acct:
            c=ws_out.cell(row=r,column=ec,value=f"⚠️ 科目{warn_acct[r]}不在规则表中")
            c.fill=ORANGE;c.font=XlFont(name="微软雅黑",color="885500",size=10);c.border=bdr

    for c in range(1,mc+1): ws_out.column_dimensions[get_column_letter(c)].width=14
    ws_out.column_dimensions[get_column_letter(ec)].width=65
    ws_out.freeze_panes=f"A{DATA_ROW}"

    # Sheet2
    ws2=out_wb.create_sheet("错误汇总")
    for i,h in enumerate(["行号","总账科目","错误类型","错误详情"],1):
        c=ws2.cell(row=1,column=i,value=h)
        c.fill=HDR;c.font=XlFont(name="微软雅黑",bold=True,color="FFFFFF",size=11)
        c.alignment=Alignment(horizontal="center",vertical="center");c.border=bdr
    ws2.column_dimensions['A'].width=8;ws2.column_dimensions['B'].width=16
    ws2.column_dimensions['C'].width=18;ws2.column_dimensions['D'].width=65

    ri=2
    for r in sorted(errors.keys()):
        rv=[ws_b.cell(row=r,column=c+1).value for c in range(ws_b.max_column)]
        av=rv[M_IDX] if M_IDX<len(rv) else ''
        a_s=_norm(av)
        for et,cols,msg in errors[r]:
            for i,v in enumerate([r,a_s,f"❌ {et}",msg],1):
                c=ws2.cell(row=ri,column=i,value=v)
                c.fill=PatternFill(start_color="FFF2CC",end_color="FFF2CC",fill_type="solid")
                c.font=XlFont(name="微软雅黑",size=10)
                c.alignment=Alignment(vertical="center",wrap_text=(i==4));c.border=bdr
            ri+=1
    for w in warnings:
        if w[2]!='科目不在规则表': continue
        for i,v in enumerate([w[0],w[1],"⚠️ 科目不在规则表","请核实"],1):
            c=ws2.cell(row=ri,column=i,value=v)
            c.fill=PatternFill(start_color="FFE0B2",end_color="FFE0B2",fill_type="solid")
            c.font=XlFont(name="微软雅黑",size=10);c.border=bdr
        ri+=1
    return out_wb


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 后台线程
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class ValidateWorker(QThread):
    progress=pyqtSignal(int); finished=pyqtSignal(object); error=pyqtSignal(str)
    def __init__(self,b,r,m=None):
        super().__init__(); self.b=b; self.r=r; self.m=m
    def run(self):
        try:
            e,w,wb,ws,rm,t=run_validate(self.b,self.r,self.m,lambda v:self.progress.emit(v))
            self.finished.emit({"errors":e,"warnings":w,"out_wb":wb,"total":t})
        except Exception as ex:
            import traceback; self.error.emit(traceback.format_exc())


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 自定义控件
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class StatCard(QFrame):
    def __init__(self,icon,title,value="—",color="#333"):
        super().__init__(); self.setFixedHeight(76)
        self.setStyleSheet(f"StatCard{{background:white;border-radius:10px;border:1px solid {C_BORDER};}}")
        sh=QGraphicsDropShadowEffect();sh.setBlurRadius(12);sh.setOffset(0,2);sh.setColor(QColor(0,0,0,25))
        self.setGraphicsEffect(sh)
        ly=QVBoxLayout(self);ly.setContentsMargins(14,8,14,8);ly.setSpacing(2)
        top=QLabel(f"{icon}  {title}");top.setStyleSheet(f"color:{C_TEXT_SEC};font-size:11px;border:none;")
        ly.addWidget(top)
        self.vl=QLabel(str(value));self.vl.setStyleSheet(f"color:{color};font-size:24px;font-weight:bold;border:none;")
        ly.addWidget(self.vl)
    def set_value(self,v): self.vl.setText(str(v))


class DropArea(QLabel):
    file_dropped=pyqtSignal(str)
    def __init__(self,text="将文件拖到此处 或 点击选择"):
        super().__init__(text); self.setAcceptDrops(True); self.setAlignment(Qt.AlignCenter)
        self.setFixedHeight(46); self._idle()
    def _idle(self):
        self.setStyleSheet(f"QLabel{{border:2px dashed {C_ACCENT};border-radius:8px;color:{C_TEXT_SEC};background:{C_PRIMARY_XL};font-size:12px;}}")
    def _loaded(self,fn):
        self.setText(f"📄 {fn}")
        self.setStyleSheet(f"QLabel{{border:2px solid {C_SUCCESS};border-radius:8px;color:{C_SUCCESS};background:#E8F5E9;font-size:12px;font-weight:bold;}}")
    def dragEnterEvent(self,e):
        if e.mimeData().hasUrls(): e.acceptProposedAction()
    def dragLeaveEvent(self,e): self._idle()
    def dropEvent(self,e):
        urls=e.mimeData().urls()
        if urls:
            p=urls[0].toLocalFile()
            if p.lower().endswith(('.xlsx','.xls')):
                self._loaded(os.path.basename(p)); self.file_dropped.emit(p)
            else:
                QMessageBox.warning(self,"格式错误","请拖入 .xlsx 或 .xls 文件"); self._idle()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 主窗口
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.b_file=""; self.rule_file=""; self.mapping_file=""
        self.mapping_data=None; self.result=None; self.all_table_rows=[]
        self._init_ui(); self._load_saved()

    def _load_saved(self):
        cfg=load_config()
        rp=cfg.get("rule_file","")
        if rp and os.path.exists(rp):
            self.rule_file=rp; self.r_drop._loaded(os.path.basename(rp))
        mp=cfg.get("mapping_file","")
        if mp and os.path.exists(mp):
            self.mapping_file=mp; self.m_drop._loaded(os.path.basename(mp))
        # 尝试加载缓存的mapping数据
        cached=load_mapping_cache()
        if cached:
            self.mapping_data=cached
            cnt=sum(len(v) for v in cached.values())
            self.m_info.setText(f"✅ 已加载 {len(cached)} 个科目 {cnt} 条规则")
            self.m_info.setStyleSheet(f"color:{C_SUCCESS};font-size:10px;padding:2px 0;")
        bp=cfg.get("b_file","")
        if bp and os.path.exists(bp):
            self.b_file=bp; self.b_drop._loaded(os.path.basename(bp))

    def _save_paths(self):
        cfg=load_config()
        if self.rule_file: cfg["rule_file"]=self.rule_file
        if self.b_file: cfg["b_file"]=self.b_file
        if self.mapping_file: cfg["mapping_file"]=self.mapping_file
        save_config(cfg)

    def _init_ui(self):
        self.setWindowTitle(f"{APP_NAME}  {VERSION}")
        self.setMinimumSize(1020,740); self.resize(1150,820)
        self.setStyleSheet(f"""
            QMainWindow{{background:{C_BG};}}
            QGroupBox{{font-weight:bold;font-size:12px;color:{C_TEXT};
                border:1px solid {C_BORDER};border-radius:8px;
                margin-top:12px;padding-top:18px;background:{C_CARD};}}
            QGroupBox::title{{subcontrol-origin:margin;left:16px;padding:0 6px;}}
            QStatusBar{{background:#F0F0F0;color:{C_TEXT_SEC};font-size:11px;}}
        """)
        QApplication.setFont(QFont("微软雅黑",10))

        central=QWidget(); self.setCentralWidget(central)
        root=QVBoxLayout(central); root.setSpacing(10); root.setContentsMargins(16,12,16,12)

        # 标题栏
        hdr=QFrame(); hdr.setFixedHeight(48)
        hdr.setStyleSheet(f"QFrame{{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 {C_PRIMARY},stop:1 {C_PRIMARY_L});border-radius:10px;}}")
        hl=QHBoxLayout(hdr);hl.setContentsMargins(20,0,20,0)
        ht=QLabel(f"📋  {APP_NAME}");ht.setStyleSheet("color:white;font-size:17px;font-weight:bold;background:transparent;")
        hv=QLabel(VERSION);hv.setStyleSheet("color:rgba(255,255,255,0.7);font-size:12px;background:transparent;")
        hl.addWidget(ht);hl.addStretch();hl.addWidget(hv)
        root.addWidget(hdr)

        # 文件上传区
        ub=QGroupBox("  文件上传（支持拖拽）"); ul=QVBoxLayout(ub); ul.setSpacing(6)

        # B表
        br=QHBoxLayout(); bl=QLabel("📄 B表："); bl.setFixedWidth(90)
        self.b_drop=DropArea("将B表拖到此处"); self.b_drop.file_dropped.connect(self._on_b_drop)
        bb=QPushButton("选择");bb.setFixedWidth(60);bb.setStyleSheet(self._bs(C_PRIMARY));bb.clicked.connect(self._choose_b)
        br.addWidget(bl);br.addWidget(self.b_drop,1);br.addWidget(bb); ul.addLayout(br)

        # 规则表
        rr=QHBoxLayout(); rl=QLabel("📋 校验规则表："); rl.setFixedWidth(90)
        self.r_drop=DropArea("将规则表拖到此处（记住路径）"); self.r_drop.file_dropped.connect(self._on_r_drop)
        rb=QPushButton("选择");rb.setFixedWidth(60);rb.setStyleSheet(self._bs(C_PRIMARY));rb.clicked.connect(self._choose_rule)
        rr.addWidget(rl);rr.addWidget(self.r_drop,1);rr.addWidget(rb); ul.addLayout(rr)

        # 费用mapping表
        mr=QHBoxLayout(); ml=QLabel("📊 费用配置表："); ml.setFixedWidth(90)
        self.m_drop=DropArea("将费用mapping规则表拖到此处（记住并缓存）"); self.m_drop.file_dropped.connect(self._on_m_drop)
        mb=QPushButton("选择");mb.setFixedWidth(60);mb.setStyleSheet(self._bs(C_PRIMARY));mb.clicked.connect(self._choose_mapping)
        mr.addWidget(ml);mr.addWidget(self.m_drop,1);mr.addWidget(mb); ul.addLayout(mr)

        # 费用mapping状态信息
        self.m_info=QLabel(""); self.m_info.setStyleSheet(f"color:{C_TEXT_SEC};font-size:10px;padding:2px 0;")
        ul.addWidget(self.m_info)

        root.addWidget(ub)

        # 操作按钮
        brow=QHBoxLayout()
        self.vbtn=QPushButton("  🔍  开始校验");self.vbtn.setFixedHeight(40)
        self.vbtn.setStyleSheet(self._bs(C_SUCCESS,12,28));self.vbtn.setCursor(Qt.PointingHandCursor)
        self.vbtn.clicked.connect(self._start)

        self.dbtn=QPushButton("  ⬇️  下载校验报告");self.dbtn.setFixedHeight(40)
        self.dbtn.setStyleSheet(self._bs(C_PRIMARY,12,28));self.dbtn.setEnabled(False)
        self.dbtn.setCursor(Qt.PointingHandCursor);self.dbtn.clicked.connect(self._download)

        brow.addWidget(self.vbtn);brow.addSpacing(10);brow.addWidget(self.dbtn);brow.addStretch()
        root.addLayout(brow)

        # 进度条
        self.prog=QProgressBar();self.prog.setFixedHeight(6);self.prog.setValue(0);self.prog.setTextVisible(False)
        self.prog.setStyleSheet(f"QProgressBar{{background:#e0e0e0;border-radius:3px;}}QProgressBar::chunk{{background:{C_ACCENT};border-radius:3px;}}")
        root.addWidget(self.prog)

        # 统计卡片
        sr=QHBoxLayout();sr.setSpacing(10)
        self.ct=StatCard("📊","总行数","—",C_TEXT)
        self.co=StatCard("✅","通过","—",C_SUCCESS)
        self.ce=StatCard("❌","错误","—",C_ERROR)
        self.cw=StatCard("⚠️","警告","—",C_WARN)
        for c in [self.ct,self.co,self.ce,self.cw]: sr.addWidget(c)
        root.addLayout(sr)

        # 筛选Tab + 表格
        rb2=QGroupBox("  校验结果"); rl2=QVBoxLayout(rb2)
        self.tabs=QTabBar()
        self.tabs.setStyleSheet(f"""
            QTabBar::tab{{background:{C_CARD};border:1px solid {C_BORDER};border-bottom:none;
                border-radius:6px 6px 0 0;padding:5px 14px;margin-right:2px;font-size:11px;}}
            QTabBar::tab:selected{{background:{C_PRIMARY};color:white;font-weight:bold;}}
            QTabBar::tab:hover{{background:{C_PRIMARY_XL};}}
        """)
        for t in ["全部","必输项","借贷平衡","记账码","禁填字段","费用类别","⚠️ 警告"]:
            self.tabs.addTab(t)
        self.tabs.currentChanged.connect(self._tab_changed)
        rl2.addWidget(self.tabs)

        self.tbl=QTableWidget();self.tbl.setColumnCount(4)
        self.tbl.setHorizontalHeaderLabels(["行号","总账科目","错误类型","错误详情"])
        self.tbl.horizontalHeader().setSectionResizeMode(3,QHeaderView.Stretch)
        self.tbl.setColumnWidth(0,60);self.tbl.setColumnWidth(1,120);self.tbl.setColumnWidth(2,140)
        self.tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl.setAlternatingRowColors(True);self.tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl.setStyleSheet(f"""
            QTableWidget{{border:1px solid {C_BORDER};gridline-color:#eee;}}
            QHeaderView::section{{background:{C_PRIMARY};color:white;font-weight:bold;padding:6px;border:none;}}
            QTableWidget::item:alternate{{background:#FFF8F0;}}
        """)
        rl2.addWidget(self.tbl)
        root.addWidget(rb2,1)
        self.statusBar().showMessage(f"  {APP_NAME} {VERSION}  |  就绪")

    def _bs(self,color,sz=10,px=16):
        return f"""QPushButton{{background:{color};color:white;border-radius:6px;padding:6px {px}px;
            font-size:{sz}px;font-weight:bold;border:none;}}
            QPushButton:hover{{background:{color}dd;}}QPushButton:pressed{{background:{color}bb;}}
            QPushButton:disabled{{background:#ccc;color:#888;}}"""

    def _on_b_drop(self,p): self.b_file=p; self._save_paths()
    def _on_r_drop(self,p): self.rule_file=p; self._save_paths()
    def _on_m_drop(self,p): self._load_mapping(p)

    def _choose_b(self):
        p,_=QFileDialog.getOpenFileName(self,"选择B表","","Excel (*.xlsx *.xls)")
        if p: self.b_file=p; self.b_drop._loaded(os.path.basename(p)); self._save_paths()
    def _choose_rule(self):
        p,_=QFileDialog.getOpenFileName(self,"选择校验规则表","","Excel (*.xlsx *.xls)")
        if p: self.rule_file=p; self.r_drop._loaded(os.path.basename(p)); self._save_paths()
    def _choose_mapping(self):
        p,_=QFileDialog.getOpenFileName(self,"选择费用mapping规则表","","Excel (*.xlsx *.xls)")
        if p: self._load_mapping(p)

    def _load_mapping(self, path):
        try:
            data = load_mapping_table(path)
            self.mapping_data = data
            self.mapping_file = path
            save_mapping_cache(data)
            self._save_paths()
            self.m_drop._loaded(os.path.basename(path))
            cnt = sum(len(v) for v in data.values())
            self.m_info.setText(f"✅ 已加载并缓存：{len(data)} 个科目 {cnt} 条规则")
            self.m_info.setStyleSheet(f"color:{C_SUCCESS};font-size:10px;padding:2px 0;")
        except Exception as ex:
            QMessageBox.critical(self,"加载失败",f"费用配置表解析失败：\n{ex}")

    def _start(self):
        if not self.b_file: QMessageBox.warning(self,"提示","请先上传B表！");return
        if not self.rule_file: QMessageBox.warning(self,"提示","请先上传校验规则表！");return
        self.vbtn.setEnabled(False);self.dbtn.setEnabled(False)
        self.prog.setValue(0);self.tbl.setRowCount(0)
        for c in [self.ct,self.co,self.ce,self.cw]: c.set_value("...")
        self.statusBar().showMessage("  校验中...")
        self.worker=ValidateWorker(self.b_file,self.rule_file,self.mapping_data)
        self.worker.progress.connect(self.prog.setValue)
        self.worker.finished.connect(self._done)
        self.worker.error.connect(self._err)
        self.worker.start()

    def _done(self,result):
        self.result=result
        errs=result["errors"];warns=result["warnings"];total=result["total"]
        ec=len(errs);wc=len([w for w in warns if w[2]=='科目不在规则表'])
        ok=max(0,total-ec-wc)
        self.ct.set_value(total);self.co.set_value(ok);self.ce.set_value(ec);self.cw.set_value(wc)

        self.all_table_rows=[]
        for r in sorted(errs.keys()):
            for et,cols,msg in errs[r]:
                if '必输' in et or '金额' in et: cat='必输项'
                elif '借贷' in et: cat='借贷平衡'
                elif '记账码' in et: cat='记账码'
                elif '禁填' in et: cat='禁填字段'
                elif '费用' in et: cat='费用类别'
                else: cat='其他'
                self.all_table_rows.append((str(r),'',f"❌ {et}",msg,cat,False))
        for w in warns:
            if w[2]=='科目不在规则表':
                self.all_table_rows.append((str(w[0]),w[1],"⚠️ 科目不在规则表","请核实",'警告',True))
            elif '记账码' in w[2]:
                self.all_table_rows.append((str(w[0]),'',f"⚠️ {w[2]}",str(w[1]),'记账码',True))

        self.tabs.setCurrentIndex(0);self._fill(self.all_table_rows)
        self.vbtn.setEnabled(True);self.dbtn.setEnabled(True)
        self.statusBar().showMessage(f"  校验完成  {datetime.now().strftime('%H:%M:%S')}")

    def _tab_changed(self,idx):
        if not self.all_table_rows: return
        fmap={0:None,1:'必输项',2:'借贷平衡',3:'记账码',4:'禁填字段',5:'费用类别',6:'警告'}
        f=fmap.get(idx)
        if f is None: self._fill(self.all_table_rows)
        else:
            iw=(f=='警告')
            self._fill([r for r in self.all_table_rows if r[4]==f or (iw and r[5])])

    def _fill(self,rows):
        self.tbl.setRowCount(len(rows))
        for i,(rno,acct,et,detail,cat,iw) in enumerate(rows):
            bg=QColor("#FFF9C4") if not iw else QColor("#FFE0B2")
            for j,v in enumerate([rno,acct,et,detail]):
                item=QTableWidgetItem(v);item.setBackground(bg)
                if j==2 and not iw: item.setForeground(QColor(C_ERROR))
                elif j==2: item.setForeground(QColor(C_WARN))
                self.tbl.setItem(i,j,item)

    def _err(self,msg):
        self.vbtn.setEnabled(True);self.prog.setValue(0)
        QMessageBox.critical(self,"校验出错",f"错误：\n{msg}")
        self.statusBar().showMessage("  校验失败")

    def _download(self):
        if not self.result: return
        dn=f"校验报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        sp,_=QFileDialog.getSaveFileName(self,"保存",dn,"Excel (*.xlsx)")
        if sp:
            try:
                self.result["out_wb"].save(sp)
                QMessageBox.information(self,"✅ 保存成功",f"报告已保存到：\n{sp}")
                self.statusBar().showMessage(f"  已保存：{os.path.basename(sp)}")
            except Exception as e: QMessageBox.critical(self,"保存失败",str(e))


if __name__=="__main__":
    app=QApplication(sys.argv);app.setStyle("Fusion")
    win=MainWindow();win.show();sys.exit(app.exec_())
