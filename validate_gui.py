#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAP凭证校验工具 v2.3 - PyQt5 GUI（橙色主题 + 表格粘贴修复 + 布局优化）
打包：pyinstaller --onefile --windowed --name SAP凭证校验工具 validate_gui.py
"""
import sys, os, json
from datetime import datetime
from decimal import Decimal, InvalidOperation
from collections import defaultdict

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem,
    QHeaderView, QProgressBar, QMessageBox, QGroupBox, QAbstractItemView,
    QTabBar, QFrame, QGraphicsDropShadowEffect, QStackedWidget, QToolTip,
    QShortcut
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QPoint
from PyQt5.QtGui import QFont, QColor, QDragEnterEvent, QDropEvent, QKeySequence

import openpyxl
from openpyxl.styles import PatternFill, Font as XlFont, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string

# ━━━━━━━ 常量 ━━━━━━━
APP_NAME="SAP凭证校验工具"; VERSION="v2.4"; DATA_ROW=4
CONFIG_FILE=os.path.join(os.path.expanduser("~"),".sap_validate_config.json")
CACHE_FILE=os.path.join(os.path.expanduser("~"),".sap_validate_cache.json")

DEBIT_CODES={'01','21','40'}; CREDIT_CODES={'11','31','50'}; ALL_CODES=DEBIT_CODES|CREDIT_CODES
RECON_CODES={'D':{'01','11'},'K':{'21','31'},None:{'40','50'},'':{'40','50'}}
RECON_LABELS={'D':'客户统驭','K':'供应商统驭',None:'普通科目','':'普通科目'}
HARDCODED_RULES=[(1002010001,1002019999,[('AD','现金流量码'),('X','利润中心')])]
FORBIDDEN_COLS_NON6=[
    ('AF','客户'),('AG','供应商'),('AH','生产（物料）'),('AI','流量类型'),
    ('AJ','渠道'),('AK','一级费用'),('AL','二级费用'),('AM','城市'),
    ('AN','项目'),('AO','销售部门'),('AP','产品'),('AQ','提报人'),
    ('AR','提报部门'),('AS','报销人'),('AT','销售大区'),('AU','销售员')]

A_IDX=0; B_IDX=1; J_IDX=column_index_from_string('J')-1; M_IDX=column_index_from_string('M')-1
O_IDX=column_index_from_string('O')-1; AK_IDX=column_index_from_string('AK')-1; AL_IDX=column_index_from_string('AL')-1

C_PRIMARY='#424242';C_PRIMARY_L='#616161';C_PRIMARY_XL='#F5F5F5'
C_ACCENT='#5C7CFA';C_BG='#FAFAFA';C_CARD='#FFFFFF'
C_TEXT='#212121';C_TEXT_SEC='#757575';C_SUCCESS='#51CF66';C_ERROR='#E53935';C_WARN='#FB8C00';C_BORDER='#E0E0E0'

HEADER_ROW1=['凭证序号','公司代码','凭证类型','凭证日期','过账日期','货币','汇率','凭证抬头文本','参照','记帐代码','屏幕科目','特别总帐标识','总账科目','事物类型','凭证货币金额','本币金额','成本中心编号','订单号','基准日期','参考','分配','项目文本','反记账标识','利润中心','数量','单位','付款条件','冻结付款','功能范围','现金流量码','贸易伙伴','客户','供应商','生产（物料）','流量类型','渠道','一级费用','二级费用','城市','项目','销售部门','产品','提报人','提报部门','报销人','销售大区','销售员','银行流水号','物料','OA单据号','销售合同号']
HEADER_ROW2=['LINEID','BUKRS','BLART','BLDAT','BUDAT','WAERS','KURSF','BKTXT','XBLNR','BSCHL','NEWKO','UMSKZ','HKONT','BEWAR','WRBTR','DMBTR','KOSTL','AUFNR','ZFBDT','XREF1','ZUONR','SGTXT','XNEGP','PRCTR','MENGE','MEINS','ZTERM','ZLSPR','FKBER','ZZCASHFLOW','RASSC','KUNNR','WWVEN','artnr','WWLL','WWOFR','WWFY1','WWFY2','WWCS','WWXM','WWXSB','WWPRC','WWTBR','WWTBM','WWBXR','WWXSQ','WWXSY','XREF3','MATNR','XREF1_HD','XREF2_HD']
HEADER_ROW3=[10,4,2,'8(YYYYMMDD)','8(YYYYMMDD)',5,'(9,5)',25,16,2,10,1,10,3,'(13,2)','(13,2)',10,12,'8(YYYYMMDD)',12,18,50,1,10,'(13,3)',3,4,1,4,5,5,6,6,13,18,18,18,18,18,18,18,18,18,10,18,18,18,20,18,20,20]
NUM_COLS=len(HEADER_ROW1)
COL_LETTERS=[get_column_letter(i+1) for i in range(NUM_COLS)]

# ━━━━━━━ 字体体系 ━━━━━━━
FONT_FAMILY = "微软雅黑"
FONT_TITLE  = 14  # 标题
FONT_H2     = 12  # 二级标题/按钮
FONT_BODY   = 11  # 正文
FONT_SMALL  = 10  # 辅助信息
FONT_GRID   = 10  # 表格（表头+数据统一）

# ━━━━━━━ 配置持久化 ━━━━━━━
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE,'r',encoding='utf-8') as f: return json.load(f)
        except: pass
    return {}
def save_config(cfg):
    try:
        with open(CONFIG_FILE,'w',encoding='utf-8') as f: json.dump(cfg,f,ensure_ascii=False,indent=2)
    except: pass
def load_mapping_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE,'r',encoding='utf-8') as f: return json.load(f)
        except: pass
    return None
def save_mapping_cache(data):
    try:
        with open(CACHE_FILE,'w',encoding='utf-8') as f: json.dump(data,f,ensure_ascii=False,indent=2)
    except: pass

# ━━━━━━━ 校验逻辑（与 v2.1 相同） ━━━━━━━
def col_letter_to_idx(letter): return column_index_from_string(letter)-1
def get_hardcoded_required(acct_str):
    try: acct_int=int(acct_str)
    except: return None
    for s,e,cols in HARDCODED_RULES:
        if s<=acct_int<=e: return cols
    return None
def load_rule_table(rule_file):
    wb=openpyxl.load_workbook(rule_file);ws=wb.active;rows=list(ws.iter_rows(values_only=True))
    fn,ci=rows[0],rows[1]; rm,rcm={},{}
    for row in rows[2:]:
        a=row[0]
        if a is None: continue
        a_s=str(int(a)) if isinstance(a,float) else str(a).strip()
        rc=row[2] if len(row)>2 else None
        if rc: rc=str(rc).strip()
        rcm[a_s]=rc if rc else None
        req=[]
        for f,c,v in zip(fn,ci,row):
            if c and c!='M' and v=='√': req.append((c,f))
        rm[a_s]=req
    return rm,rcm
def load_mapping_table(mf):
    wb=openpyxl.load_workbook(mf);ws=wb.active;rows=list(ws.iter_rows(values_only=True))
    mp=defaultdict(set)
    for row in rows[1:]:
        if row[0] is None: continue
        a=str(int(row[0])) if isinstance(row[0],float) else str(row[0]).strip()
        f1=str(int(row[1])) if isinstance(row[1],float) else str(row[1]).strip() if row[1] else ''
        f2=str(int(row[2])) if isinstance(row[2],float) else str(row[2]).strip() if row[2] else ''
        mp[a].add((f1,f2))
    return {k:list(v) for k,v in mp.items()}
def safe_decimal(val):
    if val is None: return None
    try: return Decimal(str(val).strip())
    except: return None
def check_decimal_places(val):
    d=safe_decimal(val)
    if d is None: return True
    s,dg,exp=d.as_tuple()
    return abs(exp)<=2 if exp<0 else True
def _norm(val):
    if val is None: return ''
    if isinstance(val,float): return str(int(val))
    return str(val).strip()

def run_validate_data(all_rows, rule_map, recon_map, mapping_data=None, progress_cb=None):
    mapping_sets={}
    if mapping_data:
        for a,combos in mapping_data.items(): mapping_sets[a]=set(tuple(c) for c in combos)
    total=len(all_rows)
    if total==0: return {},[]
    errors=defaultdict(list); warnings=[]
    for idx,(excel_row,rv) in enumerate(all_rows.items()):
        if progress_cb and idx%50==0: progress_cb(int(idx/total*55))
        acct_val=rv[M_IDX] if M_IDX<len(rv) else None
        if acct_val is None: continue
        acct_str=_norm(acct_val)
        hc=get_hardcoded_required(acct_str); is_hc=hc is not None
        if is_hc: req_cols=hc
        elif acct_str in rule_map: req_cols=rule_map[acct_str]
        else: warnings.append((excel_row,acct_str,'科目不在规则表')); req_cols=[]
        for cl,fn in req_cols:
            ci=col_letter_to_idx(cl); v=rv[ci] if ci<len(rv) else None
            if v is None or str(v).strip()=='':
                errors[excel_row].append(('必输项为空',[cl],f'缺失必输项：{cl}({fn})'))
        o_val=rv[O_IDX] if O_IDX<len(rv) else None
        if o_val is None or str(o_val).strip()=='':
            errors[excel_row].append(('金额为空',['O'],'凭证货币金额(O列)为空'))
        elif not check_decimal_places(o_val):
            errors[excel_row].append(('金额格式错误',['O'],f'金额小数位超过2位：{o_val}'))
        j_val=rv[J_IDX] if J_IDX<len(rv) else None; j_str=_norm(j_val)
        if len(j_str)==1: j_str='0'+j_str
        if not is_hc and acct_str in recon_map:
            rt=recon_map[acct_str]; rk=''
            if rt:
                ru=str(rt).strip().upper()
                if '客户' in str(rt) or ru=='D': rk='D'
                elif '供应' in str(rt) or ru=='K': rk='K'
            allowed=RECON_CODES.get(rk,set()); label=RECON_LABELS.get(rk,'普通科目')
            if j_str and j_str not in allowed:
                errors[excel_row].append(('记账码不匹配',['J'],f'统驭科目为{label}({rk or "空"})，应为{"/".join(sorted(allowed))}，实际{j_str}'))
        if j_str and j_str not in ALL_CODES:
            warnings.append((excel_row,j_str,'记账码不在标准范围'))
        if not acct_str.startswith('6'):
            for cl,fn in FORBIDDEN_COLS_NON6:
                ci=col_letter_to_idx(cl); v=rv[ci] if ci<len(rv) else None
                if v is not None and str(v).strip()!='':
                    errors[excel_row].append(('禁填字段有值',[cl],f'非6开头科目，{cl}({fn})不允许有值'))
        if mapping_sets and acct_str in mapping_sets:
            ak_val=_norm(rv[AK_IDX] if AK_IDX<len(rv) else None)
            al_val=_norm(rv[AL_IDX] if AL_IDX<len(rv) else None)
            if ak_val and al_val:
                if (ak_val,al_val) not in mapping_sets[acct_str]:
                    errors[excel_row].append(('费用类别不匹配',['AK','AL'],f'科目{acct_str}下，一级费用{ak_val}+二级费用{al_val}组合不在配置表中'))
    if progress_cb: progress_cb(65)
    groups=defaultdict(list)
    for r,rv in all_rows.items():
        a=rv[A_IDX] if A_IDX<len(rv) else None
        if a is None: continue
        groups[(_norm(a),_norm(rv[B_IDX] if B_IDX<len(rv) else None))].append(r)
    if progress_cb: progress_cb(75)
    for (lid,bk),rlist in groups.items():
        ds,cs=Decimal('0'),Decimal('0')
        for r in rlist:
            rv=all_rows[r]; ov=rv[O_IDX] if O_IDX<len(rv) else None
            jv=rv[J_IDX] if J_IDX<len(rv) else None; js=_norm(jv)
            if len(js)==1: js='0'+js
            amt=safe_decimal(ov)
            if amt is None: continue
            if js in DEBIT_CODES: ds+=amt
            elif js in CREDIT_CODES: cs+=amt
        diff=ds-cs
        if diff!=0:
            msg=f'借贷不平衡：凭证{lid}/公司{bk}，借方{ds} 贷方{cs} 差额{diff}'
            for r in rlist: errors[r].append(('借贷不平衡',[],msg))
    if progress_cb: progress_cb(90)
    return errors,warnings

def run_validate(b_file, rule_file, mapping_data=None, progress_cb=None):
    rule_map,recon_map=load_rule_table(rule_file)
    wb_b=openpyxl.load_workbook(b_file);ws_b=wb_b.active;mc=ws_b.max_column
    total=ws_b.max_row-DATA_ROW+1
    if total<=0: return {},[], openpyxl.Workbook(),ws_b,rule_map,0
    all_rows={}
    for r in range(DATA_ROW,ws_b.max_row+1):
        all_rows[r]=[ws_b.cell(row=r,column=c+1).value for c in range(mc)]
    errors,warnings=run_validate_data(all_rows,rule_map,recon_map,mapping_data,progress_cb)
    out_wb=build_report(errors,warnings,wb_b,ws_b,rule_map,total)
    if progress_cb: progress_cb(100)
    return errors,warnings,out_wb,ws_b,rule_map,total

def build_report(errors,warnings,wb_b,ws_b,rule_map,total):
    RED=PatternFill(start_color="FF4444",end_color="FF4444",fill_type="solid")
    YELLOW=PatternFill(start_color="FFF2CC",end_color="FFF2CC",fill_type="solid")
    ORANGE=PatternFill(start_color="FFE0B2",end_color="FFE0B2",fill_type="solid")
    HDR=PatternFill(start_color="E65100",end_color="E65100",fill_type="solid")
    GRAY=PatternFill(start_color="F2F2F2",end_color="F2F2F2",fill_type="solid")
    WHITE=PatternFill(start_color="FFFFFF",end_color="FFFFFF",fill_type="solid")
    thin=Side(style="thin",color="DDDDDD");bdr=Border(left=thin,right=thin,top=thin,bottom=thin)
    out_wb=openpyxl.Workbook();ws_out=out_wb.active;ws_out.title="校验结果"
    ecells=set()
    for r,el in errors.items():
        for _,cols,_ in el:
            for cl in cols: ecells.add((r,cl))
    wa={w[0]:w[1] for w in warnings if w[2]=='科目不在规则表'}
    mc=ws_b.max_column;ec=mc+1
    for r in range(1,ws_b.max_row+1):
        for c in range(1,mc+1):
            v=ws_b.cell(row=r,column=c).value;d=ws_out.cell(row=r,column=c,value=v)
            d.border=bdr;d.alignment=Alignment(vertical="center");cl=get_column_letter(c)
            if r<=3:
                d.fill=HDR if r==1 else GRAY
                d.font=XlFont(name="微软雅黑",bold=(r==1),color="FFFFFF" if r==1 else "333333",size=10)
            elif r in errors:
                d.fill=RED if (r,cl) in ecells else YELLOW
                d.font=XlFont(name="微软雅黑",color="FFFFFF" if (r,cl) in ecells else "000000",size=10)
            elif r in wa: d.fill=ORANGE;d.font=XlFont(name="微软雅黑",size=10)
            else: d.fill=WHITE;d.font=XlFont(name="微软雅黑",size=10)
        if r==1:
            c2=ws_out.cell(row=1,column=ec,value="错误说明");c2.fill=HDR
            c2.font=XlFont(name="微软雅黑",bold=True,color="FFFFFF",size=10);c2.border=bdr
        if r in errors:
            c2=ws_out.cell(row=r,column=ec,value=" | ".join([e[2] for e in errors[r]]))
            c2.fill=YELLOW;c2.font=XlFont(name="微软雅黑",color="CC0000",size=10);c2.border=bdr
        elif r in wa:
            c2=ws_out.cell(row=r,column=ec,value=f"⚠️ 科目{wa[r]}不在规则表中")
            c2.fill=ORANGE;c2.font=XlFont(name="微软雅黑",color="885500",size=10);c2.border=bdr
    for c in range(1,mc+1): ws_out.column_dimensions[get_column_letter(c)].width=14
    ws_out.column_dimensions[get_column_letter(ec)].width=65;ws_out.freeze_panes=f"A{DATA_ROW}"
    ws2=out_wb.create_sheet("错误汇总")
    for i,h in enumerate(["行号","总账科目","错误类型","错误详情"],1):
        c2=ws2.cell(row=1,column=i,value=h);c2.fill=HDR;c2.font=XlFont(name="微软雅黑",bold=True,color="FFFFFF",size=11)
        c2.alignment=Alignment(horizontal="center",vertical="center");c2.border=bdr
    ws2.column_dimensions['A'].width=8;ws2.column_dimensions['B'].width=16
    ws2.column_dimensions['C'].width=18;ws2.column_dimensions['D'].width=65
    ri=2
    for r in sorted(errors.keys()):
        rv=[ws_b.cell(row=r,column=c+1).value for c in range(ws_b.max_column)]
        a_s=_norm(rv[M_IDX] if M_IDX<len(rv) else '')
        for et,cols,msg in errors[r]:
            for i,v in enumerate([r,a_s,f"❌ {et}",msg],1):
                c2=ws2.cell(row=ri,column=i,value=v)
                c2.fill=PatternFill(start_color="FFF2CC",end_color="FFF2CC",fill_type="solid")
                c2.font=XlFont(name="微软雅黑",size=10);c2.alignment=Alignment(vertical="center",wrap_text=(i==4));c2.border=bdr
            ri+=1
    return out_wb

# ━━━━━━━ 后台线程 ━━━━━━━
class ValidateWorker(QThread):
    progress=pyqtSignal(int);finished=pyqtSignal(object);error=pyqtSignal(str)
    def __init__(self,b,r,m=None): super().__init__();self.b=b;self.r=r;self.m=m
    def run(self):
        try:
            e,w,wb,ws,rm,t=run_validate(self.b,self.r,self.m,lambda v:self.progress.emit(v))
            self.finished.emit({"errors":e,"warnings":w,"out_wb":wb,"total":t})
        except: import traceback;self.error.emit(traceback.format_exc())

class GridValidateWorker(QThread):
    progress=pyqtSignal(int);finished=pyqtSignal(object);error=pyqtSignal(str)
    def __init__(self,all_rows,rule_file,mapping_data=None):
        super().__init__();self.all_rows=all_rows;self.rule_file=rule_file;self.mapping_data=mapping_data
    def run(self):
        try:
            rm,rcm=load_rule_table(self.rule_file)
            e,w=run_validate_data(self.all_rows,rm,rcm,self.mapping_data,lambda v:self.progress.emit(v))
            self.finished.emit({"errors":e,"warnings":w,"total":len(self.all_rows)})
        except: import traceback;self.error.emit(traceback.format_exc())

# ━━━━━━━ 支持多行粘贴的 QTableWidget ━━━━━━━
class PasteableTable(QTableWidget):
    """重写 keyPressEvent 支持 Ctrl+V 多行多列粘贴"""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Paste):
            self._paste_from_clipboard()
        else:
            super().keyPressEvent(event)

    def _paste_from_clipboard(self):
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        if not text:
            return
        # 解析 Tab 分隔的多行文本
        lines = text.rstrip('\n').split('\n')
        rows = [line.split('\t') for line in lines]

        # 获取当前选中单元格作为起始位置
        sel = self.selectedItems()
        if sel:
            start_row = sel[0].row()
            start_col = sel[0].column()
        else:
            start_row = self.currentRow()
            start_col = self.currentColumn()

        # 如果起始行在表头区（前3行），从第4行开始
        if start_row < 3:
            start_row = 3

        # 扩展表格行数
        needed = start_row + len(rows)
        if needed > self.rowCount():
            self.setRowCount(needed + 50)

        # 填充数据
        for i, row in enumerate(rows):
            for j, val in enumerate(row):
                r = start_row + i
                c = start_col + j
                if c >= self.columnCount():
                    break
                # 跳过表头行
                if r < 3:
                    continue
                item = self.item(r, c)
                if item is None:
                    item = QTableWidgetItem()
                    self.setItem(r, c, item)
                if item.flags() & Qt.ItemIsEditable:
                    item.setText(val.strip())
                elif not (item.flags() & Qt.ItemIsEnabled):
                    # 不可编辑的表头行，跳过
                    pass
                else:
                    item.setText(val.strip())

        self.viewport().update()

# ━━━━━━━ 自定义控件 ━━━━━━━
class StatBar(QFrame):
    """一行式统计栏"""
    def __init__(self):
        super().__init__()
        self.setFixedHeight(36)
        self.setStyleSheet(f"""
            StatBar {{
                background: {C_CARD}; border-radius: 6px;
                border: 1px solid {C_BORDER}; padding: 0 12px;
            }}
        """)
        ly = QHBoxLayout(self); ly.setContentsMargins(16, 0, 16, 0); ly.setSpacing(24)
        self.labels = {}
        for icon, key, color in [("📊","total",C_TEXT),("✅","ok","#2E7D32"),("❌","err","#E53935"),("⚠️","warn","#FB8C00")]:
            name = {"total":"总行数","ok":"通过","err":"错误","warn":"警告"}[key]
            lbl = QLabel(f"{icon} {name}: —")
            lbl.setStyleSheet(f"color: {color}; font-size: {FONT_BODY}px; font-weight: bold; border: none;")
            ly.addWidget(lbl)
            self.labels[key] = lbl
        ly.addStretch()
    def set_values(self, total="—", ok="—", err="—", warn="—"):
        self.labels["total"].setText(f"📊 总行数: {total}")
        self.labels["ok"].setText(f"✅ 通过: {ok}")
        self.labels["err"].setText(f"❌ 错误: {err}")
        self.labels["warn"].setText(f"⚠️ 警告: {warn}")

class DropArea(QLabel):
    file_dropped = pyqtSignal(str)
    def __init__(self, text="拖拽文件到此处"):
        super().__init__(text)
        self.setAcceptDrops(True); self.setAlignment(Qt.AlignCenter)
        self.setFixedHeight(52); self._idle()
    def _idle(self):
        self.setStyleSheet(f"""
            QLabel {{
                border: 2px dashed {C_ACCENT}; border-radius: 8px;
                color: {C_TEXT_SEC}; background: {C_PRIMARY_XL};
                font-size: {FONT_BODY}px; padding: 4px;
            }}
        """)
    def _loaded(self, fn):
        self.setText(f"📄 {fn}")
        self.setStyleSheet(f"""
            QLabel {{
                border: 2px solid {C_SUCCESS}; border-radius: 8px;
                color: {C_SUCCESS}; background: #E8F5E9;
                font-size: {FONT_BODY}px; font-weight: bold; padding: 4px;
            }}
        """)
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls(): e.acceptProposedAction()
    def dragLeaveEvent(self, e): self._idle()
    def dropEvent(self, e):
        urls = e.mimeData().urls()
        if urls:
            p = urls[0].toLocalFile()
            if p.lower().endswith(('.xlsx','.xls')):
                self._loaded(os.path.basename(p)); self.file_dropped.emit(p)
            else:
                QMessageBox.warning(self, "格式错误", "请拖入Excel文件"); self._idle()

# ━━━━━━━ 主窗口 ━━━━━━━
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.b_file = ""; self.rule_file = ""; self.mapping_file = ""
        self.mapping_data = None; self.result = None; self.all_table_rows = []
        self._init_ui(); self._load_saved()

    def _load_saved(self):
        cfg = load_config()
        rp = cfg.get("rule_file", "")
        if rp and os.path.exists(rp):
            self.rule_file = rp; self.r_drop._loaded(os.path.basename(rp))
        mp = cfg.get("mapping_file", "")
        if mp and os.path.exists(mp):
            self.mapping_file = mp; self.m_drop._loaded(os.path.basename(mp))
        cached = load_mapping_cache()
        if cached:
            self.mapping_data = cached
            cnt = sum(len(v) for v in cached.values())
            self.m_info.setText(f"✅ 已加载 {len(cached)} 个科目 {cnt} 条规则")
            self.m_info.setStyleSheet(f"color: {C_SUCCESS}; font-size: {FONT_SMALL}px;")
        bp = cfg.get("b_file", "")
        if bp and os.path.exists(bp):
            self.b_file = bp; self.b_drop._loaded(os.path.basename(bp))

    def _save_paths(self):
        cfg = load_config()
        if self.rule_file: cfg["rule_file"] = self.rule_file
        if self.b_file: cfg["b_file"] = self.b_file
        if self.mapping_file: cfg["mapping_file"] = self.mapping_file
        save_config(cfg)

    def _bs(self, color, sz=FONT_H2, px=20):
        return f"""
            QPushButton {{
                background: {color}; color: white; border-radius: 6px;
                padding: 8px {px}px; font-size: {sz}px; font-weight: bold; border: none;
            }}
            QPushButton:hover {{ background: {color}dd; }}
            QPushButton:pressed {{ background: {color}bb; }}
            QPushButton:disabled {{ background: #ccc; color: #888; }}
        """

    def _init_ui(self):
        self.setWindowTitle(f"{APP_NAME}  {VERSION}")
        self.setMinimumSize(1100, 780)
        self.resize(1280, 900)
        self.setStyleSheet(f"""
            QMainWindow {{ background: {C_BG}; }}
            QGroupBox {{
                font-weight: bold; font-size: {FONT_H2}px; color: {C_TEXT};
                border: 1px solid {C_BORDER}; border-radius: 8px;
                margin-top: 14px; padding-top: 22px; background: {C_CARD};
            }}
            QGroupBox::title {{ subcontrol-origin: margin; left: 18px; padding: 0 8px; }}
            QStatusBar {{ background: #F0F0F0; color: {C_TEXT_SEC}; font-size: {FONT_SMALL}px; }}
        """)
        QApplication.setFont(QFont(FONT_FAMILY, FONT_BODY))

        central = QWidget(); self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(12); root.setContentsMargins(20, 14, 20, 14)

        # ── 标题栏 ──
        hdr = QFrame(); hdr.setFixedHeight(40)
        hdr.setStyleSheet(f"""
            QFrame {{
                background: {C_CARD}; border-radius: 8px;
                border: 1px solid {C_BORDER};
            }}
        """)
        hl = QHBoxLayout(hdr); hl.setContentsMargins(20, 0, 20, 0)
        ht = QLabel(f"📋  {APP_NAME}")
        ht.setStyleSheet(f"color: {C_TEXT}; font-size: {FONT_TITLE}px; font-weight: bold; background: transparent;")
        hv = QLabel(VERSION)
        hv.setStyleSheet(f"color: {C_TEXT_SEC}; font-size: {FONT_SMALL}px; background: transparent;")
        hl.addWidget(ht); hl.addStretch(); hl.addWidget(hv)
        root.addWidget(hdr)

        # ── 页面切换 Tab ──
        self.page_tab = QTabBar()
        self.page_tab.setStyleSheet(f"""
            QTabBar::tab {{
                background: {C_CARD}; border: 1px solid {C_BORDER}; border-bottom: none;
                border-radius: 8px 8px 0 0; padding: 10px 28px;
                margin-right: 4px; font-size: {FONT_H2}px; font-weight: bold; color: {C_TEXT_SEC};
            }}
            QTabBar::tab:selected {{ background: {C_ACCENT}; color: white; }}
            QTabBar::tab:hover {{ background: #E8E8E8; }}
        """)
        self.page_tab.addTab("📄  文件校验")
        self.page_tab.addTab("📋  表格校验")
        self.page_tab.currentChanged.connect(lambda idx: self.stack.setCurrentIndex(idx))
        root.addWidget(self.page_tab)

        # ── 堆叠页面 ──
        self.stack = QStackedWidget()
        self.stack.addWidget(self._build_file_page())
        self.stack.addWidget(self._build_grid_page())
        root.addWidget(self.stack, 1)

        self.statusBar().showMessage(f"  {APP_NAME} {VERSION}  |  就绪")

    # ════════════════ 页面一：文件校验 ════════════════
    def _build_file_page(self):
        page = QWidget()
        root = QVBoxLayout(page); root.setSpacing(12); root.setContentsMargins(0, 8, 0, 0)

        # 上传区
        ub = QGroupBox("  文件上传（支持拖拽）")
        ul = QVBoxLayout(ub); ul.setSpacing(10)
        for lbl_text, attr, slot, drop_text in [
            ("📄 B表：", "b_drop", "_choose_b", "将B表拖到此处"),
            ("📋 校验规则表：", "r_drop", "_choose_rule", "将规则表拖到此处（记住路径）"),
            ("📊 费用配置表：", "m_drop", "_choose_mapping", "将费用mapping表拖到此处（记住并缓存）")
        ]:
            row = QHBoxLayout(); row.setSpacing(10)
            l = QLabel(lbl_text); l.setFixedWidth(100)
            l.setStyleSheet(f"font-size: {FONT_BODY}px;")
            da = DropArea(drop_text); setattr(self, attr, da)
            b = QPushButton("选择"); b.setFixedWidth(70); b.setFixedHeight(36)
            b.setStyleSheet(self._bs(C_PRIMARY, FONT_BODY, 12))
            b.clicked.connect(getattr(self, slot))
            if attr == 'b_drop': da.file_dropped.connect(self._on_b_drop)
            elif attr == 'r_drop': da.file_dropped.connect(self._on_r_drop)
            else: da.file_dropped.connect(self._on_m_drop)
            row.addWidget(l); row.addWidget(da, 1); row.addWidget(b)
            ul.addLayout(row)
        self.m_info = QLabel("")
        self.m_info.setStyleSheet(f"color: {C_TEXT_SEC}; font-size: {FONT_SMALL}px;")
        ul.addWidget(self.m_info)
        root.addWidget(ub)

        # 按钮行
        brow = QHBoxLayout(); brow.setSpacing(12)
        self.vbtn = QPushButton("  🔍  开始校验")
        self.vbtn.setFixedHeight(42); self.vbtn.setStyleSheet(self._bs(C_SUCCESS))
        self.vbtn.clicked.connect(self._start)
        self.dbtn = QPushButton("  ⬇️  下载报告")
        self.dbtn.setFixedHeight(42); self.dbtn.setStyleSheet(self._bs(C_PRIMARY))
        self.dbtn.setEnabled(False); self.dbtn.clicked.connect(self._download)
        brow.addWidget(self.vbtn); brow.addWidget(self.dbtn); brow.addStretch()
        root.addLayout(brow)

        # 进度条
        self.prog = QProgressBar(); self.prog.setFixedHeight(6)
        self.prog.setValue(0); self.prog.setTextVisible(False)
        self.prog.setStyleSheet(f"""
            QProgressBar {{ background: #e0e0e0; border-radius: 3px; }}
            QProgressBar::chunk {{ background: {C_ACCENT}; border-radius: 3px; }}
        """)
        root.addWidget(self.prog)

        # 统计栏
        self.stat_bar = StatBar()
        root.addWidget(self.stat_bar)

        # 结果区
        rb2 = QGroupBox("  校验结果"); rl2 = QVBoxLayout(rb2); rl2.setSpacing(8)
        self.tabs = QTabBar()
        self.tabs.setStyleSheet(f"""
            QTabBar::tab {{
                background: {C_CARD}; border: 1px solid {C_BORDER}; border-bottom: none;
                border-radius: 6px 6px 0 0; padding: 6px 16px;
                margin-right: 2px; font-size: {FONT_BODY}px; color: {C_TEXT_SEC};
            }}
            QTabBar::tab:selected {{ background: {C_ACCENT}; color: white; font-weight: bold; }}
            QTabBar::tab:hover {{ background: #E8E8E8; }}
        """)
        for t in ["全部", "必输项", "借贷平衡", "记账码", "禁填字段", "费用类别", "⚠️ 警告"]:
            self.tabs.addTab(t)
        self.tabs.currentChanged.connect(self._tab_changed)
        rl2.addWidget(self.tabs)

        self.tbl = QTableWidget(); self.tbl.setColumnCount(4)
        self.tbl.setHorizontalHeaderLabels(["行号", "总账科目", "错误类型", "错误详情"])
        self.tbl.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.tbl.setColumnWidth(0, 70); self.tbl.setColumnWidth(1, 130); self.tbl.setColumnWidth(2, 160)
        self.tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl.setAlternatingRowColors(True); self.tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl.verticalHeader().setDefaultSectionSize(28)
        self.tbl.setStyleSheet(f"""
            QTableWidget {{ border: 1px solid {C_BORDER}; gridline-color: #E8E8E8; font-size: {FONT_BODY}px; }}
            QHeaderView::section {{
                background: #F5F5F5; color: {C_TEXT}; font-weight: bold;
                padding: 8px; border: none; border-bottom: 1px solid #D0D0D0;
                font-size: {FONT_BODY}px;
            }}
        """)
        rl2.addWidget(self.tbl)
        root.addWidget(rb2, 1)
        return page

    # ════════════════ 页面二：表格校验 ════════════════
    def _build_grid_page(self):
        page = QWidget()
        root = QVBoxLayout(page); root.setSpacing(10); root.setContentsMargins(0, 8, 0, 0)

        # 提示信息
        tip = QLabel("💡 从 Excel / SAP 中复制数据后，点击第4行任意单元格，按 Ctrl+V 粘贴（支持多行多列）")
        tip.setStyleSheet(f"""
            QLabel {{
                color: {C_ACCENT}; font-size: {FONT_BODY}px;
                background: #EEF1FF; padding: 8px 16px;
                border-radius: 6px; border: 1px solid #D0D8FF;
            }}
        """)
        root.addWidget(tip)

        # 按钮栏
        brow = QHBoxLayout(); brow.setSpacing(10)
        self.g_vbtn = QPushButton("  🔍  校验")
        self.g_vbtn.setFixedHeight(40); self.g_vbtn.setStyleSheet(self._bs(C_SUCCESS))
        self.g_vbtn.clicked.connect(self._grid_validate)
        self.g_clr = QPushButton("  🗑️  清空数据")
        self.g_clr.setFixedHeight(40); self.g_clr.setStyleSheet(self._bs("#757575"))
        self.g_clr.clicked.connect(self._grid_clear)
        self.g_dl = QPushButton("  ⬇️  下载报告")
        self.g_dl.setFixedHeight(40); self.g_dl.setStyleSheet(self._bs(C_PRIMARY))
        self.g_dl.setEnabled(False); self.g_dl.clicked.connect(self._grid_download)
        brow.addWidget(self.g_vbtn); brow.addWidget(self.g_clr); brow.addWidget(self.g_dl); brow.addStretch()
        root.addLayout(brow)

        # 进度 + 统计
        self.g_prog = QProgressBar(); self.g_prog.setFixedHeight(6)
        self.g_prog.setValue(0); self.g_prog.setTextVisible(False)
        self.g_prog.setStyleSheet(f"""
            QProgressBar {{ background: #e0e0e0; border-radius: 3px; }}
            QProgressBar::chunk {{ background: {C_ACCENT}; border-radius: 3px; }}
        """)
        root.addWidget(self.g_prog)

        sr = QHBoxLayout(); sr.setSpacing(12)
        self.g_stat = StatBar()
        root.addWidget(self.g_stat)

        # 大表格（使用支持粘贴的自定义控件）
        nc = NUM_COLS + 1
        self.grid = PasteableTable()
        self.grid.setColumnCount(nc)
        self.grid.setHorizontalHeaderLabels(COL_LETTERS + ["错误说明"])
        self.grid.setRowCount(3 + 200)

        # 填充表头
        for r, hrow in enumerate([HEADER_ROW1, HEADER_ROW2, HEADER_ROW3]):
            for c, v in enumerate(hrow):
                item = QTableWidgetItem(str(v) if v is not None else "")
                item.setBackground(QColor("#FFFFFF"))
                item.setForeground(QColor("#5D4037") if r == 0 else QColor("#757575"))
                item.setFont(QFont(FONT_FAMILY, FONT_GRID, QFont.Bold if r == 0 else QFont.Normal))
                item.setFlags(Qt.ItemIsEnabled)
                self.grid.setItem(r, c, item)
            # 错误说明列
            item = QTableWidgetItem(["错误说明", "VALIDATE_MSG", ""][r])
            item.setBackground(QColor("#FFFFFF"))
            item.setForeground(QColor("#5D4037") if r == 0 else QColor("#757575"))
            item.setFlags(Qt.ItemIsEnabled)
            self.grid.setItem(r, NUM_COLS, item)

        self.grid.setStyleSheet(f"""
            QTableWidget {{
                border: 1px solid {C_BORDER}; gridline-color: #E8E8E8;
                font-size: {FONT_GRID}px;
            }}
            QHeaderView::section {{
                background: #F5F5F5; color: {C_TEXT}; font-weight: bold;
                padding: 4px; border: none; border-bottom: 1px solid #D0D0D0;
                font-size: {FONT_SMALL}px;
            }}
        """)
        self.grid.horizontalHeader().setDefaultSectionSize(88)
        self.grid.setColumnWidth(NUM_COLS, 280)
        self.grid.verticalHeader().setDefaultSectionSize(24)
        self.grid.setMouseTracking(True)
        self.grid.cellEntered.connect(self._grid_cell_hover)
        root.addWidget(self.grid, 1)

        self.grid_errors = {}
        self.grid_result = None
        return page

    def _grid_cell_hover(self, row, col):
        if col < NUM_COLS:
            cl = COL_LETTERS[col]
            key = (row, cl)
            if key in self.grid_errors:
                pos = self.grid.viewport().mapToGlobal(QPoint(0, 0))
                QToolTip.showText(pos, self.grid_errors[key])

    def _grid_clear(self):
        for r in range(3, self.grid.rowCount()):
            for c in range(self.grid.columnCount()):
                item = self.grid.item(r, c)
                if item:
                    item.setText(""); item.setBackground(QColor("white"))
                    item.setForeground(QColor("black")); item.setToolTip("")
                else:
                    self.grid.setItem(r, c, QTableWidgetItem(""))
        self.grid_errors = {}; self.grid_result = None; self.g_dl.setEnabled(False)
        self.g_stat.set_values()

    def _grid_validate(self):
        if not self.rule_file:
            QMessageBox.warning(self, "提示", "请先在「文件校验」页面上传校验规则表！"); return
        all_rows = {}
        for r in range(3, self.grid.rowCount()):
            row_data = []; has_data = False
            for c in range(NUM_COLS):
                item = self.grid.item(r, c)
                v = item.text().strip() if item else ""
                if v: has_data = True
                try: v = int(v)
                except:
                    try: v = float(v)
                    except: pass
                row_data.append(v if v != "" else None)
            if has_data: all_rows[r + 1] = row_data
        if not all_rows:
            QMessageBox.warning(self, "提示", "表格中没有数据，请先粘贴数据！"); return
        self.g_vbtn.setEnabled(False); self.g_prog.setValue(0)
        self.g_worker = GridValidateWorker(all_rows, self.rule_file, self.mapping_data)
        self.g_worker.progress.connect(self.g_prog.setValue)
        self.g_worker.finished.connect(self._grid_done)
        self.g_worker.error.connect(self._grid_err)
        self.g_worker.start()

    def _grid_done(self, result):
        errors = result["errors"]; warnings = result["warnings"]; total = result["total"]
        ec = len(errors); wc = len([w for w in warnings if w[2] == '科目不在规则表'])
        ok = max(0, total - ec - wc)
        self.g_stat.set_values(total, ok, ec, wc)

        # 清除之前标色
        self.grid_errors = {}
        for r in range(3, self.grid.rowCount()):
            for c in range(self.grid.columnCount()):
                item = self.grid.item(r, c)
                if item:
                    item.setBackground(QColor("white"))
                    item.setForeground(QColor("black")); item.setToolTip("")

        ecells = set()
        for r, el in errors.items():
            for _, cols, _ in el:
                for cl in cols: ecells.add((r, cl))
        wa = {w[0]: w[1] for w in warnings if w[2] == '科目不在规则表'}

        for r, el in errors.items():
            gr = r - 1
            if gr >= self.grid.rowCount(): continue
            for c in range(NUM_COLS):
                item = self.grid.item(gr, c)
                if not item: continue
                cl = COL_LETTERS[c]
                if (r, cl) in ecells:
                    item.setBackground(QColor("#FF4444"))
                    item.setForeground(QColor("white"))
                    msgs = [e[2] for e in el if cl in e[1]]
                    if msgs:
                        tip = "\n".join(msgs); item.setToolTip(tip)
                        self.grid_errors[(gr, cl)] = tip
                else:
                    item.setBackground(QColor("#FFF2CC"))
            ei = self.grid.item(gr, NUM_COLS)
            if not ei: ei = QTableWidgetItem(); self.grid.setItem(gr, NUM_COLS, ei)
            ei.setText(" | ".join([e[2] for e in el]))
            ei.setForeground(QColor(C_ERROR)); ei.setBackground(QColor("#FFF2CC"))

        for r, acct in wa.items():
            gr = r - 1
            if gr >= self.grid.rowCount(): continue
            for c in range(NUM_COLS):
                item = self.grid.item(gr, c)
                if item: item.setBackground(QColor("#FFE0B2"))
            ei = self.grid.item(gr, NUM_COLS)
            if not ei: ei = QTableWidgetItem(); self.grid.setItem(gr, NUM_COLS, ei)
            ei.setText(f"⚠️ 科目{acct}不在规则表中")
            ei.setForeground(QColor(C_WARN)); ei.setBackground(QColor("#FFE0B2"))

        self.grid_result = result; self.g_vbtn.setEnabled(True); self.g_dl.setEnabled(True)
        self.g_prog.setValue(100)
        self.statusBar().showMessage(f"  表格校验完成：{ec}个错误 {wc}个警告  {datetime.now().strftime('%H:%M:%S')}")

    def _grid_err(self, msg):
        self.g_vbtn.setEnabled(True); self.g_prog.setValue(0)
        QMessageBox.critical(self, "校验出错", f"错误：\n{msg}")

    def _grid_download(self):
        if not self.grid_result: return
        wb = openpyxl.Workbook(); ws = wb.active
        for r, hrow in enumerate([HEADER_ROW1, HEADER_ROW2, HEADER_ROW3], 1):
            for c, v in enumerate(hrow, 1): ws.cell(row=r, column=c, value=v)
        for r in range(3, self.grid.rowCount()):
            has = False
            for c in range(NUM_COLS):
                item = self.grid.item(r, c)
                v = item.text().strip() if item else ""
                if v: has = True
                try: v = int(v)
                except:
                    try: v = float(v)
                    except: pass
                ws.cell(row=r+1, column=c+1, value=v if v != "" else None)
            if not has: break
        rm, _ = load_rule_table(self.rule_file)
        out_wb = build_report(self.grid_result["errors"], self.grid_result["warnings"], wb, ws, rm, self.grid_result["total"])
        dn = f"校验报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        sp, _ = QFileDialog.getSaveFileName(self, "保存", dn, "Excel (*.xlsx)")
        if sp:
            try: out_wb.save(sp); QMessageBox.information(self, "✅", f"报告已保存到：\n{sp}")
            except Exception as e: QMessageBox.critical(self, "失败", str(e))

    # ════════════════ 文件校验公共方法 ════════════════
    def _on_b_drop(self, p): self.b_file = p; self._save_paths()
    def _on_r_drop(self, p): self.rule_file = p; self._save_paths()
    def _on_m_drop(self, p): self._load_mapping(p)
    def _choose_b(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择B表", "", "Excel (*.xlsx *.xls)")
        if p: self.b_file = p; self.b_drop._loaded(os.path.basename(p)); self._save_paths()
    def _choose_rule(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择校验规则表", "", "Excel (*.xlsx *.xls)")
        if p: self.rule_file = p; self.r_drop._loaded(os.path.basename(p)); self._save_paths()
    def _choose_mapping(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择费用mapping规则表", "", "Excel (*.xlsx *.xls)")
        if p: self._load_mapping(p)
    def _load_mapping(self, path):
        try:
            data = load_mapping_table(path); self.mapping_data = data; self.mapping_file = path
            save_mapping_cache(data); self._save_paths()
            self.m_drop._loaded(os.path.basename(path))
            cnt = sum(len(v) for v in data.values())
            self.m_info.setText(f"✅ 已加载并缓存：{len(data)} 个科目 {cnt} 条规则")
            self.m_info.setStyleSheet(f"color: {C_SUCCESS}; font-size: {FONT_SMALL}px;")
        except Exception as ex:
            QMessageBox.critical(self, "加载失败", str(ex))

    def _start(self):
        if not self.b_file: QMessageBox.warning(self, "提示", "请先上传B表！"); return
        if not self.rule_file: QMessageBox.warning(self, "提示", "请先上传校验规则表！"); return
        self.vbtn.setEnabled(False); self.dbtn.setEnabled(False)
        self.prog.setValue(0); self.tbl.setRowCount(0)
        self.stat_bar.set_values("...", "...", "...", "...")
        self.worker = ValidateWorker(self.b_file, self.rule_file, self.mapping_data)
        self.worker.progress.connect(self.prog.setValue)
        self.worker.finished.connect(self._done)
        self.worker.error.connect(self._err)
        self.worker.start()

    def _done(self, result):
        self.result = result
        errs = result["errors"]; warns = result["warnings"]; total = result["total"]
        ec = len(errs); wc = len([w for w in warns if w[2] == '科目不在规则表'])
        ok = max(0, total - ec - wc)
        self.stat_bar.set_values(total, ok, ec, wc)
        self.all_table_rows = []
        for r in sorted(errs.keys()):
            for et, cols, msg in errs[r]:
                if '必输' in et or '金额' in et: cat = '必输项'
                elif '借贷' in et: cat = '借贷平衡'
                elif '记账码' in et: cat = '记账码'
                elif '禁填' in et: cat = '禁填字段'
                elif '费用' in et: cat = '费用类别'
                else: cat = '其他'
                self.all_table_rows.append((str(r), '', f"❌ {et}", msg, cat, False))
        for w in warns:
            if w[2] == '科目不在规则表':
                self.all_table_rows.append((str(w[0]), w[1], "⚠️ 科目不在规则表", "请核实", '警告', True))
            elif '记账码' in w[2]:
                self.all_table_rows.append((str(w[0]), '', f"⚠️ {w[2]}", str(w[1]), '记账码', True))
        self.tabs.setCurrentIndex(0); self._fill(self.all_table_rows)
        self.vbtn.setEnabled(True); self.dbtn.setEnabled(True)
        self.statusBar().showMessage(f"  校验完成  {datetime.now().strftime('%H:%M:%S')}")

    def _tab_changed(self, idx):
        if not self.all_table_rows: return
        fmap = {0: None, 1: '必输项', 2: '借贷平衡', 3: '记账码', 4: '禁填字段', 5: '费用类别', 6: '警告'}
        f = fmap.get(idx)
        if f is None: self._fill(self.all_table_rows)
        else: self._fill([r for r in self.all_table_rows if r[4] == f or (f == '警告' and r[5])])

    def _fill(self, rows):
        self.tbl.setRowCount(len(rows))
        for i, (rno, acct, et, detail, cat, iw) in enumerate(rows):
            bg = QColor("#FFF9C4") if not iw else QColor("#FFE0B2")
            for j, v in enumerate([rno, acct, et, detail]):
                item = QTableWidgetItem(v); item.setBackground(bg)
                if j == 2 and not iw: item.setForeground(QColor(C_ERROR))
                elif j == 2: item.setForeground(QColor(C_WARN))
                self.tbl.setItem(i, j, item)

    def _err(self, msg):
        self.vbtn.setEnabled(True); self.prog.setValue(0)
        QMessageBox.critical(self, "校验出错", f"错误：\n{msg}")

    def _download(self):
        if not self.result: return
        dn = f"校验报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        sp, _ = QFileDialog.getSaveFileName(self, "保存", dn, "Excel (*.xlsx)")
        if sp:
            try: self.result["out_wb"].save(sp); QMessageBox.information(self, "✅", f"报告已保存到：\n{sp}")
            except Exception as e: QMessageBox.critical(self, "失败", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv); app.setStyle("Fusion")
    win = MainWindow(); win.show(); sys.exit(app.exec_())
