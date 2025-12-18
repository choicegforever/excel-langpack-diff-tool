import sys
import os
import traceback
import threading
import time
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed

import wx
import pandas as pd
import numpy as np

# ----------------------------- 工具与数据结构 ----------------------------- #

@dataclass
class SheetConfig:
    sheet_name: str
    id_col: Optional[str] = None
    comp_a_col: Optional[str] = None
    comp_b_col: Optional[str] = None

@dataclass
class CompareOptions:
    ignore_case: bool = True
    strip_spaces: bool = True

@dataclass
class GlobalConfig:
    """全局统一配置"""
    id_col: Optional[str] = None
    comp_a_col: Optional[str] = None
    comp_b_col: Optional[str] = None

# 结果单行
@dataclass
class CompareRow:
    sheet: str
    row_index_a: Optional[int]
    id_value: str
    a_value: Optional[str]
    b_value: Optional[str]
    status: str  # 一致/不一致/表B无此ID/Sheet在B不存在

# ----------------------------- 辅助函数 ----------------------------- #

class ExcelCache:
    """Excel文件缓存管理器"""
    def __init__(self):
        self.column_cache = {}  # {(path, sheet): columns}
        self.data_cache = {}    # {(path, sheet, cols_tuple): df}
        self.file_cache = {}    # {path: ExcelFile} - 文件级缓存
        
    def get_excel_file(self, path: str):
        """获取Excel文件对象（带缓存）"""
        if path not in self.file_cache:
            self.file_cache[path] = pd.ExcelFile(path)
        return self.file_cache[path]
        
    def get_columns(self, path: str, sheet_name: str) -> List[str]:
        """获取列名（带缓存）"""
        cache_key = (path, sheet_name)
        if cache_key not in self.column_cache:
            # 使用缓存的文件对象
            excel_file = self.get_excel_file(path)
            df = pd.read_excel(excel_file, sheet_name=sheet_name, nrows=0, dtype=str)
            self.column_cache[cache_key] = list(df.columns.astype(str))
        return self.column_cache[cache_key]
    
    def get_data(self, path: str, sheet_name: str, cols: List[str], filter_empty_ids: bool = True) -> pd.DataFrame:
        """获取数据（带缓存）"""
        cols_tuple = tuple(sorted(cols)) if cols else ()
        cache_key = (path, sheet_name, cols_tuple, filter_empty_ids)
        
        if cache_key not in self.data_cache:
            excel_file = self.get_excel_file(path)
            usecols = list(dict.fromkeys([c for c in cols if c is not None]))
            
            if not usecols:
                self.data_cache[cache_key] = pd.DataFrame()
            else:
                df = pd.read_excel(excel_file, sheet_name=sheet_name, usecols=usecols, dtype=str)
                df = df.fillna("").astype(str)
                
                # 数据预过滤
                if filter_empty_ids and len(cols) > 0 and cols[0] is not None:
                    id_col = cols[0]
                    if id_col in df.columns:
                        df = df[df[id_col].str.strip() != ""].copy()
                
                self.data_cache[cache_key] = df
        
        return self.data_cache[cache_key].copy()  # 返回副本避免修改缓存
    
    def clear_cache(self):
        """清理缓存"""
        # 关闭所有Excel文件
        for excel_file in self.file_cache.values():
            try:
                excel_file.close()
            except:
                pass
        
        self.column_cache.clear()
        self.data_cache.clear()
        self.file_cache.clear()

# 全局缓存实例
excel_cache = ExcelCache()

def safe_read_excel_columns(path: str, sheet_name: str) -> List[str]:
    """只读取列名（使用缓存）。"""
    return excel_cache.get_columns(path, sheet_name)


def load_needed_columns_optimized(path: str, sheet_name: str, cols: List[str], filter_empty_ids: bool = True) -> pd.DataFrame:
    """优化的数据加载函数，使用缓存机制"""
    return excel_cache.get_data(path, sheet_name, cols, filter_empty_ids)


def load_large_file_in_chunks(path: str, sheet_name: str, usecols: List[str], filter_empty_ids: bool = True) -> pd.DataFrame:
    """分块加载大文件"""
    chunks = []
    chunk_size = 10000  # 每次读取1万行
    
    try:
        # 使用openpyxl引擎支持分块读取
        for chunk in pd.read_excel(path, sheet_name=sheet_name, usecols=usecols, 
                                 dtype=str, engine='openpyxl', chunksize=chunk_size):
            chunk = chunk.fillna("").astype(str)
            
            # 分块时也进行预过滤
            if filter_empty_ids and len(usecols) > 0:
                id_col = usecols[0]
                if id_col in chunk.columns:
                    chunk = chunk[chunk[id_col].str.strip() != ""]
            
            if not chunk.empty:
                chunks.append(chunk)
        
        if chunks:
            return pd.concat(chunks, ignore_index=True)
        else:
            return pd.DataFrame()
    except:
        # 如果分块读取也失败，回退到原方法
        df = pd.read_excel(path, sheet_name=sheet_name, usecols=usecols, dtype=str, engine=None)
        df = df.fillna("").astype(str)
        
        # 回退时也进行预过滤
        if filter_empty_ids and len(usecols) > 0:
            id_col = usecols[0]
            if id_col in df.columns:
                df = df[df[id_col].str.strip() != ""]
        
        return df


def normalize_value(s: Optional[str], opt: CompareOptions) -> str:
    if s is None:
        return ""
    val = str(s)
    if opt.strip_spaces:
        val = val.strip()
    if opt.ignore_case:
        val = val.casefold()
    return val


# ----------------------------- 全局配置对话框 ----------------------------- #

class GlobalConfigDialog(wx.Dialog):
    def __init__(self, parent, title: str, cols_a: List[str], cols_b: List[str], config: GlobalConfig):
        super().__init__(parent, title=title, size=(580, 350))
        self.cols_a = cols_a
        self.cols_b = cols_b
        self.config = config

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        tip = wx.StaticText(panel, label="设置所有表格的统一列配置")
        font = tip.GetFont()
        font.MakeBold()
        tip.SetFont(font)
        vbox.Add(tip, 0, wx.ALL, 10)

        grid = wx.FlexGridSizer(3, 3, 10, 10)
        grid.AddGrowableCol(1, 1)
        grid.AddGrowableCol(2, 1)

        # 行1：ID 列
        grid.Add(wx.StaticText(panel, label="ID 列名："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.cb_id = wx.ComboBox(panel, choices=cols_a, style=wx.CB_DROPDOWN, value=config.id_col or "")
        self.cb_id_ref = wx.ComboBox(panel, choices=cols_b, style=wx.CB_READONLY)
        grid.Add(self.cb_id, 1, wx.EXPAND)
        grid.Add(self.cb_id_ref, 1, wx.EXPAND)

        # 行2：表A对比列
        grid.Add(wx.StaticText(panel, label="表A 对比列名："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.cb_comp_a = wx.ComboBox(panel, choices=cols_a, style=wx.CB_DROPDOWN, value=config.comp_a_col or "")
        grid.AddSpacer(1)
        grid.Add(self.cb_comp_a, 1, wx.EXPAND)

        # 行3：表B对比列
        grid.Add(wx.StaticText(panel, label="表B 对比列名："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.cb_comp_b = wx.ComboBox(panel, choices=cols_b, style=wx.CB_DROPDOWN, value=config.comp_b_col or "")
        grid.AddSpacer(1)
        grid.Add(self.cb_comp_b, 1, wx.EXPAND)

        vbox.Add(grid, 1, wx.ALL|wx.EXPAND, 12)

        # 绑定事件：选择参考列时自动填充
        self.cb_id_ref.Bind(wx.EVT_COMBOBOX, self.on_id_ref_select)

        # 提示信息
        hint = wx.StaticText(panel, label="提示：可直接输入列名，或从下拉列表选择。右侧为表B的参考列（仅ID列）")
        hint.SetForegroundColour(wx.Colour(100, 100, 100))
        vbox.Add(hint, 0, wx.ALL, 8)

        # 预填充参考列
        if config.id_col and config.id_col in cols_b:
            self.cb_id_ref.SetStringSelection(config.id_col)

        # 底部按钮
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        btn_ok = wx.Button(panel, wx.ID_OK, label="保存")
        btn_cancel = wx.Button(panel, wx.ID_CANCEL, label="取消")
        hbox.AddStretchSpacer(1)
        hbox.Add(btn_ok, 0, wx.ALL, 5)
        hbox.Add(btn_cancel, 0, wx.ALL, 5)
        vbox.Add(hbox, 0, wx.EXPAND|wx.ALL, 8)

        panel.SetSizer(vbox)

    def on_id_ref_select(self, event):
        """当选择参考ID列时，自动填充到ID列输入框"""
        selected = self.cb_id_ref.GetStringSelection()
        if selected:
            self.cb_id.SetValue(selected)

    def get_config(self) -> Optional[GlobalConfig]:
        id_col = self.cb_id.GetValue().strip()
        comp_a = self.cb_comp_a.GetValue().strip()
        comp_b = self.cb_comp_b.GetValue().strip()
        
        if not id_col or not comp_a or not comp_b:
            return None
        return GlobalConfig(id_col=id_col, comp_a_col=comp_a, comp_b_col=comp_b)


# ----------------------------- 列配置对话框 ----------------------------- #

class SheetConfigDialog(wx.Dialog):
    def __init__(self, parent, title: str, sheet_name: str,
                 cols_a: List[str], cols_b: List[str], config: SheetConfig):
        super().__init__(parent, title=title, size=(520, 320))
        self.sheet_name = sheet_name
        self.cols_a = cols_a
        self.cols_b = cols_b
        self.config = config

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        tip = wx.StaticText(panel, label=f"配置 Sheet：{sheet_name}")
        font = tip.GetFont()
        font.MakeBold()
        tip.SetFont(font)
        vbox.Add(tip, 0, wx.ALL, 10)

        grid = wx.FlexGridSizer(3, 3, 10, 10)
        grid.AddGrowableCol(1, 1)
        grid.AddGrowableCol(2, 1)

        # 行1：ID 列
        grid.Add(wx.StaticText(panel, label="ID 列 (A/B 共用或任选)："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.cb_id_a = wx.ComboBox(panel, choices=cols_a, style=wx.CB_READONLY)
        self.cb_id_b = wx.ComboBox(panel, choices=cols_b, style=wx.CB_READONLY)
        grid.Add(self.cb_id_a, 1, wx.EXPAND)
        grid.Add(self.cb_id_b, 1, wx.EXPAND)

        # 行2：表A对比列
        grid.Add(wx.StaticText(panel, label="表A 对比列："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.cb_comp_a = wx.ComboBox(panel, choices=cols_a, style=wx.CB_READONLY)
        grid.Add(self.cb_comp_a, 1, wx.EXPAND)
        grid.AddSpacer(1)

        # 行3：表B对比列
        grid.Add(wx.StaticText(panel, label="表B 对比列："), 0, wx.ALIGN_CENTER_VERTICAL)
        self.cb_comp_b = wx.ComboBox(panel, choices=cols_b, style=wx.CB_READONLY)
        grid.AddSpacer(1)
        grid.Add(self.cb_comp_b, 1, wx.EXPAND)

        vbox.Add(grid, 1, wx.ALL|wx.EXPAND, 12)

        # 预填
        def set_if_in(cb: wx.ComboBox, val: Optional[str], pool: List[str]):
            if val and val in pool:
                cb.SetStringSelection(val)

        set_if_in(self.cb_id_a, config.id_col, cols_a)
        set_if_in(self.cb_id_b, config.id_col, cols_b)  # ID 可从 A 或 B 任选
        set_if_in(self.cb_comp_a, config.comp_a_col, cols_a)
        set_if_in(self.cb_comp_b, config.comp_b_col, cols_b)

        # 底部按钮
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        btn_ok = wx.Button(panel, wx.ID_OK, label="保存")
        btn_cancel = wx.Button(panel, wx.ID_CANCEL, label="取消")
        hbox.AddStretchSpacer(1)
        hbox.Add(btn_ok, 0, wx.ALL, 5)
        hbox.Add(btn_cancel, 0, wx.ALL, 5)
        vbox.Add(hbox, 0, wx.EXPAND|wx.ALL, 8)

        panel.SetSizer(vbox)

    def get_config(self) -> Optional[SheetConfig]:
        # 以 A/B 任意一个选择为准，优先 A；若 A 未选则取 B
        id_a = self.cb_id_a.GetStringSelection() if self.cb_id_a.GetSelection()!=wx.NOT_FOUND else None
        id_b = self.cb_id_b.GetStringSelection() if self.cb_id_b.GetSelection()!=wx.NOT_FOUND else None
        id_col = id_a or id_b
        comp_a = self.cb_comp_a.GetStringSelection() if self.cb_comp_a.GetSelection()!=wx.NOT_FOUND else None
        comp_b = self.cb_comp_b.GetStringSelection() if self.cb_comp_b.GetSelection()!=wx.NOT_FOUND else None
        if not id_col or not comp_a or not comp_b:
            return None
        return SheetConfig(sheet_name=self.sheet_name, id_col=id_col, comp_a_col=comp_a, comp_b_col=comp_b)


# ----------------------------- 主窗口 ----------------------------- #

class ExcelDiffFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="Excel 对比工具", size=(1100, 720))
        self.CenterOnScreen()

        self.path_a: Optional[str] = None
        self.path_b: Optional[str] = None
        self.sheets_a: List[str] = []
        self.sheets_b: List[str] = []
        self.common_sheets: List[str] = []
        self.sheet_configs: Dict[str, SheetConfig] = {}
        self.global_config = GlobalConfig()  # 全局统一配置
        self.compare_options = CompareOptions()
        self.results: List[CompareRow] = []
        
        # 缓存所有Sheet的列名，避免重复读取
        self.sheet_columns_a: Dict[str, List[str]] = {}  # {sheet_name: [col1, col2, ...]}
        self.sheet_columns_b: Dict[str, List[str]] = {}

        self._build_ui()

    # -------- UI 组装 -------- #
    def _build_ui(self):
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        # 顶部：文件选择区
        file_box = wx.StaticBox(panel, label="选择文件")
        file_sizer = wx.StaticBoxSizer(file_box, wx.VERTICAL)

        # 表 A
        h1 = wx.BoxSizer(wx.HORIZONTAL)
        self.fp_a = wx.FilePickerCtrl(panel, message="选择表A (xlsx/xls)", wildcard="Excel files (*.xlsx;*.xls)|*.xlsx;*.xls")
        btn_load_a = wx.Button(panel, label="加载表A Sheet")
        btn_load_a.Bind(wx.EVT_BUTTON, self.on_load_a)
        h1.Add(wx.StaticText(panel, label="表 A："), 0, wx.ALIGN_CENTER_VERTICAL|wx.RIGHT, 6)
        h1.Add(self.fp_a, 1, wx.EXPAND|wx.RIGHT, 6)
        h1.Add(btn_load_a, 0)
        file_sizer.Add(h1, 0, wx.EXPAND|wx.ALL, 6)

        # 表 B
        h2 = wx.BoxSizer(wx.HORIZONTAL)
        self.fp_b = wx.FilePickerCtrl(panel, message="选择表B (xlsx/xls)", wildcard="Excel files (*.xlsx;*.xls)|*.xlsx;*.xls")
        btn_load_b = wx.Button(panel, label="加载表B Sheet")
        btn_load_b.Bind(wx.EVT_BUTTON, self.on_load_b)
        h2.Add(wx.StaticText(panel, label="表 B："), 0, wx.ALIGN_CENTER_VERTICAL|wx.RIGHT, 6)
        h2.Add(self.fp_b, 1, wx.EXPAND|wx.RIGHT, 6)
        h2.Add(btn_load_b, 0)
        file_sizer.Add(h2, 0, wx.EXPAND|wx.ALL, 6)

        vbox.Add(file_sizer, 0, wx.EXPAND|wx.ALL, 8)

        # 中部：Sheet 选择与配置
        mid_box = wx.StaticBox(panel, label="Sheet 选择与列配置")
        mid_sizer = wx.StaticBoxSizer(mid_box, wx.HORIZONTAL)

        left = wx.BoxSizer(wx.VERTICAL)
        left.Add(wx.StaticText(panel, label="可选 Sheet："), 0, wx.BOTTOM, 5)
        self.clb_sheets = wx.CheckListBox(panel, choices=[])
        self.clb_sheets.Bind(wx.EVT_LISTBOX_DCLICK, self.on_config_sheet_dclick)
        left.Add(self.clb_sheets, 1, wx.EXPAND)

        btns = wx.BoxSizer(wx.VERTICAL)
        
        # 第一行按钮
        btns1 = wx.BoxSizer(wx.HORIZONTAL)
        self.btn_select_all = wx.Button(panel, label="全选")
        self.btn_clear_sel = wx.Button(panel, label="全不选")
        self.btn_select_all.Bind(wx.EVT_BUTTON, self.on_select_all)
        self.btn_clear_sel.Bind(wx.EVT_BUTTON, self.on_clear_all)
        btns1.Add(self.btn_select_all, 0, wx.RIGHT, 6)
        btns1.Add(self.btn_clear_sel, 0)
        
        # 第二行按钮
        btns2 = wx.BoxSizer(wx.HORIZONTAL)
        self.btn_global_config = wx.Button(panel, label="统一配置所有表格…")
        self.btn_config = wx.Button(panel, label="单独配置选中表格…")
        self.btn_global_config.Bind(wx.EVT_BUTTON, self.on_global_config)
        self.btn_config.Bind(wx.EVT_BUTTON, self.on_config_selected)
        btns2.Add(self.btn_global_config, 1, wx.EXPAND|wx.RIGHT, 6)
        btns2.Add(self.btn_config, 1, wx.EXPAND)
        
        btns.Add(btns1, 0, wx.EXPAND|wx.BOTTOM, 6)
        btns.Add(btns2, 0, wx.EXPAND)
        left.Add(btns, 0, wx.TOP, 6)

        mid_sizer.Add(left, 1, wx.ALL|wx.EXPAND, 6)

        # 右侧：对比选项 + 操作
        right = wx.BoxSizer(wx.VERTICAL)
        opt_box = wx.StaticBox(panel, label="对比选项")
        opt_sizer = wx.StaticBoxSizer(opt_box, wx.VERTICAL)
        self.cb_ignore_case = wx.CheckBox(panel, label="忽略大小写")
        self.cb_ignore_case.SetValue(True)
        self.cb_strip = wx.CheckBox(panel, label="去除前后空格")
        self.cb_strip.SetValue(True)
        opt_sizer.Add(self.cb_ignore_case, 0, wx.ALL, 4)
        opt_sizer.Add(self.cb_strip, 0, wx.ALL, 4)
        right.Add(opt_sizer, 0, wx.EXPAND|wx.BOTTOM, 8)

        run_box = wx.StaticBox(panel, label="执行")
        run_sizer = wx.StaticBoxSizer(run_box, wx.VERTICAL)
        self.btn_compare = wx.Button(panel, label="开始比对")
        self.btn_export = wx.Button(panel, label="导出差异报告…")
        self.btn_export.Disable()
        self.btn_compare.Bind(wx.EVT_BUTTON, self.on_compare)
        self.btn_export.Bind(wx.EVT_BUTTON, self.on_export)
        self.gauge = wx.Gauge(panel, range=100, size=(250, -1))
        run_sizer.Add(self.btn_compare, 0, wx.ALL|wx.EXPAND, 4)
        run_sizer.Add(self.gauge, 0, wx.ALL|wx.EXPAND, 4)
        run_sizer.Add(self.btn_export, 0, wx.ALL|wx.EXPAND, 4)
        right.Add(run_sizer, 0, wx.EXPAND)

        mid_sizer.Add(right, 0, wx.ALL|wx.EXPAND, 6)

        vbox.Add(mid_sizer, 1, wx.EXPAND|wx.LEFT|wx.RIGHT|wx.BOTTOM, 8)

        # 下部：日志区域
        self.txt_log = wx.TextCtrl(panel, style=wx.TE_MULTILINE|wx.TE_READONLY)
        vbox.Add(self.txt_log, 1, wx.EXPAND|wx.ALL, 8)

        panel.SetSizer(vbox)

    # -------- 事件处理 -------- #

    def log(self, msg: str):
        self.txt_log.AppendText(msg + "\n")

    def on_load_a(self, event):
        path = self.fp_a.GetPath()
        if not path:
            wx.MessageBox("请先选择表A文件", "提示", wx.OK|wx.ICON_INFORMATION)
            return
        try:
            # 清理缓存以释放内存
            if hasattr(self, 'path_a') and self.path_a != path:
                excel_cache.clear_cache()
                self.sheet_columns_a.clear()
            
            x = pd.ExcelFile(path)
            self.path_a = path
            self.sheets_a = list(x.sheet_names)
            
            # 批量预加载所有Sheet的列名到缓存
            self.log("正在预加载表A的列信息...")
            self.sheet_columns_a.clear()
            for sheet in self.sheets_a:
                try:
                    cols = safe_read_excel_columns(path, sheet)
                    self.sheet_columns_a[sheet] = cols
                except Exception as e:
                    self.log(f"读取表A Sheet '{sheet}' 列名失败：{e}")
                    self.sheet_columns_a[sheet] = []
            
            x.close()  # 显式关闭文件句柄
            self.log(f"表A加载成功，共 {len(self.sheets_a)} 个Sheet，列信息已缓存")
            self.refresh_sheet_list()
        except Exception as e:
            self.log("加载表A失败：" + str(e))
            self.log(traceback.format_exc())

    def on_load_b(self, event):
        path = self.fp_b.GetPath()
        if not path:
            wx.MessageBox("请先选择表B文件", "提示", wx.OK|wx.ICON_INFORMATION)
            return
        try:
            # 清理缓存以释放内存
            if hasattr(self, 'path_b') and self.path_b != path:
                excel_cache.clear_cache()
                self.sheet_columns_b.clear()
                
            x = pd.ExcelFile(path)
            self.path_b = path
            self.sheets_b = list(x.sheet_names)
            
            # 批量预加载所有Sheet的列名到缓存
            self.log("正在预加载表B的列信息...")
            self.sheet_columns_b.clear()
            for sheet in self.sheets_b:
                try:
                    cols = safe_read_excel_columns(path, sheet)
                    self.sheet_columns_b[sheet] = cols
                except Exception as e:
                    self.log(f"读取表B Sheet '{sheet}' 列名失败：{e}")
                    self.sheet_columns_b[sheet] = []
            
            x.close()  # 显式关闭文件句柄
            self.log(f"表B加载成功，共 {len(self.sheets_b)} 个Sheet，列信息已缓存")
            self.refresh_sheet_list()
        except Exception as e:
            self.log("加载表B失败：" + str(e))
            self.log(traceback.format_exc())

    def refresh_sheet_list(self):
        self.clb_sheets.Clear()
        self.common_sheets = []
        if self.sheets_a and self.sheets_b:
            aset = set(self.sheets_a)
            bset = set(self.sheets_b)
            self.common_sheets = sorted(list(aset & bset))
            
            # 显示详细的Sheet统计信息
            self.log(f"Sheet统计：表A有{len(self.sheets_a)}个Sheet，表B有{len(self.sheets_b)}个Sheet")
            self.log(f"共有Sheet数量：{len(self.common_sheets)}个")
            
            if not self.common_sheets:
                self.log("⚠️ 表A与表B没有共同的Sheet名称。")
            else:
                for s in self.common_sheets:
                    self.clb_sheets.Append(s)
                # 默认全选
                for i in range(self.clb_sheets.GetCount()):
                    self.clb_sheets.Check(i, True)
        else:
            self.log("请先分别加载表A与表B，以显示共同的Sheet。")

    def _get_selected_sheets(self) -> List[str]:
        res = []
        for i in range(self.clb_sheets.GetCount()):
            if self.clb_sheets.IsChecked(i):
                res.append(self.clb_sheets.GetString(i))
        return res

    def on_select_all(self, event):
        for i in range(self.clb_sheets.GetCount()):
            self.clb_sheets.Check(i, True)

    def on_clear_all(self, event):
        for i in range(self.clb_sheets.GetCount()):
            self.clb_sheets.Check(i, False)

    def on_config_sheet_dclick(self, event):
        idx = event.GetSelection()
        if idx != wx.NOT_FOUND:
            sheet = self.clb_sheets.GetString(idx)
            self.open_config_dialog(sheet)

    def on_global_config(self, event):
        """全局统一配置"""
        if not (self.path_a and self.path_b):
            wx.MessageBox("请先选择并加载表A与表B", "提示")
            return
        
        selected = self._get_selected_sheets()
        if not selected:
            wx.MessageBox("请先勾选至少一个 Sheet。", "提示")
            return
        
        # 使用缓存的列名（无需重新读取）
        first_sheet = selected[0]
        cols_a = self.sheet_columns_a.get(first_sheet, [])
        cols_b = self.sheet_columns_b.get(first_sheet, [])
        
        if not cols_a or not cols_b:
            self.log(f"未找到Sheet '{first_sheet}' 的列信息，请重新加载文件")
            return
        
        dlg = GlobalConfigDialog(self, title="统一配置", cols_a=cols_a, cols_b=cols_b, config=self.global_config)
        if dlg.ShowModal() == wx.ID_OK:
            new_config = dlg.get_config()
            if new_config is None:
                wx.MessageBox("请完整填写 ID 列、表A对比列、表B对比列的名称。", "提示")
            else:
                self.global_config = new_config
                # 应用到所有选中的sheet
                for sheet in selected:
                    self.sheet_configs[sheet] = SheetConfig(
                        sheet_name=sheet,
                        id_col=new_config.id_col,
                        comp_a_col=new_config.comp_a_col,
                        comp_b_col=new_config.comp_b_col
                    )
                self.log(f"统一配置已应用到 {len(selected)} 个表格：ID={new_config.id_col}  A列={new_config.comp_a_col}  B列={new_config.comp_b_col}")
        dlg.Destroy()

    def on_config_selected(self, event):
        selected = self._get_selected_sheets()
        if not selected:
            wx.MessageBox("请先勾选至少一个 Sheet。", "提示")
            return
        for s in selected:
            self.open_config_dialog(s)

    def open_config_dialog(self, sheet: str):
        if not (self.path_a and self.path_b):
            wx.MessageBox("请先选择并加载表A与表B", "提示")
            return
        
        # 使用缓存的列名（无需重新读取）
        cols_a = self.sheet_columns_a.get(sheet, [])
        cols_b = self.sheet_columns_b.get(sheet, [])
        
        if not cols_a or not cols_b:
            self.log(f"未找到Sheet '{sheet}' 的列信息，请重新加载文件")
            return

        cfg = self.sheet_configs.get(sheet, SheetConfig(sheet_name=sheet))
        dlg = SheetConfigDialog(self, title="配置列", sheet_name=sheet,
                                cols_a=cols_a, cols_b=cols_b, config=cfg)
        if dlg.ShowModal() == wx.ID_OK:
            new_cfg = dlg.get_config()
            if new_cfg is None:
                wx.MessageBox("请完整选择 ID 列、表A对比列、表B对比列。", "提示")
            else:
                self.sheet_configs[sheet] = new_cfg
                self.log(f"保存配置：{sheet}  ID={new_cfg.id_col}  A列={new_cfg.comp_a_col}  B列={new_cfg.comp_b_col}")
        dlg.Destroy()

    def on_compare(self, event):
        if not (self.path_a and self.path_b):
            wx.MessageBox("请先选择并加载表A与表B", "提示")
            return
        sel = self._get_selected_sheets()
        if not sel:
            wx.MessageBox("请至少勾选一个需要对比的 Sheet", "提示")
            return
        
        # 更新选项
        self.compare_options.ignore_case = self.cb_ignore_case.GetValue()
        self.compare_options.strip_spaces = self.cb_strip.GetValue()

        self.results.clear()

        self.btn_compare.Disable()
        self.btn_export.Disable()
        self.gauge.SetValue(0)
        
        # 启动后台线程进行对比
        self.compare_thread = threading.Thread(target=self._compare_worker, args=(sel,))
        self.compare_thread.daemon = True
        self.compare_thread.start()
        
        # 启动定时器更新UI
        self.compare_timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.on_compare_timer)
        self.compare_timer.Start(100)  # 每100ms更新一次

    def _compare_worker(self, selected_sheets: List[str]):
        """后台线程执行对比工作，支持多线程并行处理"""
        try:
            self.compare_progress = 0
            self.compare_status = "准备中..."
            self.compare_error = None
            
            total = len(selected_sheets)
            
            # 大量Sheet时，先预加载数据到缓存
            if total > 20:
                self.compare_status = "预加载数据中..."
                self._preload_sheet_data(selected_sheets)
                self.compare_progress = 10
            
            # 根据Sheet数量决定是否使用多线程
            if total <= 2:
                # 少量Sheet，使用单线程避免开销
                for i, sheet in enumerate(selected_sheets):
                    remaining = total - i
                    self.compare_status = f"正在对比：{sheet}（剩余 {remaining-1} 个表格）"
                    self.compare_one_sheet(sheet)
                    self.compare_progress = int((i + 1) * 90 / max(total, 1)) + 10
            else:
                # 多个Sheet，使用多线程并行处理
                self.compare_status = "多线程并行对比中..."
                self._parallel_compare_sheets(selected_sheets)
            
            self.compare_status = "对比完成"
            self.compare_progress = 100
        except Exception as e:
            self.compare_error = str(e)
            self.compare_status = f"对比出错：{str(e)}"

    def _preload_sheet_data(self, selected_sheets: List[str]):
        """预加载所有需要的Sheet数据到缓存"""
        try:
            # 批量预加载表A和表B的数据
            for sheet in selected_sheets:
                cfg = self.sheet_configs.get(sheet)
                if cfg and cfg.id_col and cfg.comp_a_col and cfg.comp_b_col:
                    # 预加载到缓存，后续使用时直接从缓存读取
                    if sheet in self.sheets_a:
                        excel_cache.get_data(self.path_a, sheet, [cfg.id_col, cfg.comp_a_col], filter_empty_ids=True)
                    if sheet in self.sheets_b:
                        excel_cache.get_data(self.path_b, sheet, [cfg.id_col, cfg.comp_b_col], filter_empty_ids=True)
        except Exception as e:
            self.log(f"预加载数据时出错：{str(e)}")

    def _parallel_compare_sheets(self, selected_sheets: List[str]):
        """多线程并行处理多个Sheet"""
        import threading
        
        # 使用线程锁保护共享资源
        results_lock = threading.Lock()
        progress_lock = threading.Lock()
        
        completed_count = 0
        total_count = len(selected_sheets)
        
        # 更激进的线程数量配置
        import multiprocessing
        cpu_count = multiprocessing.cpu_count()
        
        if total_count <= 5:
            max_workers = min(4, total_count)
        elif total_count <= 20:
            max_workers = min(cpu_count * 2, total_count)  # CPU核心数的2倍
        elif total_count <= 100:
            max_workers = min(cpu_count * 4, total_count)  # CPU核心数的4倍
        else:
            max_workers = min(cpu_count * 6, 64, total_count)  # 最多64个线程
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # 提交所有任务
            future_to_sheet = {
                executor.submit(self._compare_one_sheet_threadsafe, sheet, results_lock): sheet 
                for sheet in selected_sheets
            }
            
            # 收集结果并更新进度
            for future in as_completed(future_to_sheet):
                sheet = future_to_sheet[future]
                try:
                    future.result()  # 获取结果，如果有异常会抛出
                    
                    with progress_lock:
                        completed_count += 1
                        remaining_count = total_count - completed_count
                        self.compare_progress = int(completed_count * 100 / total_count)
                        self.compare_status = f"已完成 {completed_count}/{total_count} 个Sheet（剩余 {remaining_count} 个）"
                        
                except Exception as e:
                    self.compare_error = f"处理{sheet}时出错：{str(e)}"
                    break

    def _compare_one_sheet_threadsafe(self, sheet: str, results_lock: threading.Lock):
        """线程安全的单Sheet对比函数"""
        # 检查配置
        cfg = self.sheet_configs.get(sheet)
        if not cfg or not (cfg.id_col and cfg.comp_a_col and cfg.comp_b_col):
            return

        # 如果B没有该sheet
        if sheet not in self.sheets_b:
            with results_lock:
                self.results.append(CompareRow(sheet, None, "", "", "", "Sheet在B不存在"))
            return

        try:
            # 使用优化的数据加载函数（加上预过滤）
            df_a = load_needed_columns_optimized(self.path_a, sheet, [cfg.id_col, cfg.comp_a_col], filter_empty_ids=True)
            df_b = load_needed_columns_optimized(self.path_b, sheet, [cfg.id_col, cfg.comp_b_col], filter_empty_ids=True)
            
            if df_a.empty or df_b.empty:
                return

            # 重命名列以避免冲突
            a_id, a_val = cfg.id_col, cfg.comp_a_col
            b_id, b_val = cfg.id_col, cfg.comp_b_col
            df_a = df_a.rename(columns={a_id: "__ID__", a_val: "__A__"})
            df_b = df_b.rename(columns={b_id: "__ID__", b_val: "__B__"})

            # 统一为字符串类型
            df_a["__ID__"] = df_a["__ID__"].astype(str).fillna("")
            df_a["__A__"] = df_a["__A__"].astype(str).fillna("")
            df_b["__ID__"] = df_b["__ID__"].astype(str).fillna("")
            df_b["__B__"] = df_b["__B__"].astype(str).fillna("")

            # 创建临时结果列表
            temp_results = []
            
            # 根据数据量选择最优算法
            total_rows = len(df_a) + len(df_b)
            
            if total_rows < 500:
                # 小数据集：使用简单完整算法
                self._simple_full_compare(df_a, df_b, sheet, temp_results)
            else:
                # 大数据集：使用向量化算法
                self._compare_with_vectorized_operations_local(df_a, df_b, sheet, temp_results)
            
            # 线程安全地添加结果
            with results_lock:
                self.results.extend(temp_results)
                
        except Exception as e:
            # 线程安全地设置错误
            with results_lock:
                if not hasattr(self, 'compare_error') or not self.compare_error:
                    self.compare_error = f"处理{sheet}时出错：{str(e)}"

    def on_compare_timer(self, event):
        """定时器更新UI状态"""
        if hasattr(self, 'compare_progress'):
            self.gauge.SetValue(self.compare_progress)
            
        if hasattr(self, 'compare_status'):
            # 更新状态到日志（但不重复相同消息）
            if not hasattr(self, '_last_status') or self._last_status != self.compare_status:
                if "正在对比：" in self.compare_status or "已完成" in self.compare_status:
                    self.log(self.compare_status)
                self._last_status = self.compare_status
        
        # 检查是否完成
        if hasattr(self, 'compare_progress') and self.compare_progress >= 100:
            self.compare_timer.Stop()
            
            if hasattr(self, 'compare_error') and self.compare_error:
                self.log(f"对比过程中出错：{self.compare_error}")
            else:
                self.log_results_summary(self.results)
                self.btn_export.Enable(True)
                self.log("全部对比完成。")
            
            self.btn_compare.Enable(True)
            
            # 清理临时属性
            if hasattr(self, 'compare_progress'):
                delattr(self, 'compare_progress')
            if hasattr(self, 'compare_status'):
                delattr(self, 'compare_status')
            if hasattr(self, 'compare_error'):
                delattr(self, 'compare_error')
            if hasattr(self, '_last_status'):
                delattr(self, '_last_status')

    def compare_one_sheet(self, sheet: str):
        # 检查配置
        cfg = self.sheet_configs.get(sheet)
        if not cfg or not (cfg.id_col and cfg.comp_a_col and cfg.comp_b_col):
            return

        # 如果B没有该sheet
        if sheet not in self.sheets_b:
            self.results.append(CompareRow(sheet, None, "", "", "", "Sheet在B不存在"))
            return

        try:
            # 使用优化的数据加载函数
            df_a = load_needed_columns_optimized(self.path_a, sheet, [cfg.id_col, cfg.comp_a_col])
            df_b = load_needed_columns_optimized(self.path_b, sheet, [cfg.id_col, cfg.comp_b_col])
            if df_a.empty or df_b.empty:
                return

            # 重命名列以避免冲突
            a_id, a_val = cfg.id_col, cfg.comp_a_col
            b_id, b_val = cfg.id_col, cfg.comp_b_col
            df_a = df_a.rename(columns={a_id: "__ID__", a_val: "__A__"})
            df_b = df_b.rename(columns={b_id: "__ID__", b_val: "__B__"})

            # 统一为字符串类型
            df_a["__ID__"] = df_a["__ID__"].astype(str).fillna("")
            df_a["__A__"] = df_a["__A__"].astype(str).fillna("")
            df_b["__ID__"] = df_b["__ID__"].astype(str).fillna("")
            df_b["__B__"] = df_b["__B__"].astype(str).fillna("")

            # 根据数据量选择最优算法
            total_rows = len(df_a) + len(df_b)
            
            if total_rows < 500:
                # 小数据集：使用简单完整算法
                self._simple_full_compare_single(df_a, df_b, sheet)
            else:
                # 大数据集：使用向量化算法
                self._compare_with_vectorized_operations(df_a, df_b, sheet)
                
        except Exception as e:
            # 在后台线程中，不能直接操作UI，只能设置错误状态
            if hasattr(self, 'compare_error'):
                self.compare_error = f"处理{sheet}时出错：{str(e)}"

    def _compare_with_vectorized_operations(self, df_a: pd.DataFrame, df_b: pd.DataFrame, sheet: str):
        """使用完全向量化操作的超高速对比算法"""
        opt = self.compare_options
        
        if df_a.empty:
            return
            
        # 预处理表A：向量化标准化
        df_a_clean = df_a.copy()
        df_a_clean["__ID_CLEAN__"] = df_a_clean["__ID__"].astype(str).str.strip()
        df_a_clean["__A_CLEAN__"] = df_a_clean["__A__"].astype(str)
        
        # 向量化标准化表A的对比值
        a_norm = df_a_clean["__A_CLEAN__"]
        if opt.strip_spaces:
            a_norm = a_norm.str.strip()
        if opt.ignore_case:
            a_norm = a_norm.str.casefold()
        df_a_clean["__A_NORM__"] = a_norm
        
        # 预处理表B：向量化标准化
        df_b_clean = df_b.copy()
        df_b_clean["__ID_CLEAN__"] = df_b_clean["__ID__"].astype(str).str.strip()
        df_b_clean["__B_CLEAN__"] = df_b_clean["__B__"].astype(str)
        
        # 向量化标准化表B的对比值
        b_norm = df_b_clean["__B_CLEAN__"]
        if opt.strip_spaces:
            b_norm = b_norm.str.strip()
        if opt.ignore_case:
            b_norm = b_norm.str.casefold()
        df_b_clean["__B_NORM__"] = b_norm
        
        # 使用pandas merge进行高效连接（比字典查找更快）
        merged = df_a_clean.reset_index().merge(
            df_b_clean[["__ID_CLEAN__", "__B_CLEAN__", "__B_NORM__"]], 
            on="__ID_CLEAN__", 
            how="left"
        )
        
        # 向量化状态判断
        conditions = [
            merged["__B_CLEAN__"].isna() & (merged["__ID_CLEAN__"] != ""),  # 表B无此ID
            merged["__A_NORM__"] == merged["__B_NORM__"],  # 一致
        ]
        choices = ["表B无此ID", "一致"]
        merged["status"] = np.select(conditions, choices, default="不一致")
        
        # 批量创建结果，避免逐行处理
        batch_size = 5000  # 大批次处理，减少开销
        total_rows = len(merged)
        
        for start_idx in range(0, total_rows, batch_size):
            end_idx = min(start_idx + batch_size, total_rows)
            batch = merged.iloc[start_idx:end_idx]
            
            # 向量化创建CompareRow对象
            batch_results = [
                CompareRow(
                    sheet=sheet,
                    row_index_a=int(row["index"]) + 2,  # Excel行号
                    id_value=row["__ID_CLEAN__"],
                    a_value=row["__A_CLEAN__"],
                    b_value=row["__B_CLEAN__"] if pd.notna(row["__B_CLEAN__"]) else "",
                    status=row["status"],
                )
                for _, row in batch.iterrows()
            ]
            
            # 批量添加结果
            self.results.extend(batch_results)
            
            # 减少sleep频率
            if start_idx % (batch_size * 2) == 0:
                time.sleep(0.0005)  # 更短的sleep时间

    def _compare_with_vectorized_operations_local(self, df_a: pd.DataFrame, df_b: pd.DataFrame, sheet: str, temp_results: List):
        """线程安全版本：使用完全向量化操作的超高速对比算法"""
        opt = self.compare_options
        
        if df_a.empty:
            return
            
        # 预处理表A：向量化标准化
        df_a_clean = df_a.copy()
        df_a_clean["__ID_CLEAN__"] = df_a_clean["__ID__"].astype(str).str.strip()
        df_a_clean["__A_CLEAN__"] = df_a_clean["__A__"].astype(str)
        
        # 向量化标准化表A的对比值
        a_norm = df_a_clean["__A_CLEAN__"]
        if opt.strip_spaces:
            a_norm = a_norm.str.strip()
        if opt.ignore_case:
            a_norm = a_norm.str.casefold()
        df_a_clean["__A_NORM__"] = a_norm
        
        # 预处理表B：向量化标准化
        df_b_clean = df_b.copy()
        df_b_clean["__ID_CLEAN__"] = df_b_clean["__ID__"].astype(str).str.strip()
        df_b_clean["__B_CLEAN__"] = df_b_clean["__B__"].astype(str)
        
        # 向量化标准化表B的对比值
        b_norm = df_b_clean["__B_CLEAN__"]
        if opt.strip_spaces:
            b_norm = b_norm.str.strip()
        if opt.ignore_case:
            b_norm = b_norm.str.casefold()
        df_b_clean["__B_NORM__"] = b_norm
        
        # 使用pandas merge进行高效连接
        merged = df_a_clean.reset_index().merge(
            df_b_clean[["__ID_CLEAN__", "__B_CLEAN__", "__B_NORM__"]], 
            on="__ID_CLEAN__", 
            how="left"
        )
        
        # 向量化状态判断
        conditions = [
            merged["__B_CLEAN__"].isna() & (merged["__ID_CLEAN__"] != ""),  # 表B无此ID
            merged["__A_NORM__"] == merged["__B_NORM__"],  # 一致
        ]
        choices = ["表B无此ID", "一致"]
        merged["status"] = np.select(conditions, choices, default="不一致")
        
        # 批量创建结果
        batch_size = 10000  # 更大批次，减少开销
        total_rows = len(merged)
        
        for start_idx in range(0, total_rows, batch_size):
            end_idx = min(start_idx + batch_size, total_rows)
            batch = merged.iloc[start_idx:end_idx]
            
            # 向量化创建CompareRow对象
            batch_results = [
                CompareRow(
                    sheet=sheet,
                    row_index_a=int(row["index"]) + 2,
                    id_value=row["__ID_CLEAN__"],
                    a_value=row["__A_CLEAN__"],
                    b_value=row["__B_CLEAN__"] if pd.notna(row["__B_CLEAN__"]) else "",
                    status=row["status"],
                )
                for _, row in batch.iterrows()
            ]
            
            temp_results.extend(batch_results)





    def _simple_full_compare(self, df_a: pd.DataFrame, df_b: pd.DataFrame, sheet: str, temp_results: List):
        """针对小数据集的简单完整对比算法"""
        opt = self.compare_options
        
        if df_a.empty:
            return
            
        # 构建B表字典
        b_dict = {}
        for _, row in df_b.iterrows():
            id_val = str(row["__ID__"]).strip()
            b_val = str(row["__B__"])
            if id_val:
                b_dict[id_val] = b_val
        
        # 逐行对比（小数据集可以接受）
        for idx, row in df_a.iterrows():
            id_val = str(row["__ID__"]).strip()
            a_val = str(row["__A__"])
            
            if not id_val:
                continue
                
            if id_val in b_dict:
                b_val = b_dict[id_val]
                
                # 标准化比较
                a_norm = a_val
                b_norm = b_val
                if opt.strip_spaces:
                    a_norm = a_norm.strip()
                    b_norm = b_norm.strip()
                if opt.ignore_case:
                    a_norm = a_norm.casefold()
                    b_norm = b_norm.casefold()
                
                status = "一致" if a_norm == b_norm else "不一致"
            else:
                b_val = ""
                status = "表B无此ID"
            
            temp_results.append(CompareRow(
                sheet=sheet,
                row_index_a=idx + 2,
                id_value=id_val,
                a_value=a_val,
                b_value=b_val,
                status=status,
            ))

    def _simple_full_compare_single(self, df_a: pd.DataFrame, df_b: pd.DataFrame, sheet: str):
        """单线程版本的简单完整对比算法"""
        opt = self.compare_options
        
        if df_a.empty:
            return
            
        # 构建B表字典
        b_dict = {}
        for _, row in df_b.iterrows():
            id_val = str(row["__ID__"]).strip()
            b_val = str(row["__B__"])
            if id_val:
                b_dict[id_val] = b_val
        
        # 逐行对比（小数据集可以接受）
        for idx, row in df_a.iterrows():
            id_val = str(row["__ID__"]).strip()
            a_val = str(row["__A__"])
            
            if not id_val:
                continue
                
            if id_val in b_dict:
                b_val = b_dict[id_val]
                
                # 标准化比较
                a_norm = a_val
                b_norm = b_val
                if opt.strip_spaces:
                    a_norm = a_norm.strip()
                    b_norm = b_norm.strip()
                if opt.ignore_case:
                    a_norm = a_norm.casefold()
                    b_norm = b_norm.casefold()
                
                status = "一致" if a_norm == b_norm else "不一致"
            else:
                b_val = ""
                status = "表B无此ID"
            
            self.results.append(CompareRow(
                sheet=sheet,
                row_index_a=idx + 2,
                id_value=id_val,
                a_value=a_val,
                b_value=b_val,
                status=status,
            ))

    def _find_missing_in_a(self, compared_sheets: List[str]) -> List[dict]:
        """查找B中有但A中没有的ID"""
        missing_in_a_details = []
        
        if not compared_sheets:
            return missing_in_a_details
        
        for sheet in compared_sheets:
            try:
                cfg = self.sheet_configs.get(sheet)
                if not cfg or not (cfg.id_col and cfg.comp_a_col and cfg.comp_b_col):
                    continue
                    
                if sheet not in self.sheets_a or sheet not in self.sheets_b:
                    continue
                
                # 加载数据
                df_a = load_needed_columns_optimized(self.path_a, sheet, [cfg.id_col, cfg.comp_a_col], filter_empty_ids=True)
                df_b = load_needed_columns_optimized(self.path_b, sheet, [cfg.id_col, cfg.comp_b_col], filter_empty_ids=True)
                
                if df_a.empty or df_b.empty:
                    continue
                
                # 重命名列
                a_id, a_val = cfg.id_col, cfg.comp_a_col
                b_id, b_val = cfg.id_col, cfg.comp_b_col
                df_a = df_a.rename(columns={a_id: "__ID__", a_val: "__A__"})
                df_b = df_b.rename(columns={b_id: "__ID__", b_val: "__B__"})
                
                # 统一为字符串类型
                df_a["__ID__"] = df_a["__ID__"].astype(str).str.strip()
                df_b["__ID__"] = df_b["__ID__"].astype(str).str.strip()
                df_b["__B__"] = df_b["__B__"].astype(str).fillna("")
                
                # 过滤空ID
                df_a_ids = set(df_a[df_a["__ID__"] != ""]["__ID__"].tolist())
                df_b_clean = df_b[df_b["__ID__"] != ""]
                
                # 找出B中有但A中没有的ID
                for _, row in df_b_clean.iterrows():
                    id_val = row["__ID__"]
                    b_val = row["__B__"]
                    
                    if id_val not in df_a_ids:
                        missing_in_a_details.append({
                            "Sheet": sheet,
                            "ID": id_val,
                            "表B的值": b_val,
                        })
                        
            except Exception as e:
                # 记录错误但继续处理其他Sheet
                self.log(f"处理Sheet '{sheet}' 时出错：{str(e)}")
                continue
        
        return missing_in_a_details

    def log_results_summary(self, rows: List[CompareRow]):
        """记录结果统计信息"""
        total = len(rows)
        same = sum(1 for r in rows if r.status == "一致")
        diff = sum(1 for r in rows if r.status == "不一致")
        miss = sum(1 for r in rows if r.status == "表B无此ID")
        miss_sheet = sum(1 for r in rows if r.status == "Sheet在B不存在")
        self.log(f"结果统计：总行数={total}  一致={same}  不一致={diff}  表B无此ID={miss}  Sheet在B不存在={miss_sheet}")

    def on_export(self, event):
        if not self.results:
            wx.MessageBox("当前无结果可导出。", "提示")
            return
        with wx.FileDialog(self, "导出为 Excel", wildcard="Excel 文件 (*.xlsx)|*.xlsx",
                           style=wx.FD_SAVE|wx.FD_OVERWRITE_PROMPT, defaultFile="output.xlsx") as dlg:
            if dlg.ShowModal() != wx.ID_OK:
                return
            out_path = dlg.GetPath()
        try:
            self.export_results_to_excel(out_path)
            self.log(f"导出完成：{out_path}")
            
            # 显示完成对话框
            result = wx.MessageBox("导出完成！", "提示", wx.OK)
            
            # 点击确定后自动打开文件夹
            if result == wx.OK:
                import subprocess
                import platform
                
                folder_path = os.path.dirname(out_path)
                try:
                    if platform.system() == "Windows":
                        # Windows: 打开文件夹并选中文件
                        subprocess.run(f'explorer /select,"{out_path}"', shell=True)
                    elif platform.system() == "Darwin":  # macOS
                        subprocess.run(["open", "-R", out_path])
                    else:  # Linux
                        subprocess.run(["xdg-open", folder_path])
                except Exception as e:
                    self.log(f"打开文件夹失败：{e}")
                    
        except Exception as e:
            self.log("导出失败：" + str(e))
            self.log(traceback.format_exc())

    def export_results_to_excel(self, out_path: str):
        # 计算基础统计信息
        sheets_a_count = len(self.sheets_a) if self.sheets_a else 0
        sheets_b_count = len(self.sheets_b) if self.sheets_b else 0
        common_sheets_count = len(self.common_sheets) if self.common_sheets else 0
        
        # 获取实际对比的Sheet列表
        compared_sheets = []
        if self.results:
            compared_sheets = list(set(r.sheet for r in self.results if r.sheet))
        compared_sheets_count = len(compared_sheets)
        
        # 先创建差异详情数据
        inconsistent_details = []  # 值不一致的记录
        missing_in_b_details = []  # A中有B中没有的记录
        
        for r in self.results:
            if r.status == "不一致":
                inconsistent_details.append({
                    "Sheet": r.sheet,
                    "ID": r.id_value,
                    "表A的值": r.a_value,
                    "表B的值": r.b_value,
                })
            elif r.status == "表B无此ID":
                missing_in_b_details.append({
                    "Sheet": r.sheet,
                    "ID": r.id_value,
                    "表A的值": r.a_value,
                })
        
        # 计算B中有A中没有的ID（需要反向对比）
        missing_in_a_details = self._find_missing_in_a(compared_sheets)

        # 按Sheet统计详细信息
        sheet_details = {}
        for sheet in compared_sheets:
            sheet_results = [r for r in self.results if r.sheet == sheet]
            
            # 统计各种状态的数量
            total_ids = len([r for r in sheet_results if r.status != "快速统计"])
            matched_ids = len([r for r in sheet_results if r.status in ["一致", "不一致"]])
            unmatched_a_ids = len([r for r in sheet_results if r.status == "表B无此ID"])
            consistent_ids = len([r for r in sheet_results if r.status == "一致"])
            inconsistent_ids = len([r for r in sheet_results if r.status == "不一致"])
            
            # 计算B中有A中没有的数量
            unmatched_b_ids = len([r for r in missing_in_a_details if r["Sheet"] == sheet])
            
            # 如果是快速统计模式，从统计信息中提取数据
            fast_stats = [r for r in sheet_results if r.status == "快速统计"]
            if fast_stats:
                # 解析快速统计信息
                stat_info = fast_stats[0].id_value
                if "总数" in stat_info:
                    import re
                    numbers = re.findall(r'总数(\d+), 一致(\d+), 不一致(\d+), 缺失(\d+)', stat_info)
                    if numbers:
                        total_ids, consistent_ids, inconsistent_ids, unmatched_a_ids = map(int, numbers[0])
                        matched_ids = consistent_ids + inconsistent_ids
            
            sheet_details[sheet] = {
                "总ID数": total_ids,
                "匹配上的ID数": matched_ids,
                "A中有B中没有的ID数": unmatched_a_ids,
                "B中有A中没有的ID数": unmatched_b_ids,
                "值一致的ID数": consistent_ids,
                "值不一致的ID数": inconsistent_ids,
            }
        
        # 尝试使用不同的Excel引擎
        excel_engine = self._get_available_excel_engine()
        
        if excel_engine == "xlsxwriter":
            self._export_new_format_xlsxwriter(out_path, sheets_a_count, sheets_b_count, 
                                             common_sheets_count, compared_sheets_count, 
                                             compared_sheets, sheet_details, inconsistent_details,
                                             missing_in_b_details, missing_in_a_details)
        else:
            self._export_new_format_openpyxl(out_path, sheets_a_count, sheets_b_count, 
                                           common_sheets_count, compared_sheets_count, 
                                           compared_sheets, sheet_details, inconsistent_details,
                                           missing_in_b_details, missing_in_a_details)

    def _get_available_excel_engine(self) -> str:
        """检测可用的Excel引擎"""
        try:
            import xlsxwriter
            return "xlsxwriter"
        except ImportError:
            try:
                import openpyxl
                return "openpyxl"
            except ImportError:
                return "basic"

    def _export_new_format_xlsxwriter(self, out_path: str, sheets_a_count: int, sheets_b_count: int, 
                                     common_sheets_count: int, compared_sheets_count: int, 
                                     compared_sheets: list, sheet_details: dict, inconsistent_details: list,
                                     missing_in_b_details: list, missing_in_a_details: list):
        """使用xlsxwriter引擎导出新格式报告"""
        with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
            wb = writer.book
            
            # 第一页：总览和详细统计
            ws_summary = wb.add_worksheet("对比总览")
            
            # 格式定义
            title_format = wb.add_format({"bold": True, "font_size": 16, "bg_color": "#4472C4", "font_color": "white"})
            header_format = wb.add_format({"bold": True, "bg_color": "#D9E2F3", "border": 1})
            data_format = wb.add_format({"border": 1})
            
            row = 0
            
            # 基础统计信息
            ws_summary.write(row, 0, "Excel对比报告 - 基础统计", title_format)
            ws_summary.merge_range(row, 0, row, 3, "Excel对比报告 - 基础统计", title_format)
            row += 2
            
            basic_stats = [
                ["表A Sheet总数", sheets_a_count],
                ["表B Sheet总数", sheets_b_count],
                ["共有Sheet数量", common_sheets_count],
                ["实际对比Sheet数量", compared_sheets_count],
            ]
            
            ws_summary.write(row, 0, "项目", header_format)
            ws_summary.write(row, 1, "数量", header_format)
            row += 1
            
            for stat_name, stat_value in basic_stats:
                ws_summary.write(row, 0, stat_name, data_format)
                ws_summary.write(row, 1, stat_value, data_format)
                row += 1
            
            row += 2
            
            # 匹配Sheet列表和详细统计
            ws_summary.write(row, 0, "各Sheet详细统计", title_format)
            ws_summary.merge_range(row, 0, row, 6, "各Sheet详细统计", title_format)
            row += 2
            
            # 表头
            headers = ["Sheet名称", "总ID数", "匹配上的ID数", "A中有B中没有", "B中有A中没有", "值一致的ID数", "值不一致的ID数"]
            for col, header in enumerate(headers):
                ws_summary.write(row, col, header, header_format)
            row += 1
            
            # 数据
            for sheet in compared_sheets:
                details = sheet_details.get(sheet, {})
                ws_summary.write(row, 0, sheet, data_format)
                ws_summary.write(row, 1, details.get("总ID数", 0), data_format)
                ws_summary.write(row, 2, details.get("匹配上的ID数", 0), data_format)
                ws_summary.write(row, 3, details.get("A中有B中没有的ID数", 0), data_format)
                ws_summary.write(row, 4, details.get("B中有A中没有的ID数", 0), data_format)
                ws_summary.write(row, 5, details.get("值一致的ID数", 0), data_format)
                ws_summary.write(row, 6, details.get("值不一致的ID数", 0), data_format)
                row += 1
            
            # 设置列宽
            ws_summary.set_column(0, 0, 20)  # Sheet名称
            ws_summary.set_column(1, 6, 15)  # 数值列
            
            # 设置自动筛选
            if compared_sheets:
                ws_summary.autofilter(row - len(compared_sheets) - 1, 0, row - 1, 6)
            
            # 第二页：不一致详情
            if inconsistent_details:
                ws_inconsistent = wb.add_worksheet("不一致详情")
                
                # 标题
                ws_inconsistent.write(0, 0, f"不一致数据详情（共 {len(inconsistent_details)} 条）", title_format)
                ws_inconsistent.merge_range(0, 0, 0, 3, f"不一致数据详情（共 {len(inconsistent_details)} 条）", title_format)
                
                # 表头
                detail_headers = ["Sheet", "ID", "表A的值", "表B的值"]
                for col, header in enumerate(detail_headers):
                    ws_inconsistent.write(2, col, header, header_format)
                
                # 数据
                for row_idx, detail in enumerate(inconsistent_details):
                    ws_inconsistent.write(3 + row_idx, 0, detail["Sheet"], data_format)
                    ws_inconsistent.write(3 + row_idx, 1, detail["ID"], data_format)
                    ws_inconsistent.write(3 + row_idx, 2, detail["表A的值"], data_format)
                    ws_inconsistent.write(3 + row_idx, 3, detail["表B的值"], data_format)
                
                # 设置列宽
                ws_inconsistent.set_column(0, 0, 20)  # Sheet
                ws_inconsistent.set_column(1, 1, 15)  # ID
                ws_inconsistent.set_column(2, 3, 30)  # 值列
                
                # 设置自动筛选
                ws_inconsistent.autofilter(2, 0, 2 + len(inconsistent_details), 3)
                
                # 冻结窗格
                ws_inconsistent.freeze_panes(3, 0)
            
            # 第三页：A中有B中没有的详情
            if missing_in_b_details:
                ws_missing_b = wb.add_worksheet("A有B无详情")
                
                # 标题
                ws_missing_b.write(0, 0, f"A中有B中没有的数据（共 {len(missing_in_b_details)} 条）", title_format)
                ws_missing_b.merge_range(0, 0, 0, 2, f"A中有B中没有的数据（共 {len(missing_in_b_details)} 条）", title_format)
                
                # 表头
                detail_headers = ["Sheet", "ID", "表A的值"]
                for col, header in enumerate(detail_headers):
                    ws_missing_b.write(2, col, header, header_format)
                
                # 数据
                for row_idx, detail in enumerate(missing_in_b_details):
                    ws_missing_b.write(3 + row_idx, 0, detail["Sheet"], data_format)
                    ws_missing_b.write(3 + row_idx, 1, detail["ID"], data_format)
                    ws_missing_b.write(3 + row_idx, 2, detail["表A的值"], data_format)
                
                # 设置列宽
                ws_missing_b.set_column(0, 0, 20)  # Sheet
                ws_missing_b.set_column(1, 1, 15)  # ID
                ws_missing_b.set_column(2, 2, 30)  # 值列
                
                # 设置自动筛选
                ws_missing_b.autofilter(2, 0, 2 + len(missing_in_b_details), 2)
                
                # 冻结窗格
                ws_missing_b.freeze_panes(3, 0)
            
            # 第四页：B中有A中没有的详情
            if missing_in_a_details:
                ws_missing_a = wb.add_worksheet("B有A无详情")
                
                # 标题
                ws_missing_a.write(0, 0, f"B中有A中没有的数据（共 {len(missing_in_a_details)} 条）", title_format)
                ws_missing_a.merge_range(0, 0, 0, 2, f"B中有A中没有的数据（共 {len(missing_in_a_details)} 条）", title_format)
                
                # 表头
                detail_headers = ["Sheet", "ID", "表B的值"]
                for col, header in enumerate(detail_headers):
                    ws_missing_a.write(2, col, header, header_format)
                
                # 数据
                for row_idx, detail in enumerate(missing_in_a_details):
                    ws_missing_a.write(3 + row_idx, 0, detail["Sheet"], data_format)
                    ws_missing_a.write(3 + row_idx, 1, detail["ID"], data_format)
                    ws_missing_a.write(3 + row_idx, 2, detail["表B的值"], data_format)
                
                # 设置列宽
                ws_missing_a.set_column(0, 0, 20)  # Sheet
                ws_missing_a.set_column(1, 1, 15)  # ID
                ws_missing_a.set_column(2, 2, 30)  # 值列
                
                # 设置自动筛选
                ws_missing_a.autofilter(2, 0, 2 + len(missing_in_a_details), 2)
                
                # 冻结窗格
                ws_missing_a.freeze_panes(3, 0)

    def _export_new_format_openpyxl(self, out_path: str, sheets_a_count: int, sheets_b_count: int, 
                                   common_sheets_count: int, compared_sheets_count: int, 
                                   compared_sheets: list, sheet_details: dict, inconsistent_details: list,
                                   missing_in_b_details: list, missing_in_a_details: list):
        """使用openpyxl引擎导出新格式报告（基础格式）"""
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            # 第一页：总览数据
            summary_data = []
            
            # 基础统计
            summary_data.extend([
                ["项目", "数量"],
                ["表A Sheet总数", sheets_a_count],
                ["表B Sheet总数", sheets_b_count],
                ["共有Sheet数量", common_sheets_count],
                ["实际对比Sheet数量", compared_sheets_count],
                ["", ""],  # 空行
                ["Sheet名称", "总ID数", "匹配上的ID数", "A中有B中没有", "B中有A中没有", "值一致的ID数", "值不一致的ID数"],
            ])
            
            # 各Sheet详细统计
            for sheet in compared_sheets:
                details = sheet_details.get(sheet, {})
                summary_data.append([
                    sheet,
                    details.get("总ID数", 0),
                    details.get("匹配上的ID数", 0),
                    details.get("A中有B中没有的ID数", 0),
                    details.get("B中有A中没有的ID数", 0),
                    details.get("值一致的ID数", 0),
                    details.get("值不一致的ID数", 0),
                ])
            
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name="对比总览", index=False, header=False)
            
            # 第二页：不一致详情
            if inconsistent_details:
                df_inconsistent = pd.DataFrame(inconsistent_details)
                df_inconsistent.to_excel(writer, sheet_name="不一致详情", index=False)
            
            # 第三页：A中有B中没有的详情
            if missing_in_b_details:
                df_missing_b = pd.DataFrame(missing_in_b_details)
                df_missing_b.to_excel(writer, sheet_name="A有B无详情", index=False)
            
            # 第四页：B中有A中没有的详情
            if missing_in_a_details:
                df_missing_a = pd.DataFrame(missing_in_a_details)
                df_missing_a.to_excel(writer, sheet_name="B有A无详情", index=False)


# ----------------------------- 应用入口 ----------------------------- #

class ExcelDiffApp(wx.App):
    def OnInit(self):
        self.SetAppName("ExcelDiffTool")
        frame = ExcelDiffFrame()
        frame.Show()
        return True


def main():
    app = ExcelDiffApp(False)
    app.MainLoop()


if __name__ == "__main__":
    main()
