# -*- coding: utf-8 -*-
"""桌面布局还原工具
功能：
  1. 记录当前桌面图标的位置（图标名 -> 坐标），保存为 JSON 布局文件。
  2. 还原已记录的布局：
     - 记录中存在、桌面也存在的图标 -> 移回记录的位置；
     - 记录中有、桌面已找不到（被删除/移动）-> 直接留空；
     - 桌面上有、记录中没有的新增图标 -> 统一移入桌面上新建的
       「新增图标_时间」文件夹（系统虚拟图标如“此电脑/回收站”无法移动，保持原位）。

原理：通过 COM 接口 IFolderView::SelectAndPositionItems 设置图标位置，
      这是微软官方支持的做法。直接向桌面 SysListView32 发送 LVM_SETITEMPOSITION
      的旧方法自 Windows 10 1809 起会被系统图标管理机制忽略，故不采用。
      （读取位置用 IFolderView::GetItemPosition，同样走官方接口。）
"""
import os
import sys
import json
import time
import shutil
import re
import ctypes
import ctypes.wintypes as wt
from ctypes import (POINTER, Structure, Union, byref, c_void_p, c_int, c_uint,
                    c_long, c_ulong, c_ushort, c_wchar_p, c_char, c_size_t, c_ssize_t)

import tkinter as tk
from tkinter import ttk, messagebox

if os.name != "nt":
    raise RuntimeError("本工具仅支持 Windows。")

import comtypes
from comtypes import GUID, COMMETHOD, HRESULT, IUnknown
import comtypes.client
from comtypes.automation import VARIANT, IDispatch

user32 = ctypes.windll.user32
shell32 = ctypes.windll.shell32
shlwapi = ctypes.windll.shlwapi
ole32 = ctypes.windll.ole32

# ============================ 常量与 GUID ============================

CLSID_ShellWindows = "{9BA05972-F6A8-11CF-A442-00A0C90A8F39}"
IID_IShellBrowser = GUID("{000214E2-0000-0000-C000-000000000046}")
IID_IShellView = GUID("{000214E3-0000-0000-C000-000000000046}")
IID_IShellFolder = GUID("{000214E6-0000-0000-C000-000000000046}")
SID_STopLevelBrowser = GUID("{4C96BE40-915C-11CF-99D3-00AA004AE837}")

SVGIO_ALLVIEW = 0x2          # IFolderView::ItemCount / Items
SVSI_POSITIONITEM = 0x80     # IFolderView::SelectAndPositionItems
SHGDN_INFOLDER = 0x1         # 显示名（与桌面看到的名称一致）
SWC_DESKTOP = 0x8            # FindWindowSW 的桌面窗口类
SWFO_NEEDDISPATCH = 0x1      # FindWindowSW 要求返回 IDispatch

LVS_AUTOARRANGE = 0x1
GWL_STYLE = -16
SPI_GETWORKAREA = 0x0030


class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

# ============================ COM 接口定义 ============================
# 均为扁平虚表声明（与 MIDL 生成的 vtable 顺序一致）。

class IShellWindows(IDispatch):
    """虚表顺序取自 Windows SDK exdisp.h（10.0.19041.0）。
    注意：本机（Win11 24H2）桌面未注册进 ShellWindows 集合（Count=0），
    必须用 FindWindowSW(SWC_DESKTOP) 这条专用路径获取桌面的 IDispatch。"""
    _iid_ = GUID("{85CB6900-4D95-11CF-960C-0080C7F4EE85}")
    _methods_ = [
        COMMETHOD([], HRESULT, "get_Count", (["out"], POINTER(c_long), "Count")),
        COMMETHOD([], HRESULT, "Item", (["in"], VARIANT, "index"),
                  (["out"], POINTER(POINTER(IDispatch)), "Folder")),
        COMMETHOD([], HRESULT, "_NewEnum", (["out"], POINTER(POINTER(IUnknown)), "ppunk")),
        COMMETHOD([], HRESULT, "Register", (["in"], POINTER(IDispatch), "pid"),
                  (["in"], c_long, "hwnd"), (["in"], c_int, "swClass"),
                  (["out"], POINTER(c_long), "plCookie")),
        COMMETHOD([], HRESULT, "RegisterPending", (["in"], c_long, "lThreadId"),
                  (["in"], POINTER(VARIANT), "pvarloc"),
                  (["in"], POINTER(VARIANT), "pvarlocRoot"),
                  (["in"], c_int, "swClass"), (["out"], POINTER(c_long), "plCookie")),
        COMMETHOD([], HRESULT, "Revoke", (["in"], c_long, "lCookie")),
        COMMETHOD([], HRESULT, "OnNavigate", (["in"], c_long, "lCookie"),
                  (["in"], POINTER(VARIANT), "pvarLoc")),
        COMMETHOD([], HRESULT, "OnActivated", (["in"], c_long, "lCookie"),
                  (["in"], VARIANT, "fActive")),
        COMMETHOD([], HRESULT, "FindWindowSW",
                  (["in"], POINTER(VARIANT), "pvarLoc"),
                  (["in"], POINTER(VARIANT), "pvarLocRoot"),
                  (["in"], c_int, "swClass"),
                  (["out"], POINTER(c_long), "phwnd"),
                  (["in"], c_int, "swfwOptions"),
                  (["out"], POINTER(POINTER(IDispatch)), "ppdispOut")),
    ]


class STRRET_UNION(Union):
    _fields_ = [("pOleStr", c_void_p), ("uOffset", c_ulong), ("cStr", c_char * 260)]


class STRRET(Structure):
    _anonymous_ = ("u",)
    _fields_ = [("uType", c_uint), ("u", STRRET_UNION)]


class IServiceProvider(IUnknown):
    _iid_ = GUID("{6D5140C1-7436-11CE-8034-00AA006009FA}")
    _methods_ = [
        COMMETHOD([], HRESULT, "QueryService",
                  (["in"], POINTER(GUID), "guidService"),
                  (["in"], POINTER(GUID), "riid"),
                  (["out"], POINTER(POINTER(IUnknown)), "ppvObj")),
    ]


class IShellView(IUnknown):
    """只需 GetWindow（用于校验视图窗口句柄）。"""
    _iid_ = IID_IShellView
    _methods_ = [
        COMMETHOD([], HRESULT, "GetWindow",
                  (["out"], POINTER(c_void_p), "phwnd")),
        COMMETHOD([], HRESULT, "ContextSensitiveHelp",
                  (["in"], c_int, "fEnterMode")),
    ]


class IShellBrowser(IUnknown):
    _iid_ = IID_IShellBrowser
    _methods_ = [
        COMMETHOD([], HRESULT, "GetWindow", (["out"], POINTER(c_void_p), "phwnd")),
        COMMETHOD([], HRESULT, "ContextSensitiveHelp", (["in"], c_int, "fEnterMode")),
        COMMETHOD([], HRESULT, "InsertMenusSB", (["in"], c_void_p, "hmenuShared"),
                  (["in"], c_void_p, "lpMenuWidths")),
        COMMETHOD([], HRESULT, "SetMenuSB", (["in"], c_void_p, "hmenuShared"),
                  (["in"], c_void_p, "hholemenu"), (["in"], c_void_p, "hwndActiveObject")),
        COMMETHOD([], HRESULT, "RemoveMenusSB", (["in"], c_void_p, "hmenuShared")),
        COMMETHOD([], HRESULT, "SetStatusTextSB", (["in"], c_wchar_p, "pszStatusText")),
        COMMETHOD([], HRESULT, "EnableModelessSB", (["in"], c_int, "fEnable")),
        COMMETHOD([], HRESULT, "TranslateAcceleratorSB", (["in"], c_void_p, "pmsg"),
                  (["in"], c_ushort, "wID")),
        COMMETHOD([], HRESULT, "BrowseObject", (["in"], c_void_p, "pidl"),
                  (["in"], c_uint, "wFlags")),
        COMMETHOD([], HRESULT, "GetViewStateStream", (["in"], c_uint, "grfMode"),
                  (["out"], POINTER(c_void_p), "ppStrm")),
        COMMETHOD([], HRESULT, "GetControlWindow", (["in"], c_uint, "id"),
                  (["out"], POINTER(c_void_p), "phwnd")),
        COMMETHOD([], HRESULT, "SendControlMsg", (["in"], c_uint, "id"),
                  (["in"], c_uint, "uMsg"), (["in"], c_size_t, "wParam"),
                  (["in"], c_ssize_t, "lParam"), (["out"], POINTER(c_ssize_t), "pret")),
        COMMETHOD([], HRESULT, "QueryActiveShellView",
                  (["out"], POINTER(POINTER(IShellView)), "ppshv")),
    ]


class IShellFolder(IUnknown):
    _iid_ = IID_IShellFolder
    _methods_ = [
        COMMETHOD([], HRESULT, "ParseDisplayName", (["in"], c_void_p, "hwndOwner"),
                  (["in"], c_void_p, "pbcReserved"), (["in"], c_wchar_p, "pszDisplayName"),
                  (["out"], POINTER(c_ulong), "pchEaten"),
                  (["out"], POINTER(c_void_p), "ppidl"),
                  (["out"], POINTER(c_ulong), "pdwAttributes")),
        COMMETHOD([], HRESULT, "EnumObjects", (["in"], c_void_p, "hwndOwner"),
                  (["in"], c_uint, "grfFlags"), (["out"], POINTER(c_void_p), "ppenumIDList")),
        COMMETHOD([], HRESULT, "BindToObject", (["in"], c_void_p, "pidl"),
                  (["in"], c_void_p, "pbcReserved"), (["in"], POINTER(GUID), "riid"),
                  (["out"], POINTER(c_void_p), "ppv")),
        COMMETHOD([], HRESULT, "BindToStorage", (["in"], c_void_p, "pidl"),
                  (["in"], c_void_p, "pbcReserved"), (["in"], POINTER(GUID), "riid"),
                  (["out"], POINTER(c_void_p), "ppv")),
        COMMETHOD([], HRESULT, "CompareIDs", (["in"], c_ssize_t, "lParam"),
                  (["in"], c_void_p, "pidl1"), (["in"], c_void_p, "pidl2")),
        COMMETHOD([], HRESULT, "CreateViewObject", (["in"], c_void_p, "hwndOwner"),
                  (["in"], POINTER(GUID), "riid"), (["out"], POINTER(c_void_p), "ppv")),
        COMMETHOD([], HRESULT, "GetAttributesOf", (["in"], c_uint, "cidl"),
                  (["in"], POINTER(c_void_p), "apidl"),
                  (["in", "out"], POINTER(c_ulong), "rgfInOut")),
        COMMETHOD([], HRESULT, "GetUIObjectOf", (["in"], c_void_p, "hwndOwner"),
                  (["in"], c_uint, "cidl"), (["in"], POINTER(c_void_p), "apidl"),
                  (["in"], POINTER(GUID), "riid"), (["out"], POINTER(c_ulong), "rgfReserved"),
                  (["out"], POINTER(c_void_p), "ppv")),
        COMMETHOD([], HRESULT, "GetDisplayNameOf", (["in"], c_void_p, "pidl"),
                  (["in"], c_uint, "uFlags"), (["out"], POINTER(STRRET), "pName")),
    ]


class IFolderView(IUnknown):
    # IID 与虚表顺序取自 Windows SDK shobjidl_core.h（10.0.19041.0）
    _iid_ = GUID("{CDE725B0-CCC9-4519-917E-325D72FAB4CE}")
    _methods_ = [
        COMMETHOD([], HRESULT, "GetCurrentViewMode", (["out"], POINTER(c_uint), "pViewMode")),
        COMMETHOD([], HRESULT, "SetCurrentViewMode", (["in"], c_uint, "ViewMode")),
        COMMETHOD([], HRESULT, "GetFolder", (["in"], POINTER(GUID), "riid"),
                  (["out"], POINTER(POINTER(IShellFolder)), "ppv")),
        COMMETHOD([], HRESULT, "Item", (["in"], c_int, "iItemIndex"),
                  (["out"], POINTER(c_void_p), "ppidl")),
        COMMETHOD([], HRESULT, "ItemCount", (["in"], c_uint, "uFlags"),
                  (["out"], POINTER(c_int), "pcItems")),
        COMMETHOD([], HRESULT, "Items", (["in"], c_uint, "uFlags"),
                  (["in"], POINTER(GUID), "riid"), (["out"], POINTER(c_void_p), "ppv")),
        COMMETHOD([], HRESULT, "GetSelectionMarkedItem", (["out"], POINTER(c_int), "piItem")),
        COMMETHOD([], HRESULT, "GetFocusedItem", (["out"], POINTER(c_int), "piItem")),
        COMMETHOD([], HRESULT, "GetItemPosition", (["in"], c_void_p, "pidl"),
                  (["out"], POINTER(wt.POINT), "ppt")),
        COMMETHOD([], HRESULT, "GetSpacing", (["in", "out"], POINTER(wt.POINT), "ppt")),
        COMMETHOD([], HRESULT, "GetDefaultSpacing", (["out"], POINTER(wt.POINT), "ppt")),
        COMMETHOD([], HRESULT, "GetAutoArrange"),
        COMMETHOD([], HRESULT, "SelectItem", (["in"], c_int, "iItem"),
                  (["in"], c_uint, "dwFlags")),
        COMMETHOD([], HRESULT, "SelectAndPositionItems", (["in"], c_uint, "cidl"),
                  (["in"], POINTER(c_void_p), "apidl"),
                  (["in"], POINTER(wt.POINT), "apt"), (["in"], c_uint, "dwFlags")),
    ]


shlwapi.StrRetToBufW.argtypes = [POINTER(STRRET), c_void_p, c_wchar_p, c_uint]
ole32.CoTaskMemFree.argtypes = [c_void_p]

# ============================ 桌面 API ============================

user32.FindWindowW.restype = c_void_p
user32.FindWindowExW.restype = c_void_p
user32.FindWindowExW.argtypes = [c_void_p, c_void_p, c_wchar_p, c_wchar_p]
user32.GetWindowLongPtrW.restype = ctypes.c_long
user32.GetWindowLongPtrW.argtypes = [c_void_p, c_int]


def get_desktop_listview():
    """找到桌面 SysListView32 视图窗口（Progman/WorkerW -> SHELLDLL_DefView -> SysListView32），
    仅用于 COM 兜底匹配与自动排列检测。"""
    def find_view(parent):
        defview = user32.FindWindowExW(parent, None, "SHELLDLL_DefView", None)
        if defview:
            return user32.FindWindowExW(defview, None, "SysListView32", None)
        return None

    progman = user32.FindWindowW("Progman", None)
    lv = find_view(progman) if progman else None
    if lv:
        return lv
    worker = None
    while True:
        worker = user32.FindWindowExW(None, worker, "WorkerW", None)
        if not worker:
            return None
        lv = find_view(worker)
        if lv:
            return lv


def _folder_view_from_dispatch(disp):
    """桌面 IDispatch -> IServiceProvider -> IShellBrowser -> IFolderView。"""
    svc = disp.QueryInterface(IServiceProvider)
    punk = svc.QueryService(SID_STopLevelBrowser, IID_IShellBrowser)
    browser = punk.QueryInterface(IShellBrowser)
    view = browser.QueryActiveShellView()
    return view.QueryInterface(IFolderView)


def get_desktop_folder_view():
    """获取桌面的 IFolderView 接口。"""
    sw = comtypes.client.CreateObject(CLSID_ShellWindows, interface=IShellWindows)
    loc = VARIANT(0)        # CSIDL_DESKTOP
    empty = VARIANT()       # VT_EMPTY
    hwnd, disp = sw.FindWindowSW(byref(loc), byref(empty), SWC_DESKTOP,
                                 SWFO_NEEDDISPATCH)
    if not disp:
        raise RuntimeError("无法连接到桌面视图（FindWindowSW 未返回桌面窗口）。")
    return _folder_view_from_dispatch(disp)


def _display_name(sf, pidl):
    strret = sf.GetDisplayNameOf(pidl, SHGDN_INFOLDER)
    buf = ctypes.create_unicode_buffer(512)
    shlwapi.StrRetToBufW(byref(strret), c_void_p(pidl), buf, 512)
    return buf.value


def _enumerate_with_pidl():
    """枚举桌面项：[(name, pidl, x, y)]。调用方负责用 CoTaskMemFree 释放 pidl。
    逐项容错：视图缓存异常（如已删除文件的僵尸条目导致总数虚高）时跳过坏项。"""
    fv = get_desktop_folder_view()
    sf = fv.GetFolder(IID_IShellFolder)
    count = fv.ItemCount(SVGIO_ALLVIEW)
    out = []
    consecutive_errors = 0
    for i in range(count):
        try:
            pidl = fv.Item(i)
            if not pidl:
                continue
            pt = fv.GetItemPosition(pidl)
            try:
                name = _display_name(sf, pidl)
            except Exception:
                name = "未知图标%d" % i
            consecutive_errors = 0
            out.append((name, pidl, pt.x, pt.y))
        except Exception:
            consecutive_errors += 1
            if consecutive_errors >= 5:
                break           # 视图状态已坏，用已取到的项继续
    return out


def get_desktop_items():
    """返回桌面每个图标的 {name, x, y}（COM 官方接口）。"""
    live = _enumerate_with_pidl()
    items = [{"name": n, "x": x, "y": y} for n, _p, x, y in live]
    for _n, pidl, _x, _y in live:
        try:
            ole32.CoTaskMemFree(c_void_p(pidl))
        except Exception:
            pass
    return items


def is_auto_arrange_on():
    lv = get_desktop_listview()
    if not lv:
        return False
    return bool(user32.GetWindowLongPtrW(lv, GWL_STYLE) & LVS_AUTOARRANGE)


def get_monitor_fingerprint():
    """当前显示器布局指纹：虚拟屏范围 + 每个显示器的矩形。"""
    rects = []

    def cb(_h, _hdc, lprect, _l):
        r = RECT()
        ctypes.memmove(ctypes.byref(r), lprect, ctypes.sizeof(RECT))
        rects.append((r.left, r.top, r.right, r.bottom))
        return 1

    proto = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p,
                               ctypes.POINTER(RECT), ctypes.c_int)
    user32.EnumDisplayMonitors(None, None, proto(cb), 0)
    virt = (user32.GetSystemMetrics(76), user32.GetSystemMetrics(77),
            user32.GetSystemMetrics(76) + user32.GetSystemMetrics(78),
            user32.GetSystemMetrics(77) + user32.GetSystemMetrics(79))
    return {"virtual": virt, "monitors": sorted(rects)}


def map_targets_adaptive(items, saved_fp, cur_fp):
    """显示器布局变化后，把存档坐标按比例映射到当前屏幕布局。
    显示器数量一致时按显示器一一对应映射；数量变化时按整个虚拟屏比例映射。"""
    om, cm = saved_fp.get("monitors") or [], cur_fp["monitors"]
    out = []
    if len(om) == len(cm) and om:
        for it in items:
            x, y = it["x"], it["y"]
            hit = None
            for rect in om:
                if rect[0] <= x < rect[2] and rect[1] <= y < rect[3]:
                    hit = rect
                    break
            if hit is None:
                hit = om[0]
            tgt = cm[om.index(hit)]
            ow, oh = max(1, hit[2] - hit[0]), max(1, hit[3] - hit[1])
            nw, nh = max(1, tgt[2] - tgt[0]), max(1, tgt[3] - tgt[1])
            nx = tgt[0] + int((x - hit[0]) / ow * nw)
            ny = tgt[1] + int((y - hit[1]) / oh * nh)
            out.append((it["name"], nx, ny))
        return out
    # 显示器数量变化（或旧存档无指纹）：把存档布局整体按比例映射到
    # 主屏工作区，保证每个图标都落在可见区域内
    rect = RECT()
    if not user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, byref(rect), 0):
        rect.left, rect.top, rect.right, rect.bottom = 0, 0, 1920, 1040
    xs = [it["x"] for it in items]
    ys = [it["y"] for it in items]
    bl, bt = min(xs), min(ys)
    bw = max(1, max(xs) + 115 - bl)
    bh = max(1, max(ys) + 147 - bt)
    pw = max(1, rect.right - rect.left)
    ph = max(1, rect.bottom - rect.top)
    for it in items:
        fx = (it["x"] - bl) / bw
        fy = (it["y"] - bt) / bh
        out.append((it["name"],
                    rect.left + int(fx * max(1, pw - 115)),
                    rect.top + int(fy * max(1, ph - 147))))
    return out


def get_icon_spacing(fv):
    """读取视图的图标格距（用于容差判断），失败则用保守默认值。"""
    try:
        pt = fv.GetSpacing(wt.POINT(0, 0))
        if pt and pt.x > 0 and pt.y > 0:
            return max(8, pt.x // 2), max(8, pt.y // 2)
    except Exception:
        pass
    return 60, 75


def _safe_position_batch(fv, pairs):
    """批量落位。若因某个失效 PIDL（如已删除文件的桌面缓存残留项）导致
    访问违例，二分定位坏项并跳过，返回失败的 (pidl, x, y) 列表。"""
    if not pairs:
        return []
    try:
        n = len(pairs)
        apidl = (c_void_p * n)(*[p for p, _x, _y in pairs])
        apt = (wt.POINT * n)(*[wt.POINT(x, y) for _p, x, y in pairs])
        fv.SelectAndPositionItems(n, apidl, apt, SVSI_POSITIONITEM)
        return []
    except OSError:
        if len(pairs) == 1:
            return list(pairs)
        mid = len(pairs) // 2
        return (_safe_position_batch(fv, pairs[:mid]) +
                _safe_position_batch(fv, pairs[mid:]))


def apply_positions(moves, max_rounds=10, park_others=False):
    """moves: [(图标显示名, x, y), ...]

    精确落位算法：
      1. 同名图标按「坐标排序后顺序配对」（稳定，不因同名互换）；
      2. park_others=True 时，先把不在还原名单里的杂散图标挪到屏幕右下角，
         腾出它们占据的格子（否则系统会把移动的图标挤到最近空位，造成不精确）；
      3. 每轮只提交位置不正确的图标，落位后重新枚举校验；
      4. 最多 max_rounds 轮修复 + 1 轮终检；连续 3 轮零进展则判定无法精确；
      5. 若两轮间错位集合不再变化（如被网格吸附），提前停止避免空转。

    返回 (匹配到的图标个数, 仍未精确落位的名称列表)。
    """
    if not moves:
        return 0, []
    pending = list(moves)
    matched = 0
    prev_need_key = None
    stall_count = 0
    parked = not park_others        # 是否已清理过杂散图标
    try:
        tolx, toly = get_icon_spacing(get_desktop_folder_view())
    except Exception:
        tolx, toly = 60, 75
    for round_i in range(max_rounds + 1):
        live = _enumerate_with_pidl()
        try:
            pool = {}
            for name, pidl, x, y in live:
                pool.setdefault(name, []).append((pidl, x, y))
            need = []
            outsiders = []          # 不在还原名单里的图标
            matched = 0
            for name in sorted(pool):
                entries = sorted(pool[name], key=lambda e: (e[1], e[2]))
                targets = sorted([(x, y) for nm, x, y in pending if nm == name])
                k = min(len(entries), len(targets))
                for i in range(k):
                    pidl, cx, cy = entries[i]
                    tx, ty = targets[i]
                    matched += 1
                    if abs(cx - tx) > tolx or abs(cy - ty) > toly:
                        need.append((name, pidl, tx, ty))
                outsiders.extend(entries[k:])
            if not parked and outsiders:
                # 只挪走「正压在目标格子上」的杂散图标，其余不动（避免大面积扰动）
                rect = RECT()
                occupied_targets = set((x, y) for _n, _p, x, y in need)
                blocking = [e for e in outsiders if (e[1], e[2]) in occupied_targets]
                if blocking and user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, byref(rect), 0):
                    parked = True
                    per_row = max(1, (rect.right - rect.left - 160) // 130)
                    for i, e in enumerate(blocking):
                        need.append(("(杂散图标)", e[1],
                                     rect.right - 80 - (i % per_row) * 130,
                                     rect.bottom - 80 - (i // per_row) * 160))
            if not need:
                return matched, []
            if round_i < max_rounds:
                fv = get_desktop_folder_view()
                bad_pairs = _safe_position_batch(
                    fv, [(p, x, y) for _n, p, x, y in need])
                if bad_items := set(bad_pairs):
                    # 无法安全落位的图标（多为已删除文件的缓存残留），跳过不计
                    skipped = [n for n, p, x, y in need if (p, x, y) in bad_items]
                    need = [(n, p, x, y) for n, p, x, y in need
                            if (p, x, y) not in bad_items]
                    if not need:
                        return matched, skipped
            else:
                # 终检轮：不再提交，报告剩余误差
                return matched, [n for n, _p, _x, _y in need]
        finally:
            for _n, pidl, _x, _y in live:
                try:
                    ole32.CoTaskMemFree(c_void_p(pidl))
                except Exception:
                    pass
        # Win11 24H2 的桌面视图在图标剧烈增删后会有一段"僵直期"，
        # 移动命令可能静默失效。若连续多轮毫无进展，放慢节奏等 shell 恢复。
        need_key = frozenset((x, y) for _n, _p, x, y in need)
        if need_key == prev_need_key:
            stall_count += 1
            if stall_count >= 3:
                # 连续 3 轮零进展，判定为不可精确落位（网格吸附/视图僵直）
                return matched, [n for n, _p, _x, _y in need]
            time.sleep(1.2)             # 多等一会让 shell 缓过来
        else:
            stall_count = 0
            time.sleep(0.35)
        prev_need_key = need_key
    return matched, []


def refresh_desktop():
    shell32.SHChangeNotify(0x7FFFFFFF, 0, None, None)  # SHCNE_ALLEVENTS


def get_desktop_dirs():
    """返回 [当前用户桌面, 公共桌面] 目录。"""
    dirs = []
    buf = ctypes.create_unicode_buffer(260)
    CSIDL_DESKTOPDIRECTORY = 0x00
    CSIDL_COMMON_DESKTOPDIRECTORY = 0x19
    if shell32.SHGetFolderPathW(None, CSIDL_DESKTOPDIRECTORY, None, 0, buf) == 0:
        dirs.append(buf.value)
    if shell32.SHGetFolderPathW(None, CSIDL_COMMON_DESKTOPDIRECTORY, None, 0, buf) == 0:
        dirs.append(buf.value)
    return [d for d in dirs if d and os.path.isdir(d)]


def resolve_desktop_path(name):
    """把桌面上显示的图标名映射为真实文件/文件夹路径（找不到返回 None，即虚拟图标）。"""
    for d in get_desktop_dirs():
        try:
            entries = os.listdir(d)
        except OSError:
            continue
        for e in entries:                      # 精确匹配
            if e.lower() == name.lower():
                return os.path.join(d, e)
        for e in entries:                      # 忽略扩展名匹配（桌面不显示常见扩展名）
            base = os.path.splitext(e)[0]
            if base and base.lower() == name.lower():
                return os.path.join(d, e)
    return None


def unique_target(path):
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    n = 2
    while os.path.exists("%s (%d)%s" % (base, n, ext)):
        n += 1
    return "%s (%d)%s" % (base, n, ext)


# ============================ 程序目录 ============================

if getattr(sys, "frozen", False):          # PyInstaller 打包后
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LAYOUTS_DIR = os.path.join(BASE_DIR, "Layouts")


# ============================ 界面 ============================

class App(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title("桌面布局还原工具")
        self.geometry("660x500")
        self.resizable(False, False)
        os.makedirs(LAYOUTS_DIR, exist_ok=True)
        self._build_ui()
        self.refresh_profiles()
        self.log("就绪。提示：还原前请先关闭桌面的「自动排列图标」。")
        if is_auto_arrange_on():
            self.log("警告：检测到桌面当前开启了「自动排列图标」，还原位置不会生效！")

    def _build_ui(self):
        f = ttk.Frame(self)
        f.pack(fill="both", expand=True, padx=12, pady=10)

        row1 = ttk.Frame(f); row1.pack(fill="x", pady=3)
        ttk.Label(row1, text="布局名称:").pack(side="left")
        self.txt_name = ttk.Entry(row1)
        self.txt_name.pack(side="left", padx=6, fill="x", expand=True)
        self.txt_name.insert(0, "布局_" + time.strftime("%Y%m%d_%H%M%S"))
        ttk.Button(row1, text="保存当前布局", command=self.on_save).pack(side="left")

        row2 = ttk.Frame(f); row2.pack(fill="x", pady=3)
        ttk.Label(row2, text="已存布局:").pack(side="left")
        self.cbo = ttk.Combobox(row2, state="readonly")
        self.cbo.pack(side="left", padx=6, fill="x", expand=True)
        ttk.Button(row2, text="还原所选布局", command=self.on_restore).pack(side="left")
        ttk.Button(row2, text="删除", width=6,
                   command=self.on_delete).pack(side="left", padx=(4, 0))

        row3 = ttk.Frame(f); row3.pack(fill="x", pady=3)
        ttk.Button(row3, text="打开布局文件夹",
                   command=lambda: os.startfile(LAYOUTS_DIR)).pack(side="left")
        ttk.Button(row3, text="预览当前桌面图标",
                   command=self.on_preview).pack(side="left", padx=(6, 0))

        tip = tk.Label(
            f, fg="#c0392b", justify="left", wraplength=620,
            text="提示：还原前请先在桌面右键 → 查看，取消「自动排列图标」与"
                 "「将图标与网格对齐」，否则系统会自动重排、坐标会被吸附。")
        tip.pack(fill="x", pady=(6, 4))

        self.txt_log = tk.Text(f, height=15, state="disabled",
                               font=("Microsoft YaHei UI", 9))
        self.txt_log.pack(fill="both", expand=True)

    # ---------- 通用 ----------
    def log(self, msg):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", time.strftime("[%H:%M:%S] ") + msg + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def refresh_profiles(self):
        try:
            names = sorted(
                (os.path.splitext(n)[0] for n in os.listdir(LAYOUTS_DIR)
                 if n.lower().endswith(".json")),
                reverse=True)
        except OSError:
            names = []
        self.cbo["values"] = names
        if names and self.cbo.get() not in names:
            self.cbo.set(names[0])
        elif not names:
            self.cbo.set("")

    # ---------- 功能 1：记录布局 ----------
    def on_save(self):
        name = self.txt_name.get().strip()
        if not name:
            messagebox.showinfo("提示", "请先填写布局名称。", parent=self)
            return
        name = re.sub(r'[\\/:*?"<>|]', "_", name)
        try:
            items = get_desktop_items()
        except Exception as e:
            messagebox.showerror("错误", str(e), parent=self)
            return
        if not items:
            messagebox.showwarning("警告", "未读取到任何桌面图标。", parent=self)
            return
        data = {
            "name": name,
            "saved_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "monitors": get_monitor_fingerprint(),
            "items": [{"name": it["name"], "x": it["x"], "y": it["y"]}
                      for it in items],
        }
        path = os.path.join(LAYOUTS_DIR, name + ".json")
        with open(path, "w", encoding="utf-8") as fp:
            json.dump(data, fp, ensure_ascii=False, indent=2)
        self.log("已保存 %d 个图标的位置 → Layouts\\%s.json" % (len(items), name))
        self.refresh_profiles()
        self.cbo.set(name)

    def on_preview(self):
        try:
            items = get_desktop_items()
        except Exception as e:
            messagebox.showerror("错误", str(e), parent=self)
            return
        self.log("当前桌面共 %d 个图标：" % len(items))
        for it in items:
            self.log("  %-32s (%d, %d)" % (it["name"], it["x"], it["y"]))

    # ---------- 功能 2：还原布局 ----------
    def on_restore(self):
        sel = self.cbo.get()
        if not sel:
            messagebox.showinfo("提示", "请先选择一个已保存的布局。", parent=self)
            return
        path = os.path.join(LAYOUTS_DIR, sel + ".json")
        try:
            with open(path, "r", encoding="utf-8") as fp:
                data = json.load(fp)
        except Exception as e:
            messagebox.showerror("错误", "读取布局文件失败：%s" % e, parent=self)
            return
        saved = data.get("items") or []
        if not saved:
            messagebox.showwarning("警告", "该布局文件为空。", parent=self)
            return
        try:
            live = get_desktop_items()
        except Exception as e:
            messagebox.showerror("错误", str(e), parent=self)
            return

        # 目标池按名字分组；同名图标按「坐标排序后顺序配对」（稳定不互换）
        # 显示器布局与存档时不一致时，改用自适应映射（按比例映射到新屏幕布局）
        cur_fp = get_monitor_fingerprint()
        saved_fp = data.get("monitors")
        adaptive = bool(saved_fp and saved_fp.get("monitors") != cur_fp["monitors"])
        if adaptive:
            self.log("检测到显示器布局变化（存档 %d 屏 → 当前 %d 屏），"
                     "启用自适应映射：按比例把布局映射到当前屏幕。"
                     % (len(saved_fp["monitors"]), len(cur_fp["monitors"])))
            mapped = map_targets_adaptive(saved, saved_fp, cur_fp)
            saved_map = {}
            for n, x, y in mapped:
                saved_map.setdefault(n, []).append({"name": n, "x": x, "y": y})
        else:
            saved_map = {}
            for s in saved:
                saved_map.setdefault(s["name"], []).append(s)

        live_by_name = {}
        for it in live:
            live_by_name.setdefault(it["name"], []).append(it)

        moves, consumed = [], set()
        for name, lst in saved_map.items():
            targets = sorted((s["x"], s["y"]) for s in lst)
            litems = sorted(live_by_name.get(name, []),
                            key=lambda it: (it["x"], it["y"]))
            for i in range(min(len(targets), len(litems))):
                tx, ty = targets[i]
                moves.append((name, tx, ty))
                consumed.add(id(litems[i]))
        new_live = [it for it in live if id(it) not in consumed]

        # 第一步：把仍存在的图标放回记录的位置（先清场，再多轮校验直到精确）
        restored, unresolved = 0, []
        if moves:
            try:
                restored, unresolved = apply_positions(moves, park_others=True)
                self.log("已还原 %d 个图标的位置。" % restored)

                # 大面积失败 -> 疑似显示器布局与存档时不一致：自适应映射重试
                if len(unresolved) > max(6, int(len(moves) * 0.3)):
                    self.log("超过 30%% 的图标未能落位，疑似显示器布局与存档时"
                             "不一致，自动改用自适应映射重试…")
                    if saved_fp is None:      # 旧版存档无显示器指纹
                        xs = [s["x"] for s in saved]
                        ys = [s["y"] for s in saved]
                        fp = {"virtual": (min(xs), min(ys),
                                          max(xs) + 115, max(ys) + 147),
                              "monitors": []}
                    else:
                        fp = saved_fp
                    mapped = map_targets_adaptive(saved, fp, cur_fp)
                    saved_map2 = {}
                    for n, x, y in mapped:
                        saved_map2.setdefault(n, []).append({"name": n,
                                                             "x": x, "y": y})
                    moves2, consumed2 = [], set()
                    for name, lst in saved_map2.items():
                        targets = sorted((s["x"], s["y"]) for s in lst)
                        litems = sorted(live_by_name.get(name, []),
                                        key=lambda it: (it["x"], it["y"]))
                        for i in range(min(len(targets), len(litems))):
                            tx, ty = targets[i]
                            moves2.append((name, tx, ty))
                            consumed2.add(id(litems[i]))
                    restored2, unresolved2 = apply_positions(moves2,
                                                             park_others=True)
                    self.log("自适应重试: 还原 %d 个, 未落位 %d 个"
                             % (restored2, len(unresolved2)))
                    if len(unresolved2) < len(unresolved):
                        restored, unresolved = restored2, unresolved2
                        self.log("已采用自适应映射结果。")

                # 僵直期补偿：等待 shell 落稳后补一次校验修复
                if unresolved:
                    time.sleep(1.5)
                    r3, u3 = apply_positions(moves, park_others=False)
                    if len(u3) < len(unresolved):
                        restored, unresolved = r3, u3
                        self.log("补校验后未落位减少为 %d 个。" % len(unresolved))

                if unresolved:
                    self.log("警告：%d 个图标未能精确落位（常见原因：开启了"
                             "「对齐网格」、桌面已满无空位，或显示器切换后视图"
                             "短暂僵直）：%s"
                             % (len(unresolved),
                                "、".join(unresolved[:8]) + ("…" if len(unresolved) > 8 else "")))
            except Exception as e:
                messagebox.showerror("错误", "设置图标位置失败：%s" % e, parent=self)
                return

        # 情形 1：记录中有、桌面已没有 -> 留空
        missing = []
        for name, lst in saved_map.items():
            k = min(len(lst), len(live_by_name.get(name, [])))
            missing.extend([name] * (len(lst) - k))
        if missing:
            self.log("记录中的 %d 个图标已不存在（按要求留空）：%s"
                     % (len(missing), "、".join(missing[:10]))
                     + ("…" if len(missing) > 10 else ""))

        # 情形 2：桌面新增的图标 -> 移入桌面的「新增图标_时间」文件夹
        moved_into, virtual_left, folder_created = 0, [], None
        desktop_dirs = get_desktop_dirs()
        dest = (os.path.join(desktop_dirs[0],
                             "新增图标_" + time.strftime("%Y%m%d_%H%M%S"))
                if desktop_dirs else None)
        for it in new_live:
            n = it["name"]
            if n.startswith("新增图标_"):
                self.log("跳过上次整理产生的文件夹：%s" % n)
                continue
            p = resolve_desktop_path(n)
            if not p:
                virtual_left.append(n)
                continue
            if dest is None:
                continue
            if not os.path.isdir(dest):
                os.makedirs(dest)
                folder_created = os.path.basename(dest)
            try:
                shutil.move(p, unique_target(os.path.join(dest, os.path.basename(p))))
                moved_into += 1
            except Exception as ex:
                virtual_left.append("%s（移动失败：%s）" % (n, ex))
        if moved_into:
            self.log("新增的 %d 个图标已移入桌面文件夹：%s" % (moved_into, folder_created))
        if virtual_left:
            self.log("以下图标是系统图标或无法识别/移动，保持原位：%s"
                     % "、".join(virtual_left))

        refresh_desktop()
        self.log("还原完成。若图标位置没有变化，请检查是否已关闭「自动排列图标」。")
        messagebox.showinfo(
            "还原完成",
            "位置还原：%d 个\n缺失留空：%d 个\n新增移入文件夹：%d 个\n"
            "系统图标未移动：%d 个\n未精确落位：%d 个"
            % (restored, len(missing), moved_into, len(virtual_left),
               len(unresolved)),
            parent=self)

    def on_delete(self):
        sel = self.cbo.get()
        if not sel:
            return
        if not messagebox.askyesno("确认", "确定删除布局「%s」吗？" % sel, parent=self):
            return
        try:
            os.remove(os.path.join(LAYOUTS_DIR, sel + ".json"))
            self.log("已删除布局：%s" % sel)
        except OSError as e:
            messagebox.showerror("错误", str(e), parent=self)
        self.refresh_profiles()


def selftest():
    """无界面自检：读取图标 -> 移动一格 -> 还原。结果写入 selftest.log。"""
    lines = []
    try:
        items = get_desktop_items()
        lines.append("读取桌面图标: %d 个" % len(items))
        if items:
            it = items[0]
            orig = (it["x"], it["y"])
            n, _u = apply_positions([(it["name"], it["x"] + 115, it["y"])])
            time.sleep(0.5)
            now = {i["name"]: (i["x"], i["y"]) for i in get_desktop_items()}
            moved = n > 0 and now.get(it["name"]) != orig
            apply_positions([(it["name"],) + orig])
            time.sleep(0.5)
            fin = {i["name"]: (i["x"], i["y"]) for i in get_desktop_items()}
            back = fin.get(it["name"]) == orig
            lines.append("移动测试: %s" % ("OK" if moved else "FAIL"))
            lines.append("复原测试: %s" % ("OK" if back else "FAIL"))
        lines.append("自动排列开启: %s" % ("是（需关闭）" if is_auto_arrange_on() else "否"))
    except Exception as e:
        lines.append("自检出错: %r" % e)
    log_path = os.path.join(BASE_DIR, "selftest.log")
    with open(log_path, "w", encoding="utf-8") as fp:
        fp.write("\n".join(lines) + "\n")
    print("\n".join(lines))


def main():
    if "--selftest" in sys.argv:
        selftest()
        return
    App().mainloop()


if __name__ == "__main__":
    main()
