# -*- coding: utf-8 -*-
"""
移动硬盘清理工具 - 管理员模式
支持 USB 硬件级安全弹出（停止转动）
自动检测并恢复脱机磁盘
"""
import ctypes
import sys
import os
import subprocess
import string
import time
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from ctypes import wintypes
import threading

# ── 版本号（修改此处即可更新界面右上角显示） ──
APP_VERSION = "2.3.3"

SERVICES = [
    ("WSearch",       "Windows Search"),
    ("SysMain",       "SysMain"),
    ("VSS",           "Volume Shadow Copy"),
    ("defragsvc",     "Optimize Drives"),
    ("WMPNetworkSvc", "WMP Network Sharing"),
    ("StorSvc",       "Storage Service"),
]

DRIVE_TYPE_MAP = {
    2: "可移动",
    3: "固定",
    4: "网络",
    5: "光驱",
    6: "RAM",
}

USB_BUS_TYPES = {"USB", "USB3"}

USB_EJECT_PS1 = r'''
param([int]$DiskNumber)
$ErrorActionPreference = 'Stop'
try {
    $disk = Get-WmiObject Win32_DiskDrive | Where-Object { $_.Index -eq $DiskNumber }
    if (-not $disk) {
        Write-Output "DISK_NOT_FOUND"
        exit 1
    }
    $devId = $disk.PNPDeviceID
    Write-Output "PnP: $devId"

    Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class CfgMgr32 {
    [DllImport("cfgmgr32.dll", CharSet=CharSet.Unicode)]
    public static extern int CM_Locate_DevNode(out int pdnDevInst, string pDeviceID, int ulFlags);
    [DllImport("cfgmgr32.dll")]
    public static extern int CM_Get_Parent(out int pdnDevInst, int dnDevInst, int ulFlags);
    [DllImport("cfgmgr32.dll", CharSet=CharSet.Unicode)]
    public static extern int CM_Request_Device_Eject(int dnDevInst, out int pVetoType, StringBuilder pszVetoName, int ulNameLength, int ulFlags);
    [DllImport("cfgmgr32.dll", CharSet=CharSet.Unicode)]
    public static extern int CM_Get_Device_ID(int dnDevInst, StringBuilder Buffer, int BufferLen, int ulFlags);
}
"@

    $devInst = 0
    $r = [CfgMgr32]::CM_Locate_DevNode([ref]$devInst, $devId, 0)
    if ($r -ne 0) {
        Write-Output "CM_Locate_DevNode failed: $r"
        exit 1
    }

    $parent = 0
    $r = [CfgMgr32]::CM_Get_Parent([ref]$parent, $devInst, 0)
    if ($r -ne 0) {
        Write-Output "CM_Get_Parent failed: $r"
        exit 1
    }

    $parentId = New-Object System.Text.StringBuilder 512
    [CfgMgr32]::CM_Get_Device_ID($parent, $parentId, 512, 0) | Out-Null
    Write-Output "Parent: $($parentId.ToString())"

    $vetoType = 0
    $vetoName = New-Object System.Text.StringBuilder 512
    $r = [CfgMgr32]::CM_Request_Device_Eject($parent, [ref]$vetoType, $vetoName, 512, 0)
    if ($r -eq 0) {
        Write-Output "USB_EJECT_OK"
    } else {
        $vn = $vetoName.ToString()
        Write-Output "USB_EJECT_FAIL: result=$r vetoType=$vetoType vetoName=$vn"
    }
} catch {
    Write-Output "ERROR: $_"
    exit 1
}
'''


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def run_cmd(cmd, timeout=120):
    try:
        p = subprocess.run(
            cmd, shell=True, capture_output=True, text=True,
            timeout=timeout, encoding="gbk", errors="replace"
        )
        return p.returncode, p.stdout, p.stderr
    except subprocess.TimeoutExpired:
        return -1, "", "命令执行超时"
    except Exception as e:
        return -1, "", str(e)


def svc_status(name):
    rc, out, _ = run_cmd(f"sc query {name}")
    if "RUNNING" in out:
        return "running"
    if "STOPPED" in out:
        return "stopped"
    return "missing"


def drive_exists(letter):
    letter = letter.rstrip(":\\").upper()
    idx = ord(letter) - ord('A')
    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    return bool(bitmask & (1 << idx))


def get_drive_type_code(letter):
    letter = letter.rstrip(":\\").upper()
    return ctypes.windll.kernel32.GetDriveTypeW(f"{letter}:\\")


# ── 快速获取盘符（仅 Win32 API，不调 PowerShell，瞬间完成） ──
def get_drives_fast(min_letter='G'):
    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    drives = []
    for i, ch in enumerate(string.ascii_uppercase):
        if ch < min_letter:
            continue
        if bitmask & (1 << i):
            dt = ctypes.windll.kernel32.GetDriveTypeW(f"{ch}:\\")
            drives.append((f"{ch}:", DRIVE_TYPE_MAP.get(dt, "未知")))
    return drives


# ── 批量获取所有盘符对应的物理磁盘总线类型（需要 PowerShell，较慢） ──
def get_drive_bus_types():
    cmd = (
        'powershell -Command "'
        "Get-Partition | Where-Object { $_.DriveLetter } | ForEach-Object { "
        "try { $dk = $_ | Get-Disk -ErrorAction Stop; "
        "Write-Output ('{0}|{1}' -f $_.DriveLetter, $dk.BusType) "
        "} catch {} "
        '}"'
    )
    rc, out, _ = run_cmd(cmd, timeout=15)
    result = {}
    if rc == 0 and out.strip():
        for line in out.strip().splitlines():
            parts = line.strip().split('|', 1)
            if len(parts) == 2 and parts[0]:
                result[parts[0].upper()] = parts[1].strip()
    return result


# ── 获取单个盘符的总线类型 ──
def get_drive_bus_type(letter):
    letter = letter.rstrip(":\\").upper()
    cmd = (
        f'powershell -Command "'
        f"try {{ $p = Get-Partition -DriveLetter {letter} -ErrorAction Stop; "
        f"($p | Get-Disk).BusType }} catch {{ Write-Output 'Unknown' }}"
        f'"'
    )
    rc, out, _ = run_cmd(cmd, timeout=10)
    if rc == 0 and out.strip():
        return out.strip()
    return "Unknown"


# ── 完整获取盘符（含 PowerShell 总线类型检测） ──
def get_drives_full(min_letter='G'):
    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    candidates = []
    for i, ch in enumerate(string.ascii_uppercase):
        if ch < min_letter:
            continue
        if bitmask & (1 << i):
            dt = ctypes.windll.kernel32.GetDriveTypeW(f"{ch}:\\")
            candidates.append((ch, dt))
    if not candidates:
        return [], {}

    bus_types = get_drive_bus_types()
    drives = []
    for ch, dt in candidates:
        bus = bus_types.get(ch, "").upper()
        if dt == 3 and bus in USB_BUS_TYPES:
            label = "USB硬盘"
        elif dt == 2 and bus in USB_BUS_TYPES:
            label = "可移动"
        else:
            label = DRIVE_TYPE_MAP.get(dt, "未知")
        drives.append((f"{ch}:", label))
    return drives, bus_types


def get_offline_disks():
    cmd = (
        'powershell -Command "'
        "Get-Disk | Where-Object { $_.OperationalStatus -eq 'Offline' } | "
        "Select-Object Number, FriendlyName, Size, BusType | "
        'ConvertTo-Csv -NoTypeInformation"'
    )
    rc, out, _ = run_cmd(cmd, timeout=15)
    disks = []
    if rc == 0 and out.strip():
        lines = out.strip().splitlines()
        for line in lines[1:]:
            parts = line.strip('"').split('","')
            if len(parts) >= 4:
                try:
                    num = int(parts[0])
                    name = parts[1]
                    size_bytes = int(parts[2])
                    bus = parts[3]
                    size_gb = size_bytes / (1024**3)
                    disks.append({
                        "number": num, "name": name,
                        "size_gb": size_gb, "bus": bus,
                    })
                except (ValueError, IndexError):
                    pass
    return disks


def set_disk_online(disk_number):
    cmd = (
        f'powershell -Command "'
        f'Set-Disk -Number {disk_number} -IsOffline $false"'
    )
    rc, _, _ = run_cmd(cmd, timeout=15)
    return rc == 0


def eject_volume_api(letter):
    letter = letter.rstrip(":\\")
    volume = f"\\\\.\\{letter}:"
    k32 = ctypes.windll.kernel32
    SHARE_RW = 0x1 | 0x2
    OPEN_EXISTING = 3
    FSCTL_LOCK = 0x00090018
    FSCTL_DISMOUNT = 0x00090020
    IOCTL_EJECT = 0x002D4808

    for access in [0xC0000000, 0x80000000, 0]:
        h = k32.CreateFileW(volume, access, SHARE_RW, None, OPEN_EXISTING, 0, None)
        if h != -1:
            break
    else:
        return False, "无法打开卷句柄"

    br = wintypes.DWORD(0)
    k32.DeviceIoControl(h, FSCTL_LOCK, None, 0, None, 0, ctypes.byref(br), None)
    k32.DeviceIoControl(h, FSCTL_DISMOUNT, None, 0, None, 0, ctypes.byref(br), None)
    ok = k32.DeviceIoControl(h, IOCTL_EJECT, None, 0, None, 0, ctypes.byref(br), None)
    k32.CloseHandle(h)
    return bool(ok), ("API 弹出指令已发送" if ok else "API IOCTL 失败")


# ── 推荐按钮配色 ──
REC_BG       = "#dae8fc"
REC_FG       = "#1a3a6b"
REC_ACTIVE   = "#b8d4f0"
STAR_COLOR   = "#c8a000"
STAR_HOVER   = "#ffe066"
REC_HOVER_BG = "#3b7dd8"
REC_HOVER_FG = "#ffffff"


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("移动硬盘清理工具 - 管理员模式")
        # ── 宽度 +60 → 660，高度保持 540 ──
        self.root.geometry("560x690")
        self.root.minsize(460, 530)
        self._busy = False
        self._svc_was_running = {}
        self._bus_cache = {}        # 总线类型缓存 {'G': 'USB', 'C': 'NVMe', ...}
        self._detecting = False     # 是否正在后台检测总线类型
        self.build_ui()
        # 先显示界面，再后台检测
        self.root.after(100, self._start_bus_detection)
        self.root.after(500, self._check_offline_on_start)
        self.root.mainloop()

    # ── 构建推荐按钮 ──
    def _make_rec_btn(self, parent, command):
        frm = tk.Frame(parent, bg=REC_BG, relief=tk.RAISED, bd=2, cursor="hand2")
        inner = tk.Frame(frm, bg=REC_BG)
        inner.pack(expand=True, pady=(4, 3))
        line1 = tk.Frame(inner, bg=REC_BG)
        line1.pack()
        lbl_star = tk.Label(
            line1, text="★", fg=STAR_COLOR, bg=REC_BG,
            font=("Segoe UI", 13), cursor="hand2",
        )
        lbl_star.pack(side="left")
        lbl_title = tk.Label(
            line1, text=" 安全弹出（推荐）", fg=REC_FG, bg=REC_BG,
            font=("Microsoft YaHei UI", 11, "bold"), cursor="hand2",
        )
        lbl_title.pack(side="left")
        lbl_sub = tk.Label(
            inner, text="停止服务 + 弹出硬盘", fg="#555", bg=REC_BG,
            font=("Microsoft YaHei UI", 10), cursor="hand2",
        )
        lbl_sub.pack()
        ws = [frm, inner, line1, lbl_star, lbl_title, lbl_sub]

        def _hover_enter(e):
            for w in ws:
                w.configure(bg=REC_HOVER_BG)
            lbl_star.config(fg=STAR_HOVER)
            lbl_title.config(fg=REC_HOVER_FG)
            lbl_sub.config(fg="#ccc")

        def _hover_leave(e):
            try:
                px = frm.winfo_pointerx() - frm.winfo_rootx()
                py = frm.winfo_pointery() - frm.winfo_rooty()
                if 0 <= px <= frm.winfo_width() and 0 <= py <= frm.winfo_height():
                    return
            except Exception:
                pass
            for w in ws:
                w.configure(bg=REC_BG)
            lbl_star.config(fg=STAR_COLOR)
            lbl_title.config(fg=REC_FG)
            lbl_sub.config(fg="#555")

        def _click(e):
            frm.config(relief=tk.SUNKEN)
            frm.after(80, lambda: frm.config(relief=tk.RAISED))
            frm.after(100, command)

        for w in ws:
            w.bind("<Button-1>", _click)
            w.bind("<Enter>", _hover_enter)
            w.bind("<Leave>", _hover_leave)
        return frm

    # ------------------------------------------------------------------ UI
    def build_ui(self):
        style = ttk.Style()
        style.theme_use("clam")
        UI_FONT = ("Microsoft YaHei UI", 11)
        UI_BOLD = ("Microsoft YaHei UI", 11, "bold")
        self.root.option_add("*Font", UI_FONT)
        style.configure("TButton",           font=UI_FONT)
        style.configure("TLabel",            font=UI_FONT)
        style.configure("TCheckbutton",      font=UI_FONT)
        style.configure("TLabelframe",       font=UI_BOLD)
        style.configure("TLabelframe.Label", font=UI_BOLD)
        style.configure("TNotebook.Tab",     font=UI_FONT, padding=(12, 4))
        style.configure("TCombobox",         font=UI_FONT)

        m = ttk.Frame(self.root, padding=8)
        m.pack(fill="both", expand=True)

        # ── 右上角版本号（浅灰色） ──
        ver_lbl = ttk.Label(
            m, text=f"v{APP_VERSION}",
            foreground="#b0b0b0",
            font=("Consolas", 9),
        )
        ver_lbl.pack(anchor="e", pady=(0, 2))

        # ── 盘符选择区 ──
        self.drive_frame = ttk.LabelFrame(
            m, text="盘符选择（仅 G: 及之后）", padding=8
        )
        self.drive_frame.pack(fill="x", pady=(0, 4))

        row1 = ttk.Frame(self.drive_frame)
        row1.pack(fill="x")

        # ── 启动时用快速检测（不调 PowerShell），界面秒开 ──
        drives = get_drives_fast('G')
        values = [f"{d[0]}  [{d[1]}]" for d in drives]
        self.drive_var = tk.StringVar(value=values[0] if values else "")
        self.combo = ttk.Combobox(
            row1, textvariable=self.drive_var,
            values=values, state="readonly",
            width=18, font=("Consolas", 11)
        )
        self.combo.pack(side="left", padx=(0, 8))
        self.combo.bind("<<ComboboxSelected>>", self._on_drive_selected)

        ttk.Button(row1, text="刷新盘符", command=self.refresh).pack(side="left")
        self.status_lbl = ttk.Label(row1, text="", foreground="gray")
        self.status_lbl.pack(side="left", padx=10)
        if not drives:
            self.status_lbl.config(
                text="!! 未检测到 G: 及之后的盘符", foreground="red"
            )
        else:
            self.status_lbl.config(
                text="正在识别磁盘类型...", foreground="blue"
            )

        row2 = ttk.Frame(self.drive_frame)
        row2.pack(fill="x", pady=(6, 0))
        self.show_def_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            row2,
            text="启用 D: / E: / F: 盘（⚠ 通常为本地硬盘，谨慎操作）",
            variable=self.show_def_var,
            command=self._toggle_def,
        ).pack(side="left")

        # ── Notebook ──
        nb = ttk.Notebook(m)
        nb.pack(fill="x", pady=4)
        gk = dict(sticky="nsew", padx=3, pady=3, ipady=2)

        # ---------- Tab 1: 解除占用 / 弹出 ----------
        t1 = ttk.Frame(nb, padding=6)
        nb.add(t1, text=" 解除占用 / 弹出 ")
        t1.columnconfigure(0, weight=1)
        t1.columnconfigure(1, weight=1)

        ttk.Button(t1, text="检测占用进程和服务",
                   command=self.detect).grid(row=0, column=0, **gk)
        ttk.Button(t1, text="一键停止占用服务",
                   command=self.stop_svc).grid(row=0, column=1, **gk)
        ttk.Button(t1, text="恢复已停止的服务",
                   command=self.start_svc).grid(row=1, column=0, **gk)
        ttk.Button(t1, text="恢复脱机磁盘",
                   command=self.recover_offline).grid(row=1, column=1, **gk)
        ttk.Separator(t1).grid(row=2, column=0, columnspan=2, sticky="ew", pady=4)

        rec_btn = self._make_rec_btn(t1, self.smart_eject)
        rec_btn.grid(row=3, column=0, sticky="nsew", padx=3, pady=3)
        ttk.Button(t1, text="强制弹出\n直接弹出硬盘",
                   command=self.force_eject).grid(row=3, column=1, **gk)

        # ---------- Tab 2: 删除系统文件夹 ----------
        t2 = ttk.Frame(nb, padding=6)
        nb.add(t2, text=" 删除系统文件夹 ")
        t2.columnconfigure(0, weight=1)
        t2.columnconfigure(1, weight=1)

        ttk.Button(t2, text="删除\nSystem Volume Information",
                   command=self.del_svi).grid(row=0, column=0, **gk)
        ttk.Button(t2, text="删除\n$RECYCLE.BIN",
                   command=self.del_rec).grid(row=0, column=1, **gk)
        ttk.Button(t2, text="一键删除以上两个文件夹",
                   command=self.del_both).grid(row=1, column=0, columnspan=2, **gk)

        # ---------- Tab 3: SYSTEM 写入权限 ----------
        t3 = ttk.Frame(nb, padding=6)
        nb.add(t3, text=" SYSTEM 写入权限 ")
        t3.columnconfigure(0, weight=1)
        t3.columnconfigure(1, weight=1)

        ttk.Button(t3, text="禁止 SYSTEM 写入",
                   command=self.deny_write).grid(row=0, column=0, **gk)
        ttk.Button(t3, text="恢复 SYSTEM 写入",
                   command=self.allow_write).grid(row=0, column=1, **gk)
        ttk.Label(
            t3, foreground="gray",
            text="提示：禁止写入后，系统服务将无法在该盘创建任何文件。"
                 "\n如需恢复，请在拔盘前点击【恢复】按钮。"
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # ── 日志区域 ──
        f4 = ttk.LabelFrame(m, text="执行日志", padding=4)
        f4.pack(fill="both", expand=True, pady=(4, 0))

        # ── 先 pack 底部按钮栏（保证窗口再小也能显示） ──
        btn_bar = ttk.Frame(f4)
        btn_bar.pack(side="bottom", fill="x", pady=(2, 0))
        ttk.Button(
            btn_bar, text="清空日志",
            command=lambda: self.log.delete("1.0", tk.END)
        ).pack(anchor="e")

        # ── 再 pack 日志文本框（填满剩余空间） ──
        self.log = scrolledtext.ScrolledText(
            f4, height=8, font=("Consolas", 10), wrap=tk.WORD
        )
        self.log.pack(fill="both", expand=True)

        self.log_msg("[OK] 工具已启动（管理员模式）")
        ds = ", ".join(f"{d[0]}[{d[1]}]" for d in drives) if drives else "无"
        self.log_msg(f"检测到盘符：{ds}")
        self.log_msg("正在后台识别磁盘总线类型...\n")

    # ========== 后台总线类型检测 ==========

    def _start_bus_detection(self):
        """启动后台 PowerShell 检测磁盘总线类型"""
        self._detecting = True
        threading.Thread(target=self._do_bus_detection, daemon=True).start()

    def _do_bus_detection(self):
        min_letter = 'D' if self.show_def_var.get() else 'G'
        drives, bus_types = get_drives_full(min_letter)
        self._bus_cache = bus_types
        self.root.after(0, lambda: self._apply_bus_detection(drives))

    def _apply_bus_detection(self, drives):
        """后台检测完成，更新 UI"""
        self._detecting = False
        values = [f"{d[0]}  [{d[1]}]" for d in drives]

        cur = self.drive_var.get().split()[0] if self.drive_var.get() else ""
        self.combo["values"] = values
        hit = next((v for v in values if v.startswith(cur)),
                   values[0] if values else "")
        self.drive_var.set(hit)

        ds = ", ".join(f"{d[0]}[{d[1]}]" for d in drives) if drives else "无"
        self.log_msg(f"[识别完成] 盘符：{ds}\n")

        if not drives:
            self.status_lbl.config(text="!! 未检测到可用盘符", foreground="red")
        else:
            self._on_drive_selected()

    # ========== 获取总线类型（带缓存） ==========

    def _get_bus_type(self, letter):
        """优先用缓存，缓存没有则单独查询"""
        letter = letter.rstrip(":\\").upper()
        if letter in self._bus_cache:
            return self._bus_cache[letter]
        bus = get_drive_bus_type(letter)
        self._bus_cache[letter] = bus
        return bus

    # ========== D/E/F 开关 ==========

    def _toggle_def(self):
        if self.show_def_var.get():
            msg = (
                "⚠ 安全提示\n\n"
                "D: / E: / F: 通常是本地固定硬盘分区。\n\n"
                "对本地硬盘执行弹出操作可能导致：\n"
                "• 系统不稳定或蓝屏\n"
                "• 正在运行的程序崩溃\n"
                "• 该分区上的数据丢失\n\n"
                "请确认你了解风险后再操作。\n\n"
                "确定要启用吗？"
            )
            if not messagebox.askyesno("安全提示", msg, icon="warning"):
                self.show_def_var.set(False)
                return
            self.drive_frame.config(text="盘符选择（D: 及之后 ⚠ 含本地硬盘）")
            self.log_msg("[设置] 已启用 D:/E:/F: 盘显示（请谨慎操作）")
        else:
            self.drive_frame.config(text="盘符选择（仅 G: 及之后）")
            self.log_msg("[设置] 已关闭 D:/E:/F: 盘显示")
        self.refresh()

    def _on_drive_selected(self, event=None):
        if self._busy:
            return
        v = self.drive_var.get()
        if not v:
            return
        if self._detecting:
            self.status_lbl.config(text="正在识别磁盘类型...", foreground="blue")
            return
        if "[固定]" in v:
            self.status_lbl.config(
                text="⚠ 固定硬盘 谨慎操作", foreground="red"
            )
        elif "[网络]" in v:
            self.status_lbl.config(
                text="⚠ 网络硬盘 谨慎操作", foreground="#b34700"
            )
        elif "[光驱]" in v:
            self.status_lbl.config(text="光驱", foreground="gray")
        elif "[USB硬盘]" in v:
            self.status_lbl.config(text="USB 移动硬盘", foreground="green")
        elif "[可移动]" in v:
            self.status_lbl.config(text="可移动设备", foreground="green")
        else:
            self.status_lbl.config(text="", foreground="gray")

    # ========== 驱动器安全检查 ==========

    def _check_drive_safety(self, d):
        dt = get_drive_type_code(d)
        bus = self._get_bus_type(d[0])
        is_usb = bus.upper() in USB_BUS_TYPES

        if dt == 3 and is_usb:
            return True

        if dt == 3 and not is_usb:
            msg = (
                f"⚠ 安全警告\n\n"
                f"{d}\\ 被识别为【本地固定硬盘】（总线: {bus}），\n"
                f"不是 USB 可移动设备！\n\n"
                f"弹出本地固定硬盘可能导致：\n"
                f"• 系统不稳定、蓝屏或崩溃\n"
                f"• 正在运行的程序意外关闭\n"
                f"• 该磁盘分区上的数据丢失\n\n"
                f"强烈建议：仅对 USB 移动硬盘 / U盘 执行弹出。\n\n"
                f"是否仍要继续？"
            )
            return messagebox.askyesno("⚠ 安全警告", msg, icon="warning")

        elif dt == 4:
            msg = (
                f"⚠ 安全提示\n\n"
                f"{d}\\ 被识别为【网络驱动器】，\n"
                f"不是本地可移动设备！\n\n"
                f"断开网络驱动器可能导致：\n"
                f"• 正在访问的网络文件 / 程序中断\n"
                f"• 需要重新映射网络驱动器\n\n"
                f"是否仍要继续？"
            )
            return messagebox.askyesno("⚠ 安全提示", msg, icon="warning")

        elif dt == 5:
            msg = f"提示：{d}\\ 是光驱，弹出将打开光驱托盘。\n\n继续？"
            return messagebox.askyesno("提示", msg)

        return True

    # ========== 启动时检查脱机磁盘 ==========

    def _check_offline_on_start(self):
        threading.Thread(target=self._do_check_offline_start, daemon=True).start()

    def _do_check_offline_start(self):
        offline = get_offline_disks()
        if offline:
            usb_offline = [d for d in offline if d["bus"] in ("USB", "USB3")]
            if usb_offline:
                info_lines = []
                for d in usb_offline:
                    info_lines.append(
                        f"  磁盘 {d['number']}: {d['name']} "
                        f"({d['size_gb']:.1f} GB, {d['bus']})"
                    )
                self.log_msg("[注意] 检测到以下 USB 磁盘处于脱机状态：")
                for ln in info_lines:
                    self.log_msg(ln)
                self.log_msg("这可能是上次使用 Set-Disk -IsOffline 弹出导致的。")
                self.log_msg("点击【恢复脱机磁盘】按钮可恢复。\n")
                self.root.after(0, lambda: self.status_lbl.config(
                    text=f"⚠ 检测到 {len(usb_offline)} 个脱机USB磁盘",
                    foreground="orange"
                ))
            elif offline:
                self.log_msg(
                    f"[信息] 检测到 {len(offline)} 个脱机磁盘"
                    f"（非USB），可能为正常状态。\n"
                )

    # ========== 通用 ==========

    def get_drive(self):
        v = self.drive_var.get()
        if not v:
            messagebox.showwarning("提示", "请先选择一个盘符")
            return None
        d = v.split()[0]
        if not drive_exists(d):
            messagebox.showwarning("提示", f"{d}\\ 不可访问")
            return None
        return d

    def refresh(self):
        min_letter = 'D' if self.show_def_var.get() else 'G'
        # 先用快速检测立即刷新列表
        drives = get_drives_fast(min_letter)
        values = [f"{d[0]}  [{d[1]}]" for d in drives]
        self.combo["values"] = values
        cur = self.drive_var.get().split()[0] if self.drive_var.get() else ""
        hit = next((v for v in values if v.startswith(cur)),
                   values[0] if values else "")
        self.drive_var.set(hit)

        ds = ", ".join(f"{d[0]}[{d[1]}]" for d in drives) if drives else "无"
        self.log_msg(f"[刷新] 盘符：{ds}（正在识别类型...）")

        if not drives:
            self.status_lbl.config(text="!! 未检测到可用盘符", foreground="red")
        else:
            self.status_lbl.config(text="正在识别磁盘类型...", foreground="blue")

        # 后台完整检测
        self._start_bus_detection()
        threading.Thread(target=self._do_check_offline_start, daemon=True).start()

    def log_msg(self, msg):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.root.update_idletasks()

    def exec_cmd(self, cmd, timeout=120):
        self.log_msg(f"  > {cmd}")
        rc, out, err = run_cmd(cmd, timeout)
        if out.strip():
            for ln in out.strip().splitlines():
                self.log_msg(f"    {ln}")
        if err.strip():
            for ln in err.strip().splitlines():
                self.log_msg(f"    [!] {ln}")
        self.root.update_idletasks()
        return rc, out, err

    def run_in_thread(self, func):
        if self._busy:
            messagebox.showinfo("提示", "有操作正在执行，请稍候")
            return
        self._busy = True
        self.status_lbl.config(text="执行中...", foreground="blue")

        def wrapper():
            try:
                func()
            except Exception as e:
                self.log_msg(f"[异常] {e}")
            finally:
                self._busy = False
                self.root.after(0, lambda: self._on_drive_selected())

        threading.Thread(target=wrapper, daemon=True).start()

    # ========== 恢复脱机磁盘 ==========

    def recover_offline(self):
        self.run_in_thread(self._do_recover_offline)

    def _do_recover_offline(self):
        self.log_msg("\n--- 检测脱机磁盘 ---")
        offline = get_offline_disks()
        if not offline:
            self.log_msg("  未检测到脱机磁盘，一切正常。\n")
            return

        self.log_msg(f"  发现 {len(offline)} 个脱机磁盘：")
        for d in offline:
            self.log_msg(
                f"    磁盘 {d['number']}: {d['name']} "
                f"({d['size_gb']:.1f} GB, 总线: {d['bus']})"
            )

        self.log_msg("\n  正在恢复联机...")
        success = 0
        for d in offline:
            self.log_msg(f"    Set-Disk -Number {d['number']} -IsOffline $false ...")
            ok = set_disk_online(d["number"])
            if ok:
                self.log_msg(f"    [OK] 磁盘 {d['number']} 已恢复联机")
                success += 1
            else:
                self.log_msg(f"    [!!] 磁盘 {d['number']} 恢复失败")

        self.log_msg(f"\n[OK] 共恢复 {success}/{len(offline)} 个磁盘")
        self.log_msg("正在刷新盘符...\n")
        time.sleep(2)
        self.root.after(0, self.refresh)

    # ========== USB 安全移除 ==========

    def _get_disk_number(self, letter):
        rc, out, _ = run_cmd(
            f'powershell -Command "'
            f"(Get-Partition -DriveLetter {letter} -ErrorAction Stop).DiskNumber"
            f'"', timeout=10
        )
        if rc == 0 and out.strip().isdigit():
            return int(out.strip())
        return None

    def _usb_safe_remove(self, disk_number):
        ps_file = os.path.join(os.environ.get("TEMP", "."), "_usb_eject.ps1")
        try:
            with open(ps_file, "w", encoding="utf-8-sig") as f:
                f.write(USB_EJECT_PS1)
            cmd = (
                f'powershell -ExecutionPolicy Bypass '
                f'-File "{ps_file}" -DiskNumber {disk_number}'
            )
            rc, out, err = run_cmd(cmd, timeout=20)
            if out.strip():
                for ln in out.strip().splitlines():
                    self.log_msg(f"    {ln}")
            if err.strip():
                for ln in err.strip().splitlines():
                    self.log_msg(f"    [!] {ln}")
            return "USB_EJECT_OK" in out
        except Exception as e:
            self.log_msg(f"    [!] 异常: {e}")
            return False
        finally:
            try:
                os.remove(ps_file)
            except OSError:
                pass

    # ========== 多方法弹出 ==========

    def _try_eject(self, d):
        letter = d[0]

        self.log_msg("\n  获取磁盘信息...")
        disk_number = self._get_disk_number(letter)
        if disk_number is not None:
            self.log_msg(f"    物理磁盘号: {disk_number}")
        else:
            self.log_msg("    [!] 无法获取磁盘号，USB 安全移除可能不可用")

        if disk_number is not None:
            self.log_msg("\n  方法1: USB 安全移除 (CM_Request_Device_Eject) ...")
            usb_ok = self._usb_safe_remove(disk_number)
            if usb_ok:
                time.sleep(2)
                if not drive_exists(d):
                    self.log_msg("    盘符已消失，硬盘已安全移除并停转!")
                    return True
                self.log_msg("    弹出指令成功但盘符仍在，继续...")
            else:
                self.log_msg("    失败（可能有程序占用），尝试下一方法...")

        self.log_msg("\n  方法2: DeviceIoControl API ...")
        ok, msg = eject_volume_api(d)
        self.log_msg(f"    {msg}")
        time.sleep(2)
        if not drive_exists(d):
            self.log_msg("    盘符已消失!")
            if disk_number is not None:
                self.log_msg("    追加 USB 安全移除以停转硬盘...")
                self._usb_safe_remove(disk_number)
            return True
        self.log_msg("    盘符仍在，尝试下一方法...")

        self.log_msg("\n  方法3: Shell.Application Eject ...")
        cmd = (
            'powershell -Command "'
            "(New-Object -ComObject Shell.Application)"
            ".NameSpace(17).ParseName('"
            + letter
            + ":').InvokeVerb('Eject')"
            '"'
        )
        self.exec_cmd(cmd, timeout=15)
        time.sleep(3)
        if not drive_exists(d):
            self.log_msg("    盘符已消失!")
            if disk_number is not None:
                self.log_msg("    追加 USB 安全移除以停转硬盘...")
                self._usb_safe_remove(disk_number)
            return True
        self.log_msg("    盘符仍在，尝试下一方法...")

        self.log_msg("\n  方法4: Set-Disk -IsOffline + USB 安全移除 ...")
        self.log_msg("    [注意] 此方法如果 USB 移除失败，下次插入可能需要手动恢复联机")
        cmd = (
            'powershell -Command "'
            "$p = Get-Partition -DriveLetter "
            + letter
            + " -ErrorAction Stop; "
            "$dk = $p | Get-Disk; "
            'Set-Disk -Number $dk.Number -IsOffline $true"'
        )
        self.exec_cmd(cmd, timeout=15)
        time.sleep(2)
        if not drive_exists(d):
            self.log_msg("    盘符已消失!")
            if disk_number is not None:
                self.log_msg("    追加 USB 安全移除以停转硬盘并清除脱机标记...")
                usb_ok = self._usb_safe_remove(disk_number)
                if usb_ok:
                    self.log_msg("    硬盘已安全移除并停转!")
                else:
                    self.log_msg("    [!] USB 安全移除失败")
                    self.log_msg("    [!] 硬盘可能仍在转动，数据已安全可拔出")
                    self.log_msg("    [!] 下次插入如无盘符，请点击【恢复脱机磁盘】")
            return True
        self.log_msg("    盘符仍在，尝试下一方法...")

        self.log_msg("\n  方法5: diskpart ...")
        tmp = os.path.join(os.environ.get("TEMP", "."), "_eject.txt")
        with open(tmp, "w") as f:
            f.write(f"select volume {letter}\nremove all dismount\n")
        self.exec_cmd(f'diskpart /s "{tmp}"', timeout=30)
        try:
            os.remove(tmp)
        except OSError:
            pass
        time.sleep(2)
        if not drive_exists(d):
            self.log_msg("    盘符已消失!")
            if disk_number is not None:
                self.log_msg("    追加 USB 安全移除以停转硬盘...")
                self._usb_safe_remove(disk_number)
            return True

        return False

    # ========== 检测 ==========

    def detect(self):
        d = self.get_drive()
        if not d:
            return
        self.run_in_thread(lambda: self._detect(d))

    def _detect(self, d):
        self.log_msg(f"\n{'='*50}")
        self.log_msg(f"  检测占用 {d}\\ 的进程和服务")
        self.log_msg(f"{'='*50}")

        bus = self._get_bus_type(d[0])
        dt = get_drive_type_code(d)
        type_name = DRIVE_TYPE_MAP.get(dt, "未知")
        self.log_msg(f"\n  磁盘类型: {type_name}  总线: {bus}")

        self.log_msg("\n[1] 在该盘上运行的进程：")
        cmd = (
            'powershell -Command "'
            "Get-Process | Where-Object { $_.Path -like '"
            + d
            + "\\*' } | "
            'Format-Table Id,Name,Path -AutoSize | Out-String -Width 300"'
        )
        rc, out, _ = self.exec_cmd(cmd)
        if not out.strip():
            self.log_msg("    （无）")

        self.log_msg("\n[2] 加载了该盘文件的进程：")
        ps = (
            'powershell -Command "Get-Process | ForEach-Object { $p=$_; try { '
            "$p.Modules | Where-Object { $_.FileName -like '"
            + d
            + "\\*' } | "
            "ForEach-Object { Write-Output ('PID={0}  {1}  {2}' -f "
            "$p.Id,$p.Name,$_.FileName) }"
            ' } catch {} }"'
        )
        rc, out, _ = self.exec_cmd(ps)
        if not out.strip():
            self.log_msg("    （无）")

        self.log_msg("\n[3] 常见占用服务状态：")
        for name, display in SERVICES:
            st = svc_status(name)
            if st == "running":
                self.log_msg(f"    * {display} ({name})  ->  运行中 !!")
            elif st == "stopped":
                self.log_msg(f"    o {display} ({name})  ->  已停止")
            else:
                self.log_msg(f"    - {display} ({name})  ->  未安装")

        self.log_msg("\n[4] openfiles 查询：")
        rc, out, err = run_cmd("openfiles /query /fo csv", timeout=10)
        if rc == 0 and out.strip():
            hits = [l for l in out.splitlines() if d.upper() in l.upper()]
            if hits:
                for h in hits[:30]:
                    self.log_msg(f"    {h}")
            else:
                self.log_msg("    未查到占用记录")
        else:
            self.log_msg("    不可用（需先运行 openfiles /local on 并重启）")
        self.log_msg("")

    # ========== 服务管理 ==========

    def stop_svc(self):
        self.run_in_thread(self._stop_svc)

    def _stop_svc(self):
        self.log_msg("\n--- 停止占用服务 ---")
        self._svc_was_running = {}
        for name, display in SERVICES:
            st = svc_status(name)
            if st == "running":
                self.log_msg(f"  停止 {display} ({name}) ...")
                self.exec_cmd(f"net stop {name} /y", timeout=30)
                self._svc_was_running[name] = True
            elif st == "stopped":
                self.log_msg(f"  跳过 {display} - 已是停止状态")
            else:
                self.log_msg(f"  跳过 {display} - 未安装")
        stopped = len(self._svc_was_running)
        self.log_msg(f"\n[OK] 共停止 {stopped} 个服务\n")

    def start_svc(self):
        self.run_in_thread(self._start_svc)

    def _start_svc(self):
        self.log_msg("\n--- 恢复服务 ---")
        targets = self._svc_was_running if self._svc_was_running else {
            name: True for name, _ in SERVICES
        }
        count = 0
        for name, display in SERVICES:
            if name not in targets:
                continue
            st = svc_status(name)
            if st == "stopped":
                self.log_msg(f"  启动 {display} ({name}) ...")
                self.exec_cmd(f"net start {name}", timeout=30)
                count += 1
            elif st == "running":
                self.log_msg(f"  跳过 {display} - 已在运行")
            else:
                self.log_msg(f"  跳过 {display} - 未安装")
        self._svc_was_running = {}
        self.log_msg(f"\n[OK] 共恢复 {count} 个服务\n")

    # ========== 弹出 ==========

    def smart_eject(self):
        d = self.get_drive()
        if not d:
            return
        if not self._check_drive_safety(d):
            self.log_msg(f"[取消] 用户取消了对 {d} 的弹出操作\n")
            return
        msg = (
            f"将执行以下步骤：\n\n"
            f"1. 停止常见占用服务\n"
            f"2. 多种方式尝试弹出 {d}（含 USB 硬件级安全移除）\n"
            f"3. 恢复服务\n\n继续？"
        )
        if not messagebox.askyesno("安全弹出", msg):
            return
        self.run_in_thread(lambda: self._smart_eject(d))

    def _smart_eject(self, d):
        self.log_msg(f"\n{'='*50}")
        self.log_msg(f"  安全弹出 {d}")
        self.log_msg(f"{'='*50}")

        self.log_msg("\n步骤 1：停止占用服务...")
        self._svc_was_running = {}
        for name, display in SERVICES:
            st = svc_status(name)
            if st == "running":
                self.log_msg(f"  停止 {display} ...")
                run_cmd(f"net stop {name} /y", timeout=20)
                self._svc_was_running[name] = True
            else:
                self.log_msg(f"  跳过 {display} ({st})")

        time.sleep(1)

        self.log_msg("\n步骤 2：弹出硬盘（逐一尝试多种方法）...")
        ok = self._try_eject(d)

        if ok:
            self.log_msg(f"\n[OK] {d} 已成功弹出！可以安全拔出硬盘。")
        else:
            self.log_msg(f"\n[!!] {d} 仍然存在，所有弹出方法均失败。")
            self.log_msg("     请关闭该盘上所有打开的文件/窗口后重试，")
            self.log_msg("     或使用【检测】功能查看哪些进程在占用。")

        self.log_msg("\n步骤 3：恢复服务...")
        for name, display in SERVICES:
            if name in self._svc_was_running:
                run_cmd(f"net start {name}", timeout=10)
                self.log_msg(f"  已恢复 {display}")
        self._svc_was_running = {}
        self.log_msg("[OK] 服务恢复完成\n")

    def force_eject(self):
        d = self.get_drive()
        if not d:
            return
        if not self._check_drive_safety(d):
            self.log_msg(f"[取消] 用户取消了对 {d} 的弹出操作\n")
            return
        msg = f"跳过停止服务，直接弹出 {d}？"
        if not messagebox.askyesno("强制弹出", msg):
            return
        self.run_in_thread(lambda: self._force_eject(d))

    def _force_eject(self, d):
        self.log_msg(f"\n--- 强制弹出 {d} ---")
        ok = self._try_eject(d)
        if ok:
            self.log_msg(f"\n[OK] {d} 已弹出！可以安全拔出硬盘。\n")
        else:
            self.log_msg(f"\n[!!] {d} 仍然存在，弹出失败。")
            self.log_msg("     请关闭占用该盘的程序后重试。\n")

    # ========== 删除文件夹 ==========

    def _del_folder(self, d, name, takeown=False):
        path = f"{d}\\{name}"
        if not os.path.exists(path):
            self.log_msg(f"  {path} 不存在，跳过")
            return
        self.log_msg(f"\n--- 删除 {path} ---")
        if takeown:
            self.exec_cmd(f'takeown /f "{path}" /r /d y', timeout=180)
            self.exec_cmd(
                f'icacls "{path}" /grant administrators:F /t', timeout=180
            )
        self.exec_cmd(f'cmd /c rd /s /q "{path}"')
        if os.path.exists(path):
            self.log_msg(f"  [!!] {path} 可能未完全删除")
        else:
            self.log_msg(f"  [OK] {path} 已删除")

    def del_svi(self):
        d = self.get_drive()
        if not d:
            return
        if not messagebox.askyesno(
            "确认", f"删除 {d}\\System Volume Information？"
        ):
            return
        self.run_in_thread(lambda: self._do_del_svi(d))

    def _do_del_svi(self, d):
        self._del_folder(d, "System Volume Information", True)
        self.log_msg("")

    def del_rec(self):
        d = self.get_drive()
        if not d:
            return
        if not messagebox.askyesno("确认", f"删除 {d}\\$RECYCLE.BIN？"):
            return
        self.run_in_thread(lambda: self._do_del_rec(d))

    def _do_del_rec(self, d):
        self._del_folder(d, "$RECYCLE.BIN", True)
        self.log_msg("")

    def del_both(self):
        d = self.get_drive()
        if not d:
            return
        if not messagebox.askyesno("确认", f"删除 {d} 上两个系统文件夹？"):
            return
        self.run_in_thread(lambda: self._do_del_both(d))

    def _do_del_both(self, d):
        self._del_folder(d, "System Volume Information", True)
        self._del_folder(d, "$RECYCLE.BIN", True)
        self.log_msg("")

    # ========== 权限控制 ==========

    def deny_write(self):
        d = self.get_drive()
        if not d:
            return
        msg = (
            f"禁止 SYSTEM 写入 {d}\\ ？\n\n"
            f"效果：系统无法在该盘自动创建文件夹\n"
            f"恢复：随时点击【恢复】按钮"
        )
        if not messagebox.askyesno("确认", msg):
            return
        self.run_in_thread(lambda: self._deny(d))

    def _deny(self, d):
        self.log_msg(f"\n--- 禁止 SYSTEM 写入 {d}\\ ---")
        self.exec_cmd(f'icacls {d}\\ /deny "SYSTEM:(WD)" /T /C')
        self.log_msg("[OK] 已禁止\n")

    def allow_write(self):
        d = self.get_drive()
        if not d:
            return
        if not messagebox.askyesno("确认",
                                    f"恢复 SYSTEM 对 {d}\\ 的写入权限？"):
            return
        self.run_in_thread(lambda: self._allow(d))

    def _allow(self, d):
        self.log_msg(f"\n--- 恢复 SYSTEM 写入 {d}\\ ---")
        self.exec_cmd(f'icacls {d}\\ /remove:d "SYSTEM" /T /C')
        self.log_msg("[OK] 已恢复\n")


if __name__ == "__main__":
    if not is_admin():
        script = os.path.abspath(sys.argv[0])
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, f'"{script}"', None, 1
        )
        sys.exit(0)

    try:
        App()
    except Exception as e:
        messagebox.showerror("启动失败", str(e))