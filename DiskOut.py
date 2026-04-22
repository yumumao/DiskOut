# -*- coding: utf-8 -*-
"""
移动硬盘安全清理工具
支持 USB 硬件级安全弹出（停止转动）
自动检测并恢复脱机磁盘
支持检测文件/文件夹占用进程和服务并一键结束
支持普通模式运行，需要时提示提升管理员权限
提权后自动恢复上次操作状态（含检测结果）
管理员模式下自动放行 UIPI 拖放限制
"""
import ctypes
import sys
import os
import subprocess
import string
import time
import json
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from ctypes import wintypes
import threading

# ── 尝试导入拖放支持库（pip install windnd） ──
try:
    import windnd
    HAS_WINDND = True
except ImportError:
    HAS_WINDND = False

# ── 版本号（修改此处即可更新界面右上角显示） ──
APP_VERSION = "3.3.2"    # ★ 修改：版本号更新

SERVICES = [
    ("WSearch",       "Windows Search"),
    ("SysMain",       "SysMain"),
    ("VSS",           "Volume Shadow Copy"),
    ("defragsvc",     "Optimize Drives"),
    ("WMPNetworkSvc", "WMP Network Sharing"),
    ("StorSvc",       "Storage Service"),
]

# ── 服务恢复建议 ──
SERVICE_RECOMMENDATIONS = {
    "WSearch":       ("建议恢复", "文件搜索和索引需要此服务"),
    "SysMain":       ("建议恢复", "优化应用启动速度和系统性能"),
    "VSS":           ("建议恢复", "系统还原和备份软件依赖此服务"),
    "defragsvc":     ("可选",     "磁盘优化按计划运行，不急需可稍后恢复"),
    "WMPNetworkSvc": ("可选",     "仅在使用 Windows Media Player 共享时需要"),
    "StorSvc":       ("建议恢复", "管理存储设置和可移动存储策略"),
}

DRIVE_TYPE_MAP = {
    2: "可移动",
    3: "固定",
    4: "网络",
    5: "光驱",
    6: "RAM",
}

USB_BUS_TYPES = {"USB", "USB3"}

# ── 提权状态文件路径 ──
STATE_FILE_PATH = os.path.join(
    os.environ.get("TEMP", os.path.dirname(os.path.abspath(sys.argv[0]))),
    "_diskout_elevation_state.json"
)

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


# ★ 资源路径兼容函数（PyInstaller --onefile 打包后资源解压到 sys._MEIPASS）
def resource_path(relative_path):
    """获取资源文件的绝对路径，兼容 PyInstaller 打包和开发环境"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), relative_path)


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


# ════════════════════════════════════════════════════════════════
#  PowerShell 查询总线类型（保留作为 fallback）
# ════════════════════════════════════════════════════════════════

def get_drive_bus_types():
    cmd = (
        'powershell -NoProfile -Command "'
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


def get_drive_bus_type(letter):
    letter = letter.rstrip(":\\").upper()
    cmd = (
        f'powershell -NoProfile -Command "'
        f"try {{ $p = Get-Partition -DriveLetter {letter} -ErrorAction Stop; "
        f"($p | Get-Disk).BusType }} catch {{ Write-Output 'Unknown' }}"
        f'"'
    )
    rc, out, _ = run_cmd(cmd, timeout=10)
    if rc == 0 and out.strip():
        return out.strip()
    return "Unknown"


# ════════════════════════════════════════════════════════════════
#  IOCTL 快速磁盘信息查询（纯 ctypes，无需 subprocess，毫秒级）
# ════════════════════════════════════════════════════════════════

_IOCTL_STORAGE_QUERY_PROPERTY    = 0x002D1400
_IOCTL_STORAGE_GET_DEVICE_NUMBER = 0x002D1080

_BUS_TYPE_NAMES = {
    0: "Unknown", 1: "SCSI", 2: "ATAPI", 3: "ATA", 4: "1394",
    5: "SSA", 6: "Fibre", 7: "USB", 8: "RAID", 9: "iSCSI",
    10: "SAS", 11: "SATA", 12: "SD", 13: "MMC", 14: "Virtual",
    15: "FileBackedVirtual", 16: "Spaces", 17: "NVMe", 18: "SCM", 19: "UFS",
}


class _STORAGE_PROPERTY_QUERY(ctypes.Structure):
    _fields_ = [
        ("PropertyId", ctypes.c_ulong),
        ("QueryType",  ctypes.c_ulong),
        ("Extra",      ctypes.c_byte * 1),
    ]


class _STORAGE_DEVICE_DESCRIPTOR(ctypes.Structure):
    _fields_ = [
        ("Version",               ctypes.c_ulong),
        ("Size",                  ctypes.c_ulong),
        ("DeviceType",            ctypes.c_ubyte),
        ("DeviceTypeModifier",    ctypes.c_ubyte),
        ("RemovableMedia",        ctypes.c_ubyte),
        ("CommandQueueing",       ctypes.c_ubyte),
        ("VendorIdOffset",        ctypes.c_ulong),
        ("ProductIdOffset",       ctypes.c_ulong),
        ("ProductRevisionOffset", ctypes.c_ulong),
        ("SerialNumberOffset",    ctypes.c_ulong),
        ("BusType",               ctypes.c_ulong),
    ]


class _STORAGE_DEVICE_NUMBER(ctypes.Structure):
    _fields_ = [
        ("DeviceType",      ctypes.c_ulong),
        ("DeviceNumber",    ctypes.c_ulong),
        ("PartitionNumber", ctypes.c_ulong),
    ]


def _open_volume_handle(letter):
    letter = letter.rstrip(":\\").upper()
    k32 = ctypes.windll.kernel32
    k32.CreateFileW.restype = ctypes.c_void_p
    INVALID = ctypes.c_void_p(-1).value
    h = k32.CreateFileW(f"\\\\.\\{letter}:", 0, 0x3, None, 3, 0, None)
    if h is None or h == INVALID:
        return None
    return h


def get_bus_type_ioctl(letter):
    h = _open_volume_handle(letter)
    if h is None:
        return "Unknown"
    k32 = ctypes.windll.kernel32
    try:
        query = _STORAGE_PROPERTY_QUERY()
        query.PropertyId = 0
        query.QueryType  = 0

        buf = (ctypes.c_byte * 1024)()
        returned = wintypes.DWORD(0)

        ok = k32.DeviceIoControl(
            ctypes.c_void_p(h), _IOCTL_STORAGE_QUERY_PROPERTY,
            ctypes.byref(query), ctypes.sizeof(query),
            ctypes.byref(buf), 1024,
            ctypes.byref(returned), None,
        )
        if ok and returned.value >= ctypes.sizeof(_STORAGE_DEVICE_DESCRIPTOR):
            desc = _STORAGE_DEVICE_DESCRIPTOR.from_buffer(buf)
            return _BUS_TYPE_NAMES.get(desc.BusType, f"Type{desc.BusType}")
        return "Unknown"
    except Exception:
        return "Unknown"
    finally:
        k32.CloseHandle(ctypes.c_void_p(h))


def get_disk_number_ioctl(letter):
    h = _open_volume_handle(letter)
    if h is None:
        return None
    k32 = ctypes.windll.kernel32
    try:
        sdn = _STORAGE_DEVICE_NUMBER()
        returned = wintypes.DWORD(0)

        ok = k32.DeviceIoControl(
            ctypes.c_void_p(h), _IOCTL_STORAGE_GET_DEVICE_NUMBER,
            None, 0,
            ctypes.byref(sdn), ctypes.sizeof(sdn),
            ctypes.byref(returned), None,
        )
        return sdn.DeviceNumber if ok else None
    except Exception:
        return None
    finally:
        k32.CloseHandle(ctypes.c_void_p(h))


def get_all_bus_types_ioctl(min_letter='G'):
    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    result = {}
    for i, ch in enumerate(string.ascii_uppercase):
        if ch < min_letter:
            continue
        if bitmask & (1 << i):
            result[ch] = get_bus_type_ioctl(ch)
    return result


# ════════════════════════════════════════════════════════════════
#  组合检测函数
# ════════════════════════════════════════════════════════════════

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

    bus_types = get_all_bus_types_ioctl(min_letter)
    all_unknown = all(v == "Unknown" for v in bus_types.values())
    if all_unknown and candidates:
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
        'powershell -NoProfile -Command "'
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
        f'powershell -NoProfile -Command "'
        f'Set-Disk -Number {disk_number} -IsOffline $false"'
    )
    rc, _, _ = run_cmd(cmd, timeout=15)
    return rc == 0


def eject_volume_api(letter):
    letter = letter.rstrip(":\\")
    volume = f"\\\\.\\{letter}:"
    k32 = ctypes.windll.kernel32

    k32.CreateFileW.restype = ctypes.c_void_p

    SHARE_RW       = 0x1 | 0x2
    OPEN_EXISTING  = 3
    FSCTL_LOCK     = 0x00090018
    FSCTL_DISMOUNT = 0x00090020
    IOCTL_EJECT    = 0x002D4808
    INVALID_HANDLE = ctypes.c_void_p(-1).value

    warnings = []
    write_access = False

    for access in [0xC0000000, 0x80000000, 0]:
        h = k32.CreateFileW(volume, access, SHARE_RW, None, OPEN_EXISTING, 0, None)
        if h is not None and h != INVALID_HANDLE:
            write_access = (access == 0xC0000000)
            break
    else:
        return False, "无法打开卷句柄", []

    br = wintypes.DWORD(0)

    if write_access:
        if not k32.FlushFileBuffers(h):
            warnings.append("FlushFileBuffers 失败，写缓冲区可能未完全刷新")
    else:
        warnings.append("无写权限打开卷，FlushFileBuffers 可能无效")
        k32.FlushFileBuffers(h)

    locked = k32.DeviceIoControl(
        h, FSCTL_LOCK, None, 0, None, 0, ctypes.byref(br), None
    )
    if not locked:
        warnings.append("卷锁定失败（有程序占用），强制继续")

    dismounted = k32.DeviceIoControl(
        h, FSCTL_DISMOUNT, None, 0, None, 0, ctypes.byref(br), None
    )
    if not dismounted:
        warnings.append("卸载文件系统失败")

    ok = k32.DeviceIoControl(
        h, IOCTL_EJECT, None, 0, None, 0, ctypes.byref(br), None
    )
    k32.CloseHandle(h)

    msg = "API 弹出指令已发送" if ok else "API IOCTL 失败"
    return bool(ok), msg, warnings


# ════════════════════════════════════════════════════════════════
#  Restart Manager API — 检测文件/文件夹占用进程
# ════════════════════════════════════════════════════════════════

class RM_UNIQUE_PROCESS(ctypes.Structure):
    _fields_ = [
        ("dwProcessId", wintypes.DWORD),
        ("ProcessStartTime", wintypes.FILETIME),
    ]


class RM_PROCESS_INFO(ctypes.Structure):
    _fields_ = [
        ("Process", RM_UNIQUE_PROCESS),
        ("strAppName", ctypes.c_wchar * 256),
        ("strServiceShortName", ctypes.c_wchar * 64),
        ("ApplicationType", wintypes.DWORD),
        ("AppStatus", wintypes.DWORD),
        ("TSSessionId", wintypes.DWORD),
        ("bRestartable", wintypes.BOOL),
    ]


def find_locking_processes_rm(paths):
    if isinstance(paths, str):
        paths = [paths]
    if not paths:
        return []
    try:
        rm = ctypes.windll.rstrtmgr
    except OSError:
        return []

    results = []
    session_handle = wintypes.DWORD()
    session_key = ctypes.create_unicode_buffer(33)

    ret = rm.RmStartSession(ctypes.byref(session_handle), 0, session_key)
    if ret != 0:
        return results

    try:
        n = len(paths)
        arr = (ctypes.c_wchar_p * n)(*paths)
        ret = rm.RmRegisterResources(
            session_handle.value, n, arr, 0, None, 0, None
        )
        if ret != 0:
            return results

        needed = wintypes.UINT(0)
        count = wintypes.UINT(0)
        reason = wintypes.DWORD(0)

        ret = rm.RmGetList(
            session_handle.value,
            ctypes.byref(needed), ctypes.byref(count),
            None, ctypes.byref(reason),
        )

        if ret == 234 and needed.value > 0:
            count = wintypes.UINT(needed.value)
            info = (RM_PROCESS_INFO * needed.value)()
            ret = rm.RmGetList(
                session_handle.value,
                ctypes.byref(needed), ctypes.byref(count),
                info, ctypes.byref(reason),
            )
            if ret == 0:
                for i in range(count.value):
                    pid = info[i].Process.dwProcessId
                    if pid == 0:
                        continue
                    results.append({
                        "pid": pid,
                        "name": info[i].strAppName,
                        "service": info[i].strServiceShortName,
                    })
    except Exception:
        pass
    finally:
        rm.RmEndSession(session_handle.value)

    return results


def collect_files_in_dir(path, max_files=300, max_depth=3):
    files = []

    def _scan(dir_path, depth):
        if depth > max_depth or len(files) >= max_files:
            return
        try:
            for entry in os.scandir(dir_path):
                if len(files) >= max_files:
                    return
                if entry.is_file(follow_symlinks=False):
                    files.append(entry.path)
                elif entry.is_dir(follow_symlinks=False) and depth < max_depth:
                    _scan(entry.path, depth + 1)
        except PermissionError:
            pass

    _scan(path, 0)
    return files


# ════════════════════════════════════════════════════════════════
#  SHFileOperation — 删除到回收站
# ════════════════════════════════════════════════════════════════

class _SHFILEOPSTRUCTW(ctypes.Structure):
    _fields_ = [
        ("hwnd",                  ctypes.c_void_p),
        ("wFunc",                 ctypes.c_uint),
        ("pFrom",                 ctypes.c_wchar_p),
        ("pTo",                   ctypes.c_wchar_p),
        ("fFlags",                ctypes.c_ushort),
        ("fAnyOperationsAborted", ctypes.c_int),
        ("hNameMappings",         ctypes.c_void_p),
        ("lpszProgressTitle",     ctypes.c_wchar_p),
    ]


def send_to_recycle_bin(path):
    FO_DELETE          = 0x0003
    FOF_ALLOWUNDO      = 0x0040
    FOF_NOCONFIRMATION = 0x0010

    op = _SHFILEOPSTRUCTW()
    op.hwnd = None
    op.wFunc = FO_DELETE
    op.pFrom = path + '\0'
    op.pTo = None
    op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION
    op.fAnyOperationsAborted = 0
    op.hNameMappings = None
    op.lpszProgressTitle = None

    result = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    return result == 0 and not op.fAnyOperationsAborted


# ════════════════════════════════════════════════════════════════
#  进程是否存活检查
# ════════════════════════════════════════════════════════════════

def is_process_alive(pid):
    """检查 PID 是否仍然存活"""
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    STILL_ACTIVE = 259
    k32 = ctypes.windll.kernel32
    h = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not h:
        return False
    try:
        exit_code = wintypes.DWORD()
        if k32.GetExitCodeProcess(h, ctypes.byref(exit_code)):
            return exit_code.value == STILL_ACTIVE
        return False
    finally:
        k32.CloseHandle(h)


# ════════════════════════════════════════════════════════════════

REC_BG       = "#dae8fc"
REC_FG       = "#1a3a6b"
REC_ACTIVE   = "#b8d4f0"
STAR_COLOR   = "#c8a000"
STAR_HOVER   = "#ffe066"
REC_HOVER_BG = "#3b7dd8"
REC_HOVER_FG = "#ffffff"


class App:
    def __init__(self):
        self._is_admin = is_admin()
        self.root = tk.Tk()
        title_mode = "管理员模式" if self._is_admin else "普通模式"
        self.root.title(f"移动硬盘弹出工具 - {title_mode}")
        self.root.geometry("560x750")
        self.root.minsize(460, 550)
        self._set_icon(self.root)
        self._busy = False
        self._all_stopped_services = {}
        self._bus_cache = {}
        self._detecting = False
        self._file_lock_processes = []
        self._file_lock_services = {}
        self._file_lock_path = ""
        self._detection_is_restored = False   # ★ 标记检测结果来自提权前
        self._dnd_ok = False                  # ★ 新增：标记拖放是否成功初始化
        self.build_ui()
        self._restore_state_from_elevation()
        self.root.after(100, self._start_bus_detection)
        self.root.after(500, self._check_offline_on_start)
        self.root.mainloop()

    def _set_icon(self, window):
        try:
            for ico_name in ("diskout.ico", "DiskOut.ico", "Diskout.ico"):
                ico_path = resource_path(ico_name)
                if os.path.isfile(ico_path):
                    window.iconbitmap(ico_path)
                    return
        except Exception:
            pass

    # ════════════════════════════════════════════════════════════
    #  管理员权限检查与提升
    # ════════════════════════════════════════════════════════════

    def _require_admin(self, operation="此操作"):
        if self._is_admin:
            return True
        msg = (
            f"「{operation}」需要管理员权限。\n\n"
            f"是否以管理员身份重新启动程序？\n"
            f"（当前窗口将关闭，操作状态将自动恢复）"
        )
        if messagebox.askyesno("需要管理员权限", msg, icon="warning"):
            self._restart_as_admin()
        return False

    def _save_state_for_elevation(self):
        """保存当前 UI 状态和检测结果到临时文件，供提权后恢复"""
        try:
            tab_idx = 0
            try:
                tab_idx = self.notebook.index(self.notebook.select())
            except Exception:
                pass
            state = {
                "file_path": self.file_path_var.get(),
                "drive": self.drive_var.get(),
                "tab": tab_idx,
                "show_def": self.show_def_var.get(),
                "timestamp": time.time(),
                # ★ 保存检测结果
                "file_lock_path": self._file_lock_path,
                "file_lock_processes": self._file_lock_processes,
                "file_lock_services": self._file_lock_services,
            }
            with open(STATE_FILE_PATH, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False)
        except Exception:
            pass

    def _restore_state_from_elevation(self):
        """从临时文件恢复提权前的 UI 状态和检测结果"""
        try:
            if not os.path.exists(STATE_FILE_PATH):
                return
            age = time.time() - os.path.getmtime(STATE_FILE_PATH)
            if age > 120:
                os.remove(STATE_FILE_PATH)
                return
            with open(STATE_FILE_PATH, "r", encoding="utf-8") as f:
                state = json.load(f)
            os.remove(STATE_FILE_PATH)

            restored = []

            if state.get("file_path"):
                self.file_path_var.set(state["file_path"])
                restored.append(f"文件路径: {state['file_path']}")

            if state.get("show_def"):
                self.show_def_var.set(True)
                self.drive_frame.config(
                    text="盘符选择（D: 及之后 ⚠ 含本地硬盘）")
                restored.append("D/E/F 盘显示: 已启用")

            if state.get("drive"):
                drive_prefix = state["drive"].split()[0] if state["drive"] else ""
                if drive_prefix:
                    self.drive_var.set(state["drive"])
                    restored.append(f"盘符: {drive_prefix}")

            if state.get("tab") is not None:
                try:
                    self.notebook.select(state["tab"])
                except Exception:
                    pass
                tab_names = ["解除磁盘占用/弹出", "文件/文件夹占用", "进阶功能"]
                idx = state["tab"]
                if 0 <= idx < len(tab_names):
                    restored.append(f"标签页: {tab_names[idx]}")

            # ★ 恢复检测结果
            has_detection = False
            if state.get("file_lock_path"):
                self._file_lock_path = state["file_lock_path"]
            if state.get("file_lock_processes"):
                self._file_lock_processes = state["file_lock_processes"]
                has_detection = True
            if state.get("file_lock_services"):
                self._file_lock_services = state["file_lock_services"]
                has_detection = True

            if restored:
                self.log_msg("[恢复] 已从提权前恢复以下状态：")
                for item in restored:
                    self.log_msg(f"  • {item}")

            # ★ 如有检测结果，显示恢复信息并弹出引导对话框
            if has_detection:
                n_proc = len(self._file_lock_processes)
                n_svc = len(self._file_lock_services)
                self._detection_is_restored = True
                self.log_msg("")
                self.log_msg(f"[恢复] 已恢复提权前的检测结果：")
                self.log_msg(f"  检测路径: {self._file_lock_path}")
                self.log_msg(f"  占用进程: {n_proc} 个")
                if self._file_lock_processes:
                    for p in self._file_lock_processes:
                        self.log_msg(
                            f"    PID={p['pid']:<6}  {p['name']:<24}"
                            f"  {p.get('detail','')}")
                self.log_msg(f"  占用服务: {n_svc} 个")
                if self._file_lock_services:
                    for name, display in self._file_lock_services.items():
                        self.log_msg(f"    {display} ({name})")
                self.log_msg("")

                # ★ 验证进程是否仍存活
                alive_procs = []
                dead_procs = []
                for p in self._file_lock_processes:
                    if is_process_alive(p["pid"]):
                        alive_procs.append(p)
                    else:
                        dead_procs.append(p)

                if dead_procs:
                    self.log_msg(f"[验证] 以下 {len(dead_procs)} 个进程已不存在"
                                 f"（提权期间退出）：")
                    for p in dead_procs:
                        self.log_msg(f"    PID={p['pid']}  {p['name']}  →  已退出")
                    self._file_lock_processes = alive_procs
                    self.log_msg(f"[验证] 仍存活: {len(alive_procs)} 个进程\n")

                # ★ 弹出引导对话框
                self.root.after(
                    400,
                    lambda: self._show_restored_detection_dialog(
                        len(alive_procs),
                        len(self._file_lock_services)))
            else:
                self.log_msg("")

        except Exception:
            pass
        finally:
            try:
                if os.path.exists(STATE_FILE_PATH):
                    os.remove(STATE_FILE_PATH)
            except Exception:
                pass

    def _show_restored_detection_dialog(self, n_proc, n_svc):
        """提权后弹出引导对话框：选择重新检测 / 直接解除 / 稍后决定"""
        if n_proc == 0 and n_svc == 0:
            self.log_msg("[提示] 提权前的检测结果已全部失效，请重新检测。\n")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("提权完成 — 检测结果已恢复")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()
        self._set_icon(dlg)

        ttk.Label(
            dlg,
            text="✓ 已成功提升为管理员权限",
            font=("Microsoft YaHei UI", 12, "bold"),
            foreground="#1a7f1a",
        ).pack(padx=24, pady=(18, 6))

        path_text = self._file_lock_path if self._file_lock_path else "(未知)"
        info_text = (
            f"提权前的检测结果已恢复：\n\n"
            f"  检测路径：{path_text}\n"
            f"  占用进程：{n_proc} 个（已验证仍存活）\n"
            f"  占用服务：{n_svc} 个\n"
        )
        ttk.Label(
            dlg, text=info_text,
            font=("Microsoft YaHei UI", 10),
            justify="left",
        ).pack(padx=24, pady=(4, 6))

        ttk.Label(
            dlg,
            text=(
                "请选择下一步操作：\n\n"
                "• 重新检测（推荐）— 确保结果最新最准确\n"
                "• 直接解除占用 — 使用已恢复的结果立即解除\n"
                "• 稍后决定 — 关闭对话框，自行操作"
            ),
            font=("Microsoft YaHei UI", 10),
            justify="left",
            wraplength=440,
        ).pack(padx=24, pady=(0, 12))

        btn_frame = ttk.Frame(dlg, padding=8)
        btn_frame.pack(fill="x", padx=24, pady=(0, 18))

        def do_redetect():
            dlg.destroy()
            self._detection_is_restored = False
            self.detect_file_lock()

        def do_kill_now():
            dlg.destroy()
            self._detection_is_restored = False
            self.kill_all_file_lock()

        def do_later():
            dlg.destroy()

        # ★ 推荐按钮
        redetect_btn = ttk.Button(
            btn_frame, text="🔍 重新检测（推荐）",
            command=do_redetect)
        redetect_btn.pack(side="left", padx=(0, 6))

        ttk.Button(
            btn_frame, text="⚡ 直接解除占用",
            command=do_kill_now,
        ).pack(side="left", padx=(0, 6))

        ttk.Button(
            btn_frame, text="稍后决定",
            command=do_later,
        ).pack(side="right")

        dlg.protocol("WM_DELETE_WINDOW", do_later)

        dlg.update_idletasks()
        dlg_w = dlg.winfo_width()
        dlg_h = dlg.winfo_height()
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()
        x = root_x + (root_w - dlg_w) // 2
        y = root_y + (root_h - dlg_h) // 2
        dlg.geometry(f"+{max(0, x)}+{max(0, y)}")

    def _cleanup_state_file(self):
        try:
            if os.path.exists(STATE_FILE_PATH):
                os.remove(STATE_FILE_PATH)
        except Exception:
            pass

    def _restart_as_admin(self):
        self._save_state_for_elevation()
        try:
            if getattr(sys, 'frozen', False):
                exe = sys.executable
                args = ""
            else:
                exe = sys.executable
                script = os.path.abspath(sys.argv[0])
                args = f'"{script}"'
            ctypes.windll.shell32.ShellExecuteW.restype = ctypes.c_long
            ret = ctypes.windll.shell32.ShellExecuteW(
                None, "runas", exe, args, None, 1
            )
            if ret > 32:
                self.root.destroy()
                sys.exit(0)
            else:
                self._cleanup_state_file()
                self.log_msg("[提示] 用户取消了权限提升，或提升失败")
        except Exception as e:
            self._cleanup_state_file()
            self.log_msg(f"[错误] 提升权限失败: {e}")

    def _request_admin_elevation(self):
        msg = (
            "将以管理员身份重新启动程序。\n"
            "当前窗口将关闭，操作状态（含检测结果）将自动恢复。\n\n"
            "是否继续？"
        )
        if messagebox.askyesno("提升为管理员", msg):
            self._restart_as_admin()

    # ════════════════════════════════════════════════════════════
    #  ★ 新增：管理员模式下放行 UIPI 拖放消息限制
    # ════════════════════════════════════════════════════════════

    def _allow_drag_drop_admin(self):
        """
        管理员模式下，Windows UIPI 会阻止普通权限进程（如资源管理器）
        向高权限窗口发送拖放消息。此方法通过 ChangeWindowMessageFilter /
        ChangeWindowMessageFilterEx 放行相关消息，使拖放在管理员窗口中也能工作。
        """
        if not self._is_admin:
            return

        MSGFLT_ALLOW        = 1
        WM_DROPFILES        = 0x0233
        WM_COPYDATA         = 0x004A
        WM_COPYGLOBALDATA   = 0x0049

        user32 = ctypes.windll.user32

        # 方法 1：全局消息过滤器（进程级，Vista+）
        try:
            _filter = user32.ChangeWindowMessageFilter
            _filter.argtypes = [wintypes.UINT, wintypes.DWORD]
            _filter.restype = wintypes.BOOL
            _filter(WM_DROPFILES, MSGFLT_ALLOW)
            _filter(WM_COPYDATA, MSGFLT_ALLOW)
            _filter(WM_COPYGLOBALDATA, MSGFLT_ALLOW)
        except Exception:
            pass

        # 方法 2：窗口级过滤器（Win7+），对 windnd 使用的 HWND 精确放行
        try:
            self.root.update_idletasks()
            hwnd = self.root.winfo_id()

            _filterex = user32.ChangeWindowMessageFilterEx
            _filterex.argtypes = [
                wintypes.HWND, wintypes.UINT, wintypes.DWORD, ctypes.c_void_p
            ]
            _filterex.restype = wintypes.BOOL
            _filterex(hwnd, WM_DROPFILES, MSGFLT_ALLOW, None)
            _filterex(hwnd, WM_COPYDATA, MSGFLT_ALLOW, None)
            _filterex(hwnd, WM_COPYGLOBALDATA, MSGFLT_ALLOW, None)

            # 同时对顶层窗口 HWND 放行（tkinter 内部 HWND 与顶层可能不同）
            try:
                top_hwnd = int(self.root.frame(), 16)
                if top_hwnd != hwnd:
                    _filterex(top_hwnd, WM_DROPFILES, MSGFLT_ALLOW, None)
                    _filterex(top_hwnd, WM_COPYDATA, MSGFLT_ALLOW, None)
                    _filterex(top_hwnd, WM_COPYGLOBALDATA, MSGFLT_ALLOW, None)
            except Exception:
                pass
        except Exception:
            pass

    # ════════════════════════════════════════════════════════════

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

        # ── 顶部栏 ──
        top_row = ttk.Frame(m)
        top_row.pack(fill="x", pady=(0, 2))

        ver_lbl = ttk.Label(
            top_row, text=f"v{APP_VERSION}",
            foreground="#909090",
            font=("Consolas", 10),
        )
        ver_lbl.pack(side="left")

        if self._is_admin:
            ttk.Label(
                top_row, text="✓ 管理员模式",
                foreground="#1a7f1a",
                font=("Microsoft YaHei UI", 9),
            ).pack(side="right")
        else:
            ttk.Button(
                top_row, text="⬆ 提升为管理员",
                command=self._request_admin_elevation,
            ).pack(side="right")
            ttk.Label(
                top_row, text="普通模式  ",
                foreground="#c07000",
                font=("Microsoft YaHei UI", 9),
            ).pack(side="right")

        # ── 盘符选择 ──
        self.drive_frame = ttk.LabelFrame(
            m, text="盘符选择（仅 G: 及之后）", padding=8
        )
        self.drive_frame.pack(fill="x", pady=(0, 4))

        row1 = ttk.Frame(self.drive_frame)
        row1.pack(fill="x")

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
        self.notebook = nb
        gk = dict(sticky="nsew", padx=3, pady=3, ipady=2)

        # ══════════ Tab 1: 解除磁盘占用 / 弹出 ══════════
        t1 = ttk.Frame(nb, padding=6)
        nb.add(t1, text=" 解除磁盘占用 / 弹出 ")
        t1.columnconfigure(0, weight=1)
        t1.columnconfigure(1, weight=1)

        ttk.Button(t1, text="检测占用进程和服务",
                   command=self.detect).grid(row=0, column=0, **gk)
        ttk.Button(t1, text="一键停止占用服务",
                   command=self.stop_svc).grid(row=0, column=1, **gk)
        ttk.Button(t1, text="恢复已停止的服务",
                   command=lambda: self._show_service_restore_dialog(
                       auto_popup=False)
                   ).grid(row=1, column=0, **gk)
        ttk.Button(t1, text="恢复脱机磁盘",
                   command=self.recover_offline).grid(row=1, column=1, **gk)
        ttk.Separator(t1).grid(row=2, column=0, columnspan=2, sticky="ew", pady=4)

        rec_btn = self._make_rec_btn(t1, self.smart_eject)
        rec_btn.grid(row=3, column=0, sticky="nsew", padx=3, pady=3)
        ttk.Button(t1, text="强制弹出\n直接弹出硬盘",
                   command=self.force_eject).grid(row=3, column=1, **gk)

        # ══════════ Tab 2: 文件/文件夹占用检测 ══════════
        t2 = ttk.Frame(nb, padding=6)
        nb.add(t2, text=" 文件/文件夹占用 ")

        path_row = ttk.Frame(t2)
        path_row.pack(fill="x", pady=(0, 6))
        ttk.Label(path_row, text="路径:").pack(side="left", padx=(0, 4))
        self.file_path_var = tk.StringVar()
        self.file_path_entry = ttk.Entry(
            path_row, textvariable=self.file_path_var,
            font=("Consolas", 10),
        )
        self.file_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 4))
        ttk.Button(
            path_row, text="文件", width=5,
            command=self.browse_file,
        ).pack(side="left", padx=1)
        ttk.Button(
            path_row, text="文件夹", width=6,
            command=self.browse_folder,
        ).pack(side="left", padx=1)

        t2_btns = ttk.Frame(t2)
        t2_btns.pack(fill="x", pady=(0, 4))
        t2_btns.columnconfigure(0, weight=1)
        t2_btns.columnconfigure(1, weight=1)
        t2_btns.columnconfigure(2, weight=1)
        ttk.Button(
            t2_btns, text="检测占用",
            command=self.detect_file_lock,
        ).grid(row=0, column=0, sticky="ew", padx=3, pady=2)
        ttk.Button(
            t2_btns, text="一键停止所有占用",
            command=self.kill_all_file_lock,
        ).grid(row=0, column=1, sticky="ew", padx=3, pady=2)
        ttk.Button(
            t2_btns, text="恢复已停止的服务",
            command=lambda: self._show_service_restore_dialog(
                auto_popup=False),
        ).grid(row=0, column=2, sticky="ew", padx=3, pady=2)

        # ★ 修改：拖放提示区分管理员 / 普通模式
        if HAS_WINDND:
            try:
                self._allow_drag_drop_admin()           # ★ 新增：先放行 UIPI
                windnd.hook_dropfiles(self.root, func=self._on_file_drop)
                self._dnd_ok = True
            except Exception:
                self._dnd_ok = False

            if self._dnd_ok and self._is_admin:
                dnd_hint = ("✓ 拖放已启用（管理员模式，已自动放行 UIPI；"
                            "如拖放仍失效请用浏览按钮）")
            elif self._dnd_ok:
                dnd_hint = "✓ 拖放已启用 — 可将文件或文件夹直接拖入窗口"
            else:
                dnd_hint = "⚠ 拖放初始化失败，请使用浏览按钮选择路径"
        else:
            dnd_hint = ("提示：安装 windnd（pip install windnd）可启用"
                        "拖放功能，也可手动粘贴路径")

        ttk.Label(t2, foreground="gray", text=dnd_hint,
                  wraplength=600).pack(anchor="w", pady=(2, 0))

        # ══════════ Tab 3: 进阶功能 ══════════
        t3 = ttk.Frame(nb, padding=6)
        nb.add(t3, text=" 进阶功能 ")

        del_frame = ttk.LabelFrame(t3, text="删除系统文件夹", padding=8)
        del_frame.pack(fill="x", pady=(0, 8))
        del_frame.columnconfigure(0, weight=1)
        del_frame.columnconfigure(1, weight=1)

        ttk.Button(del_frame, text="删除\nSystem Volume Information",
                   command=self.del_svi).grid(row=0, column=0, **gk)
        ttk.Button(del_frame, text="删除\n$RECYCLE.BIN",
                   command=self.del_rec).grid(row=0, column=1, **gk)
        ttk.Button(del_frame, text="一键删除以上两个文件夹",
                   command=self.del_both).grid(
                       row=1, column=0, columnspan=2, **gk)

        perm_frame = ttk.LabelFrame(t3, text="SYSTEM 写入权限", padding=8)
        perm_frame.pack(fill="x")
        perm_frame.columnconfigure(0, weight=1)
        perm_frame.columnconfigure(1, weight=1)

        ttk.Button(perm_frame, text="禁止 SYSTEM 写入",
                   command=self.deny_write).grid(row=0, column=0, **gk)
        ttk.Button(perm_frame, text="恢复 SYSTEM 写入",
                   command=self.allow_write).grid(row=0, column=1, **gk)
        ttk.Label(
            perm_frame, foreground="gray",
            text="提示：禁止写入后，系统服务将无法在该盘创建任何文件。"
                 "\n如需恢复，请在拔盘前点击【恢复】按钮。"
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        self._tab_frames = [t1, t2, t3]

        nb.pack(fill="x", pady=(4, 0))
        nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # ── 日志区域 ──
        f4 = ttk.LabelFrame(m, text="执行日志", padding=4)
        f4.pack(fill="both", expand=True, pady=(4, 0))

        btn_bar = ttk.Frame(f4)
        btn_bar.pack(side="bottom", fill="x", pady=(2, 0))
        ttk.Button(
            btn_bar, text="清空日志",
            command=lambda: self.log.delete("1.0", tk.END)
        ).pack(anchor="e")

        self.log = scrolledtext.ScrolledText(
            f4, height=1, font=("Consolas", 10), wrap=tk.WORD
        )
        self.log.pack(fill="both", expand=True)

        self.root.after(150, self._resize_notebook_to_current)

        mode_text = "管理员模式" if self._is_admin else "普通模式"
        self.log_msg(f"[OK] 工具已启动（{mode_text}）")
        if not self._is_admin:
            self.log_msg("[提示] 当前为普通模式，部分功能需要管理员权限")
            self.log_msg("[提示] 可点击右上角「提升为管理员」按钮获取完整功能")
        ds = ", ".join(f"{d[0]}[{d[1]}]" for d in drives) if drives else "无"
        self.log_msg(f"检测到盘符：{ds}")
        self.log_msg("正在后台识别磁盘总线类型...")

        # ★ 修改：拖放启动日志区分管理员模式
#        if HAS_WINDND:
#            if self._dnd_ok:
#                self.log_msg("[拖放] windnd 已加载，支持拖放文件/文件夹到窗口")
#                if self._is_admin:
#                    self.log_msg("[拖放] 管理员模式：已放行 UIPI 拖放消息"
#                                 "（WM_DROPFILES / WM_COPYGLOBALDATA）")
#                    self.log_msg("[拖放] 如拖放仍不可用，"
#                                 "请使用「文件」或「文件夹」浏览按钮")
#            else:
#                self.log_msg("[拖放] windnd 已安装但初始化失败，请使用浏览按钮")
#        else:
#            self.log_msg("[拖放] 未安装 windnd，可用 pip install windnd 启用拖放")
#
#        self.log_msg("")

    # ════════════════════════════════════════════════════════════
    #  Notebook 动态高度
    # ════════════════════════════════════════════════════════════

    def _on_tab_changed(self, event=None):
        self.root.after(20, self._resize_notebook_to_current)

    def _resize_notebook_to_current(self):
        try:
            current = self.notebook.select()
            if not current:
                return
            tab_frame = self.notebook.nametowidget(current)
            tab_frame.update_idletasks()
            needed = tab_frame.winfo_reqheight()
            self.notebook.configure(height=needed)
        except Exception:
            pass

    # ════════════════════════════════════════════════════════════
    #  服务恢复对话框
    # ════════════════════════════════════════════════════════════

    def _show_service_restore_dialog(self, auto_popup=False):
        to_restore = dict(self._all_stopped_services)

        if not to_restore:
            if not auto_popup:
                messagebox.showinfo("提示", "没有需要恢复的服务。")
            return

        if not auto_popup and not self._is_admin:
            if not self._require_admin("恢复服务"):
                return

        dlg = tk.Toplevel(self.root)
        dlg.title("恢复已停止的服务")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()
        self._set_icon(dlg)

        if auto_popup:
            msg_text = ("以下服务已被停止。完成磁盘操作后建议恢复。\n"
                        "请勾选要恢复的服务，或点击「稍后恢复」：")
        else:
            msg_text = "以下服务处于停止状态，请勾选要恢复的服务："

        ttk.Label(
            dlg, text=msg_text,
            font=("Microsoft YaHei UI", 11),
            wraplength=460,
        ).pack(padx=20, pady=(16, 10))

        cb_frame = ttk.Frame(dlg, padding=(8, 0))
        cb_frame.pack(fill="x", padx=20)

        check_vars = {}

        for name, display in to_restore.items():
            rec_level, rec_text = SERVICE_RECOMMENDATIONS.get(
                name, ("", ""))

            var = tk.BooleanVar(value=True)
            check_vars[name] = var

            row = ttk.Frame(cb_frame)
            row.pack(fill="x", pady=(4, 0))

            cb = ttk.Checkbutton(
                row, text=f"{display} ({name})", variable=var,
                style="TCheckbutton",
            )
            cb.pack(anchor="w")

            if rec_level:
                if rec_level == "建议恢复":
                    hint_color = "#1a7f1a"
                    prefix = "→ 建议恢复"
                else:
                    hint_color = "#888888"
                    prefix = "→ 可选"
                hint_text = f"    {prefix} — {rec_text}"
                ttk.Label(
                    row, text=hint_text, foreground=hint_color,
                    font=("Microsoft YaHei UI", 9),
                ).pack(anchor="w", padx=(20, 0))

        btn_frame = ttk.Frame(dlg, padding=8)
        btn_frame.pack(fill="x", padx=20, pady=(10, 16))

        def select_all():
            for v in check_vars.values():
                v.set(True)

        def select_none():
            for v in check_vars.values():
                v.set(False)

        def do_restore():
            selected = {
                n: to_restore[n]
                for n, v in check_vars.items() if v.get()
            }
            dlg.destroy()
            if selected:
                self.run_in_thread(
                    lambda: self._do_restore_selected(selected))

        def do_later():
            dlg.destroy()

        ttk.Button(btn_frame, text="全选", command=select_all,
                   width=6).pack(side="left", padx=(0, 4))
        ttk.Button(btn_frame, text="全不选", command=select_none,
                   width=7).pack(side="left", padx=(0, 4))
        ttk.Button(btn_frame, text="恢复选中的服务",
                   command=do_restore).pack(side="right", padx=(4, 0))
        ttk.Button(btn_frame, text="稍后恢复",
                   command=do_later).pack(side="right", padx=(4, 0))

        dlg.protocol("WM_DELETE_WINDOW", do_later)

        dlg.update_idletasks()
        dlg_w = dlg.winfo_width()
        dlg_h = dlg.winfo_height()
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()
        x = root_x + (root_w - dlg_w) // 2
        y = root_y + (root_h - dlg_h) // 2
        dlg.geometry(f"+{max(0, x)}+{max(0, y)}")

    def _do_restore_selected(self, selected):
        self.log_msg("\n--- 恢复已停止的服务 ---")
        count = 0
        for name, display in selected.items():
            st = svc_status(name)
            if st == "stopped":
                self.log_msg(f"  启动 {display} ({name}) ...")
                self.exec_cmd(f"net start {name}", timeout=30)
                count += 1
                self._all_stopped_services.pop(name, None)
            elif st == "running":
                self.log_msg(f"  跳过 {display} - 已在运行")
                self._all_stopped_services.pop(name, None)
            else:
                self.log_msg(f"  跳过 {display} - 未安装")
                self._all_stopped_services.pop(name, None)
        self.log_msg(f"\n[OK] 共恢复 {count} 个服务\n")

    # ========== 后台总线类型检测 ==========

    def _start_bus_detection(self):
        self._detecting = True
        threading.Thread(target=self._do_bus_detection, daemon=True).start()

    def _do_bus_detection(self):
        min_letter = 'D' if self.show_def_var.get() else 'G'
        drives, bus_types = get_drives_full(min_letter)
        self._bus_cache = bus_types
        self.root.after(0, lambda: self._apply_bus_detection(drives))

    def _apply_bus_detection(self, drives):
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

    def _get_bus_type(self, letter):
        letter = letter.rstrip(":\\").upper()
        if letter in self._bus_cache:
            return self._bus_cache[letter]
        bus = get_bus_type_ioctl(letter)
        if bus == "Unknown":
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
        if not self._require_admin("恢复脱机磁盘"):
            return
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
        num = get_disk_number_ioctl(letter)
        if num is not None:
            return num
        rc, out, _ = run_cmd(
            f'powershell -NoProfile -Command "'
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
                f'powershell -NoProfile -ExecutionPolicy Bypass '
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

    def _flush_volume(self, d):
        letter = d.rstrip(":\\")
        volume = f"\\\\.\\{letter}:"
        k32 = ctypes.windll.kernel32
        k32.CreateFileW.restype = ctypes.c_void_p
        INVALID_HANDLE = ctypes.c_void_p(-1).value
        SHARE_RW = 0x3
        OPEN_EXISTING = 3

        h = k32.CreateFileW(volume, 0xC0000000, SHARE_RW, None, OPEN_EXISTING, 0, None)
        if h is None or h == INVALID_HANDLE:
            h = k32.CreateFileW(volume, 0x80000000, SHARE_RW, None, OPEN_EXISTING, 0, None)
        if h is None or h == INVALID_HANDLE:
            self.log_msg("    [注意] 无法打开卷句柄，跳过刷缓冲区")
            return
        ok = k32.FlushFileBuffers(h)
        k32.CloseHandle(h)
        if ok:
            self.log_msg("    [OK] 缓冲区已刷新")
        else:
            self.log_msg("    [注意] FlushFileBuffers 返回失败，缓冲区可能未完全刷新")

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
        ok, msg, warnings = eject_volume_api(d)
        if warnings:
            for w in warnings:
                self.log_msg(f"    [注意] {w}")
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
            'powershell -NoProfile -Command "'
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

        if self._is_admin:
            self.log_msg("\n  方法4: Set-Disk -IsOffline + USB 安全移除 ...")
            self.log_msg("    [注意] 此方法如果 USB 移除失败，下次插入可能需要手动恢复联机")
            self.log_msg("    刷新卷缓冲区...")
            self._flush_volume(d)
            cmd = (
                'powershell -NoProfile -Command "'
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
                        self.log_msg("    [!] 数据已安全（文件系统已脱机），但硬盘可能仍在转动")
                        self.log_msg("    [!] 建议等待约 5 秒再拔出 USB 线缆")
                        self.log_msg("    [!] 下次插入如无盘符，请点击【恢复脱机磁盘】")
                return True
            self.log_msg("    盘符仍在，尝试下一方法...")
        else:
            self.log_msg("\n  方法4: 跳过（需要管理员权限）")

        if self._is_admin:
            self.log_msg("\n  方法5: diskpart ...")
            self.log_msg("    刷新卷缓冲区...")
            self._flush_volume(d)
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
        else:
            self.log_msg("\n  方法5: 跳过（需要管理员权限）")

        return False

    # ========== 检测（整盘） ==========

    def detect(self):
        d = self.get_drive()
        if not d:
            return
        self.run_in_thread(lambda: self._detect(d))

    def _detect(self, d):
        self.log_msg(f"\n{'='*50}")
        self.log_msg(f"  检测占用 {d}\\ 的进程和服务")
        self.log_msg(f"{'='*50}")

        if not self._is_admin:
            self.log_msg("  [注意] 当前为普通模式，部分检测功能可能受限")

        bus = self._get_bus_type(d[0])
        dt = get_drive_type_code(d)
        type_name = DRIVE_TYPE_MAP.get(dt, "未知")
        self.log_msg(f"\n  磁盘类型: {type_name}  总线: {bus}")

        self.log_msg("\n[1] 在该盘上运行的进程：")
        cmd = (
            'powershell -NoProfile -Command "'
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
            'powershell -NoProfile -Command "Get-Process | ForEach-Object { $p=$_; try { '
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
            if not self._is_admin:
                self.log_msg("    不可用（需要管理员权限）")
            else:
                self.log_msg("    不可用（需先运行 openfiles /local on 并重启）")
        self.log_msg("")

    # ========== 服务管理 ==========

    def stop_svc(self):
        if not self._require_admin("停止系统服务"):
            return
        self.run_in_thread(self._stop_svc)

    def _stop_svc(self):
        self.log_msg("\n--- 停止占用服务 ---")
        stopped_now = {}
        for name, display in SERVICES:
            st = svc_status(name)
            if st == "running":
                self.log_msg(f"  停止 {display} ({name}) ...")
                self.exec_cmd(f"net stop {name} /y", timeout=30)
                if svc_status(name) == "stopped":
                    stopped_now[name] = display
                    self._all_stopped_services[name] = display
            elif st == "stopped":
                self.log_msg(f"  跳过 {display} - 已是停止状态")
            else:
                self.log_msg(f"  跳过 {display} - 未安装")
        count = len(stopped_now)
        self.log_msg(f"\n[OK] 共停止 {count} 个服务\n")

        if stopped_now:
            self.root.after(
                0, lambda: self._show_service_restore_dialog(auto_popup=True))

    # ========== 弹出 ==========

    def smart_eject(self):
        d = self.get_drive()
        if not d:
            return
        if not self._require_admin("安全弹出（需停止服务）"):
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
        smart_stopped = {}
        for name, display in SERVICES:
            st = svc_status(name)
            if st == "running":
                self.log_msg(f"  停止 {display} ...")
                run_cmd(f"net stop {name} /y", timeout=20)
                smart_stopped[name] = display
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
        all_to_restore = dict(smart_stopped)
        all_to_restore.update(self._all_stopped_services)
        for name, display in SERVICES:
            if name in all_to_restore:
                run_cmd(f"net start {name}", timeout=10)
                self.log_msg(f"  已恢复 {display}")
                self._all_stopped_services.pop(name, None)
        self.log_msg("[OK] 服务恢复完成\n")

        if ok:
            self.log_msg("正在刷新盘符列表...")
            time.sleep(1)
            self.root.after(0, self.refresh)

    def force_eject(self):
        d = self.get_drive()
        if not d:
            return
        if not self._check_drive_safety(d):
            self.log_msg(f"[取消] 用户取消了对 {d} 的弹出操作\n")
            return
        if not self._is_admin:
            msg = (
                f"当前为普通模式，部分弹出方法（diskpart、Set-Disk 等）\n"
                f"将不可用，但仍可尝试 USB 安全移除等方法。\n\n"
                f"尝试弹出 {d}？\n\n"
                f"提示：如需完整功能，请点击右上角「提升为管理员」。"
            )
        else:
            msg = f"跳过停止服务，直接弹出 {d}？"
        if not messagebox.askyesno("强制弹出", msg):
            return
        self.run_in_thread(lambda: self._force_eject(d))

    def _force_eject(self, d):
        self.log_msg(f"\n--- 强制弹出 {d} ---")
        if not self._is_admin:
            self.log_msg("  [注意] 当前为普通模式，方法4/5（需管理员）将被跳过")
        ok = self._try_eject(d)
        if ok:
            self.log_msg(f"\n[OK] {d} 已弹出！可以安全拔出硬盘。\n")
            self.log_msg("正在刷新盘符列表...")
            time.sleep(1)
            self.root.after(0, self.refresh)
        else:
            self.log_msg(f"\n[!!] {d} 仍然存在，弹出失败。")
            if not self._is_admin:
                self.log_msg("     建议以管理员身份运行后重试。")
            else:
                self.log_msg("     请关闭占用该盘的程序后重试。")
            self.log_msg("")

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
        if not self._require_admin("删除系统文件夹"):
            return
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
        if not self._require_admin("删除回收站"):
            return
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
        if not self._require_admin("删除系统文件夹"):
            return
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
        if not self._require_admin("修改磁盘权限"):
            return
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
        if not self._require_admin("修改磁盘权限"):
            return
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

    # ════════════════════════════════════════════════════════════
    #  文件/文件夹占用检测与结束
    # ════════════════════════════════════════════════════════════

    def _on_file_drop(self, files):
        if not files or self._busy:
            return
        path = files[0]
        if isinstance(path, bytes):
            try:
                path = path.decode('utf-8')
            except UnicodeDecodeError:
                path = path.decode('gbk', errors='replace')
        path = os.path.normpath(path.strip().strip('"'))
        self.file_path_var.set(path)
        try:
            self.notebook.select(1)
        except Exception:
            pass
        self.log_msg(f"[拖放] 已接收路径: {path}")

    def browse_file(self):
        path = filedialog.askopenfilename(title="选择要检测占用的文件")
        if path:
            self.file_path_var.set(os.path.normpath(path))

    def browse_folder(self):
        path = filedialog.askdirectory(title="选择要检测占用的文件夹")
        if path:
            self.file_path_var.set(os.path.normpath(path))

    @staticmethod
    def _drive_letter_of(path):
        normed = os.path.normpath(path)
        if len(normed) >= 2 and normed[1] == ':':
            return normed[0].upper()
        return None

    def detect_file_lock(self):
        path = self.file_path_var.get().strip().strip('"')
        if not path:
            messagebox.showwarning("提示", "请先输入或选择一个文件/文件夹路径")
            return
        path = os.path.normpath(path)
        self.file_path_var.set(path)
        if not os.path.exists(path):
            messagebox.showwarning("提示", f"路径不存在:\n{path}")
            return
        self._detection_is_restored = False
        self.run_in_thread(lambda: self._do_detect_file_lock(path))

    def _do_detect_file_lock(self, path):
        is_dir = os.path.isdir(path)
        self._file_lock_path = path
        self.log_msg(f"\n{'='*50}")
        self.log_msg(f"  检测占用: {path}")
        self.log_msg(f"  类型: {'文件夹' if is_dir else '文件'}")
        self.log_msg(f"{'='*50}")

        all_procs = {}

        # ══════════════════════════════════════════════════════
        # [1] Restart Manager API 检测
        # ══════════════════════════════════════════════════════
        self.log_msg("\n[1] Restart Manager API 检测:")
        if is_dir:
            self.log_msg("    收集目录内文件（最多300个，深度3层）...")
            files = collect_files_in_dir(path, max_files=300, max_depth=3)
            self.log_msg(f"    收集到 {len(files)} 个文件")
            if files:
                for batch_start in range(0, len(files), 100):
                    batch = files[batch_start:batch_start + 100]
                    rm_results = find_locking_processes_rm(batch)
                    for r in rm_results:
                        if r["pid"] not in all_procs:
                            svc = (f" (服务: {r['service']})"
                                   if r["service"] else "")
                            all_procs[r["pid"]] = {
                                "pid": r["pid"],
                                "name": r["name"],
                                "detail": f"文件占用{svc}",
                            }
        else:
            rm_results = find_locking_processes_rm(path)
            for r in rm_results:
                svc = (f" (服务: {r['service']})"
                       if r["service"] else "")
                all_procs[r["pid"]] = {
                    "pid": r["pid"],
                    "name": r["name"],
                    "detail": f"文件占用{svc}",
                }

        if all_procs:
            for info in all_procs.values():
                self.log_msg(
                    f"    PID={info['pid']:<6}  "
                    f"{info['name']:<24}  {info['detail']}"
                )
        else:
            self.log_msg("    （未检测到）")

        # ══════════════════════════════════════════════════════
        # [2] 进程路径 / 模块匹配检测
        # ══════════════════════════════════════════════════════
        if is_dir:
            self.log_msg("\n[2] PowerShell 进程路径/模块检测:")
            escaped = path.replace("'", "''").rstrip('\\')
            ps_cmd = (
                'powershell -NoProfile -Command "Get-Process | ForEach-Object { '
                '$p=$_; $found=$false; '
                "if ($p.Path -and ($p.Path -like '"
                + escaped
                + "\\*')) { $found=$true }; "
                'if (-not $found) { try { '
                "$p.Modules | ForEach-Object { "
                "if ($_.FileName -like '"
                + escaped
                + "\\*') { $found=$true } } "
                '} catch {} }; '
                'if ($found) { '
                'Write-Output ("{0}|{1}|{2}" -f $p.Id,$p.Name,$p.Path) '
                '} }"'
            )
            rc, out, _ = run_cmd(ps_cmd, timeout=20)
            ps_found = 0
            if out.strip():
                for line in out.strip().splitlines():
                    parts = line.strip().split('|', 2)
                    if len(parts) >= 2:
                        try:
                            pid = int(parts[0])
                            pname = parts[1]
                            ppath = parts[2] if len(parts) > 2 else ""
                            ps_found += 1
                            self.log_msg(
                                f"    PID={pid:<6}  {pname:<24}  {ppath}"
                            )
                            if pid not in all_procs:
                                all_procs[pid] = {
                                    "pid": pid,
                                    "name": pname,
                                    "detail": f"进程路径: {ppath}",
                                }
                        except ValueError:
                            pass
            if not ps_found:
                self.log_msg("    （未检测到）")
        else:
            self.log_msg("\n[2] 进程可执行文件/模块匹配检测:")
            escaped = path.replace("'", "''")
            ps_cmd = (
                'powershell -NoProfile -Command "Get-Process | ForEach-Object { '
                "$p=$_; $tag=$null; "
                "if ($p.Path -and ($p.Path -eq '"
                + escaped
                + "')) { $tag='EXE' }; "
                "if (-not $tag) { try { "
                "$p.Modules | ForEach-Object { "
                "if ($_.FileName -eq '"
                + escaped
                + "') { $tag='MOD' } } "
                "} catch {} }; "
                "if ($tag) { "
                "Write-Output ('{0}|{1}|{2}|{3}' -f "
                "$tag,$p.Id,$p.Name,$p.Path) "
                '} }"'
            )
            rc, out, _ = run_cmd(ps_cmd, timeout=20)
            ps_found = 0
            if rc == 0 and out.strip():
                for line in out.strip().splitlines():
                    parts = line.strip().split('|', 3)
                    if len(parts) >= 3:
                        try:
                            tag = parts[0]
                            pid = int(parts[1])
                            pname = parts[2]
                            ppath = parts[3] if len(parts) > 3 else ""
                            ps_found += 1
                            if tag == "EXE":
                                detail = "★ 该文件正在作为进程运行"
                            else:
                                detail = "该文件被加载为模块"
                            self.log_msg(
                                f"    PID={pid:<6}  {pname:<24}  {detail}"
                            )
                            if pid not in all_procs:
                                all_procs[pid] = {
                                    "pid": pid,
                                    "name": pname,
                                    "detail": detail,
                                }
                        except ValueError:
                            pass
            if not ps_found:
                self.log_msg("    （未检测到进程匹配）")

        # ══════════════════════════════════════════════════════
        # [3] 常见占用服务状态
        # ══════════════════════════════════════════════════════
        drive_letter = self._drive_letter_of(path)
        self.log_msg("\n[3] 常见占用服务状态：")
        if drive_letter:
            self.log_msg(f"    （目标路径所在盘符: {drive_letter}:）")
        self._file_lock_services = {}
        for name, display in SERVICES:
            st = svc_status(name)
            if st == "running":
                self.log_msg(f"    * {display} ({name})  ->  运行中 !!")
                self._file_lock_services[name] = display
            elif st == "stopped":
                self.log_msg(f"    o {display} ({name})  ->  已停止")
            else:
                self.log_msg(f"    - {display} ({name})  ->  未安装")

        # ══════════════════════════════════════════════════════
        # [4] openfiles 查询
        # ══════════════════════════════════════════════════════
        if drive_letter:
            self.log_msg(f"\n[4] openfiles 查询 ({drive_letter}:):")
            rc, out, err = run_cmd("openfiles /query /fo csv", timeout=10)
            if rc == 0 and out.strip():
                hits = [l for l in out.splitlines()
                        if f"{drive_letter}:" in l.upper()
                        or f"{drive_letter}:\\" in l.upper()]
                if hits:
                    for h in hits[:30]:
                        self.log_msg(f"    {h}")
                else:
                    self.log_msg("    未查到占用记录")
            else:
                if not self._is_admin:
                    self.log_msg("    不可用（需要管理员权限）")
                else:
                    self.log_msg(
                        "    不可用（需先运行 openfiles /local on 并重启）"
                    )

        self._file_lock_processes = list(all_procs.values())

        total_issues = (len(self._file_lock_processes)
                        + len(self._file_lock_services))
        if total_issues > 0:
            self.log_msg(
                f"\n[结果] 检测到 {len(self._file_lock_processes)} 个占用进程"
                f"，{len(self._file_lock_services)} 个运行中的占用服务")
            self.log_msg(
                "  → 点击【一键停止所有占用】可停止服务 + 结束进程")
            self.log_msg("  → 事后可点击【恢复已停止的服务】恢复\n")
        else:
            self.log_msg("\n[结果] 未检测到占用进程和运行中的占用服务")
            self.log_msg("  可能原因：")
            self.log_msg("  • 系统内核/驱动级占用（如杀毒软件实时防护）")
            self.log_msg("  • 句柄已关闭但 Windows 仍缓存引用（稍后重试）")
            self.log_msg(
                "  • 深层子目录中的文件被占用（当前最多扫描3层）\n")

    def kill_all_file_lock(self):
        has_proc = bool(self._file_lock_processes)
        has_svc = bool(self._file_lock_services)

        if not has_proc and not has_svc:
            messagebox.showinfo(
                "提示",
                "没有检测到占用进程或运行中的服务。\n请先点击【检测占用】。"
            )
            return

        # ★ 如果检测结果来自提权前，先验证并提示
        if self._detection_is_restored:
            alive = [p for p in self._file_lock_processes
                     if is_process_alive(p["pid"])]
            dead_count = len(self._file_lock_processes) - len(alive)
            self._file_lock_processes = alive
            has_proc = bool(alive)

            # 重新检查服务状态
            still_running = {}
            for name, display in self._file_lock_services.items():
                if svc_status(name) == "running":
                    still_running[name] = display
            self._file_lock_services = still_running
            has_svc = bool(still_running)

            if not has_proc and not has_svc:
                messagebox.showinfo(
                    "提示",
                    "提权前检测到的占用已全部失效（进程已退出、服务已停止）。\n\n"
                    "如仍有问题，请重新点击【检测占用】。"
                )
                self._detection_is_restored = False
                return

        if has_svc and not self._require_admin("停止占用服务并结束进程"):
            return

        lines = []

        # ★ 如果是恢复的结果，给出明确提示
        if self._detection_is_restored:
            lines.append("⚠ 以下为提权前的检测结果（已验证仍有效）：\n")

        if has_svc:
            lines.append(f"将停止 {len(self._file_lock_services)} 个服务：")
            for name, display in self._file_lock_services.items():
                lines.append(f"  ● {display} ({name})")
            lines.append("")
        if has_proc:
            lines.append(f"将结束 {len(self._file_lock_processes)} 个进程：")
            for p in self._file_lock_processes[:15]:
                lines.append(f"  ● PID={p['pid']}  {p['name']}")
            if len(self._file_lock_processes) > 15:
                lines.append(
                    f"  ... 还有 {len(self._file_lock_processes)-15} 个"
                )
            lines.append("")
        lines.append("⚠ 未保存的数据可能丢失！确定继续？")

        msg = "\n".join(lines)
        if not messagebox.askyesno("确认停止所有占用", msg, icon="warning"):
            return

        self._detection_is_restored = False
        self.run_in_thread(self._do_kill_all_file_lock)

    def _do_kill_all_file_lock(self):
        self.log_msg(f"\n{'='*50}")
        self.log_msg("  一键停止所有占用")
        self.log_msg(f"{'='*50}")

        svc_stopped = 0
        stopped_now = {}
        if self._file_lock_services:
            self.log_msg("\n步骤 1：停止占用服务...")
            for name, display in list(self._file_lock_services.items()):
                st = svc_status(name)
                if st == "running":
                    self.log_msg(f"  停止 {display} ({name}) ...")
                    rc, _, _ = run_cmd(f"net stop {name} /y", timeout=30)
                    if rc == 0 or svc_status(name) == "stopped":
                        self.log_msg(f"    [OK] 已停止")
                        stopped_now[name] = display
                        self._all_stopped_services[name] = display
                        svc_stopped += 1
                    else:
                        self.log_msg(f"    [!] 停止失败")
                else:
                    self.log_msg(f"  跳过 {display} - 已不在运行")
            self.log_msg(f"  共停止 {svc_stopped} 个服务")
        else:
            self.log_msg("\n步骤 1：无需停止的服务")

        time.sleep(0.5)

        killed = 0
        failed = 0
        if self._file_lock_processes:
            self.log_msg("\n步骤 2：结束占用进程...")
            for proc in self._file_lock_processes:
                pid = proc["pid"]
                name = proc["name"]
                if not is_process_alive(pid):
                    self.log_msg(f"  跳过 PID={pid} ({name}) — 已不存在")
                    continue
                self.log_msg(f"  结束 PID={pid} ({name}) ...")
                rc, _, err = run_cmd(f"taskkill /F /PID {pid}", timeout=10)
                if rc == 0:
                    self.log_msg(f"    [OK] 已结束")
                    killed += 1
                else:
                    err_s = (err.strip().split('\n')[0]
                             if err.strip() else "未知错误")
                    self.log_msg(f"    [!] 失败: {err_s}")
                    failed += 1
            self.log_msg(f"  成功结束 {killed} 个进程")
            if failed:
                self.log_msg(
                    f"  {failed} 个进程无法结束"
                    f"（可能是系统关键进程或已退出）")
                if not self._is_admin:
                    self.log_msg(
                        "  [提示] 以管理员身份运行可结束更多进程")
        else:
            self.log_msg("\n步骤 2：无需结束的进程")

        self._file_lock_processes = []
        self._file_lock_services = {}

        self.log_msg(f"\n[OK] 操作完成：停止 {svc_stopped} 个服务"
                     f"，结束 {killed} 个进程\n")

        # ★ 解除占用后询问是否删除
        saved_path = self._file_lock_path
        if saved_path and os.path.exists(saved_path):
            self.root.after(
                300,
                lambda p=saved_path: self._prompt_delete_after_unlock(p))
        elif self._all_stopped_services:
            self.root.after(
                0,
                lambda: self._show_service_restore_dialog(auto_popup=True))

    # ════════════════════════════════════════════════════════════
    #  解除占用后删除到回收站
    # ════════════════════════════════════════════════════════════

    def _prompt_delete_after_unlock(self, path):
        if not os.path.exists(path):
            self.log_msg(f"[信息] 路径已不存在（可能随进程退出已删除）: {path}")
            if self._all_stopped_services:
                self._show_service_restore_dialog(auto_popup=True)
            return

        is_dir = os.path.isdir(path)
        type_str = "文件夹" if is_dir else "文件"
        msg = (
            f"占用已解除。\n\n"
            f"是否将该{type_str}移到回收站？\n\n"
            f"路径: {path}\n\n"
            f"• 【是】→ 删除到回收站（可从回收站恢复）\n"
            f"• 【否】→ 保留不删除"
        )
        do_delete = messagebox.askyesno("删除确认", msg)

        if do_delete:
            self.log_msg(f"\n--- 删除到回收站 ---")
            self.log_msg(f"  路径: {path}")
            ok = send_to_recycle_bin(path)
            if ok and not os.path.exists(path):
                self.log_msg(f"  [OK] 已将{type_str}移到回收站\n")
            else:
                self.log_msg(f"  [!] 删除失败（可能仍有占用或权限不足）")
                self.log_msg(f"  [提示] 可手动删除或稍后重试\n")

        if self._all_stopped_services:
            self._show_service_restore_dialog(auto_popup=True)


if __name__ == "__main__":
    try:
        App()
    except Exception as e:
        messagebox.showerror("启动失败", str(e))