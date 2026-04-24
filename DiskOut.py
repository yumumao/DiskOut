# -*- coding: utf-8 -*-
"""
移动硬盘安全清理工具 v3.7
────────────────────────────────────────────────
功能清单：
  ■ 支持 USB 硬件级安全弹出（停止转动）
  ■ 自动检测并恢复脱机磁盘
  ■ 支持检测文件/文件夹占用进程和服务并一键结束
  ■ 支持普通模式运行，需要时提示提升管理员权限
  ■ 提权后自动恢复上次操作状态（含检测结果）
  ■ 管理员模式下自动放行 UIPI 拖放限制
  ■ 支持多分区 USB 设备安全弹出（整设备级别）
  ■ 全盘符磁盘分区分组映射（含本地硬盘）
  ■ 多层启发式虚拟盘识别（IOCTL / 设备路径 / Get-Partition 交叉验证）
  ■ 下拉菜单同磁盘分区合并显示，一键弹出整设备
  ■ 日志统一分组格式，始终显示全部盘符
  ■ DEF 盘开关切换瞬时响应（使用缓存，不重复检测）

变更记录（v3.7）：
  - 精简虚拟盘日志：仅输出识别结果行，不再输出启发式详情
  - DEF 盘开关切换使用缓存重建下拉菜单，不再触发后台重新检测
  - 虚拟磁盘状态栏/警告对话框使用通用描述（不假定具体工具）
  - 虚拟磁盘弹出前显示操作警告
  - 补全所有注释
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

# ══════════════════════════════════════════════════════════════
#   全局常量
# ══════════════════════════════════════════════════════════════

APP_VERSION = "3.7.0"  # ★ 版本号，v3.7: 精简虚拟盘日志 + DEF缓存刷新 + 虚拟盘通用标注

# ── 常见占用移动硬盘的 Windows 服务 ──
# 这些服务可能会在后台打开 USB 磁盘上的文件/目录，阻止安全弹出
SERVICES = [
    ("WSearch",       "Windows Search"),       # 文件索引服务
    ("SysMain",       "SysMain"),              # 预取/超级预取
    ("VSS",           "Volume Shadow Copy"),   # 卷影副本
    ("defragsvc",     "Optimize Drives"),      # 磁盘优化
    ("WMPNetworkSvc", "WMP Network Sharing"),  # 媒体共享
    ("StorSvc",       "Storage Service"),       # 存储服务
]

# ── 服务恢复建议（在恢复对话框中显示给用户参考） ──
SERVICE_RECOMMENDATIONS = {
    "WSearch":       ("建议恢复", "文件搜索和索引需要此服务"),
    "SysMain":       ("建议恢复", "优化应用启动速度和系统性能"),
    "VSS":           ("建议恢复", "系统还原和备份软件依赖此服务"),
    "defragsvc":     ("可选",     "磁盘优化按计划运行，不急需可稍后恢复"),
    "WMPNetworkSvc": ("可选",     "仅在使用 Windows Media Player 共享时需要"),
    "StorSvc":       ("建议恢复", "管理存储设置和可移动存储策略"),
}

# ── GetDriveTypeW 返回值映射 ──
# Windows API GetDriveTypeW 对每个盘符根路径返回一个整数类型码
DRIVE_TYPE_MAP = {
    2: "可移动",   # DRIVE_REMOVABLE — U盘、SD读卡器等
    3: "固定",     # DRIVE_FIXED    — 本地硬盘 / USB硬盘(有些报告为固定)
    4: "网络",     # DRIVE_REMOTE   — 网络映射驱动器
    5: "光驱",     # DRIVE_CDROM    — CD/DVD/BD 光驱
    6: "RAM",      # DRIVE_RAMDISK  — RAM 盘
}

# ── USB 总线类型集合 ──
# IOCTL_STORAGE_QUERY_PROPERTY 可能返回 "USB" 或 "USB3"
USB_BUS_TYPES = {"USB", "USB3"}

# ── 虚拟磁盘总线类型集合 ──
# 包含 IOCTL 报告的标准虚拟类型 + 启发式标记的 SIMULATED 类型
VIRTUAL_BUS_TYPES = {
    "VIRTUAL",             # BusType=14  Windows 原生 VHD/VHDX
    "FILEBACKEDVIRTUAL",   # BusType=15  文件支持的虚拟磁盘
    "SPACES",              # BusType=16  存储空间
    "SIMULATED",           # 启发式检测标记：Dokan/WinFsp/ImDisk 等模拟本地磁盘
}

# ── 已知虚拟文件系统的 NT 设备路径关键词 ──
# 通过 QueryDosDeviceW 查询盘符的 NT 设备路径可以判断底层驱动类型
# 真实磁盘分区的设备路径一定以 \Device\HarddiskVolume 开头
VIRTUAL_DEVICE_KEYWORDS = [
    "\\Dokan",       # Dokan 虚拟文件系统（多种云盘工具的后端）
    "\\WinFsp",      # WinFsp 虚拟文件系统（SSHFS 等工具的后端）
    "\\ImDisk",      # ImDisk RAM 盘 / 虚拟磁盘工具
]

# ── 提权状态传递临时文件路径 ──
# 提权前保存当前操作状态到此文件，提权后读取并恢复
STATE_FILE_PATH = os.path.join(
    os.environ.get("TEMP", os.path.dirname(os.path.abspath(sys.argv[0]))),
    "_diskout_elevation_state.json"
)

# ── USB 安全弹出 PowerShell 脚本 ──
# 通过 CfgMgr32.dll 的 CM_Request_Device_Eject 实现硬件级安全移除
# 这会使 USB 硬盘停止转动，而不仅仅是卸载文件系统
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


# ══════════════════════════════════════════════════════════════
#   工具函数
# ══════════════════════════════════════════════════════════════

def resource_path(relative_path):
    """
    获取资源文件路径（兼容 PyInstaller 打包后的临时目录）。
    PyInstaller 打包后，资源被解压到 sys._MEIPASS 指向的临时目录中。
    """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), relative_path)


def is_admin():
    """检查当前进程是否以管理员权限运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def run_cmd(cmd, timeout=120):
    """
    执行命令行命令并返回 (返回码, 标准输出, 标准错误)。
    使用 GBK 编码以兼容 Windows 中文系统的默认控制台编码。
    """
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
    """
    查询 Windows 服务状态。
    返回: "running" / "stopped" / "missing"
    """
    rc, out, _ = run_cmd(f"sc query {name}")
    if "RUNNING" in out:
        return "running"
    if "STOPPED" in out:
        return "stopped"
    return "missing"


def drive_exists(letter):
    """
    检查盘符是否存在（通过 GetLogicalDrives 位掩码判断）。
    GetLogicalDrives 返回一个 26 位的掩码，每一位对应一个字母 A-Z。
    """
    letter = letter.rstrip(":\\").upper()
    idx = ord(letter) - ord('A')
    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    return bool(bitmask & (1 << idx))


def get_drive_type_code(letter):
    """
    获取盘符的 GetDriveTypeW 返回值。
    返回: 2=可移动 3=固定 4=网络 5=光驱 6=RAM
    """
    letter = letter.rstrip(":\\").upper()
    return ctypes.windll.kernel32.GetDriveTypeW(f"{letter}:\\")


def get_drives_fast(min_letter='G'):
    """
    快速获取盘符列表（不含总线类型，仅基于 GetDriveTypeW）。
    用于初始化时的占位显示，后台识别完成后会更新为精确结果。
    """
    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    drives = []
    for i, ch in enumerate(string.ascii_uppercase):
        if ch < min_letter:
            continue
        if bitmask & (1 << i):
            dt = ctypes.windll.kernel32.GetDriveTypeW(f"{ch}:\\")
            drives.append((f"{ch}:", DRIVE_TYPE_MAP.get(dt, "未知")))
    return drives


def get_drive_bus_types():
    """
    通过 PowerShell 的 Get-Partition + Get-Disk 获取各盘符的总线类型。
    只会返回有真实物理磁盘支撑的分区（虚拟文件系统挂载的盘符不会出现）。
    因此该函数的结果可用于交叉验证：不在结果中的"固定盘"可能是虚拟盘。
    返回: {盘符大写字母: 总线类型字符串}
    """
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
    """
    通过 PowerShell 获取单个盘符的总线类型（单次查询，较慢约 1-2 秒）。
    仅在 IOCTL 查询失败时作为回退方案使用。
    """
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


# ══════════════════════════════════════════════════════════════
#   IOCTL / 底层磁盘操作函数
# ══════════════════════════════════════════════════════════════

# ── IOCTL 控制码 ──
_IOCTL_STORAGE_QUERY_PROPERTY    = 0x002D1400  # 查询存储属性（总线类型等）
_IOCTL_STORAGE_GET_DEVICE_NUMBER = 0x002D1080  # 获取设备编号（物理磁盘号）

# ── STORAGE_BUS_TYPE 枚举值到名称的映射 ──
# 对应 Windows SDK 中 STORAGE_BUS_TYPE 枚举
_BUS_TYPE_NAMES = {
    0: "Unknown", 1: "SCSI", 2: "ATAPI", 3: "ATA", 4: "1394",
    5: "SSA", 6: "Fibre", 7: "USB", 8: "RAID", 9: "iSCSI",
    10: "SAS", 11: "SATA", 12: "SD", 13: "MMC", 14: "Virtual",
    15: "FileBackedVirtual", 16: "Spaces", 17: "NVMe", 18: "SCM", 19: "UFS",
}


class _STORAGE_PROPERTY_QUERY(ctypes.Structure):
    """IOCTL_STORAGE_QUERY_PROPERTY 的输入结构体"""
    _fields_ = [
        ("PropertyId", ctypes.c_ulong),   # 0 = StorageDeviceProperty
        ("QueryType",  ctypes.c_ulong),   # 0 = PropertyStandardQuery
        ("Extra",      ctypes.c_byte * 1),
    ]


class _STORAGE_DEVICE_DESCRIPTOR(ctypes.Structure):
    """IOCTL_STORAGE_QUERY_PROPERTY 的输出结构体（仅读取前几个关键字段）"""
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
        ("BusType",               ctypes.c_ulong),   # STORAGE_BUS_TYPE 枚举值
    ]


class _STORAGE_DEVICE_NUMBER(ctypes.Structure):
    """IOCTL_STORAGE_GET_DEVICE_NUMBER 的输出结构体"""
    _fields_ = [
        ("DeviceType",      ctypes.c_ulong),   # FILE_DEVICE_DISK = 7
        ("DeviceNumber",    ctypes.c_ulong),   # 物理磁盘编号（0, 1, 2...）
        ("PartitionNumber", ctypes.c_ulong),   # 分区编号（1, 2, 3...）
    ]


def _open_volume_handle(letter):
    """
    以零权限打开卷句柄（\\\\.\\X:）。
    不需要管理员权限，不受文件占用影响，仅用于 IOCTL 查询。
    返回句柄整数，失败返回 None。
    """
    letter = letter.rstrip(":\\").upper()
    k32 = ctypes.windll.kernel32
    k32.CreateFileW.restype = ctypes.c_void_p
    INVALID = ctypes.c_void_p(-1).value
    # access=0 表示零权限打开，只用于发送 IOCTL 查询
    h = k32.CreateFileW(f"\\\\.\\{letter}:", 0, 0x3, None, 3, 0, None)
    if h is None or h == INVALID:
        return None
    return h


def get_bus_type_ioctl(letter):
    """
    通过 IOCTL_STORAGE_QUERY_PROPERTY 查询盘符的总线类型。
    直接调用内核驱动，速度极快（<1ms），不需要管理员权限。
    对于 Dokan/WinFsp 等虚拟文件系统，IOCTL 通常失败，返回 "Unknown"。
    """
    h = _open_volume_handle(letter)
    if h is None:
        return "Unknown"
    k32 = ctypes.windll.kernel32
    try:
        query = _STORAGE_PROPERTY_QUERY()
        query.PropertyId = 0   # StorageDeviceProperty
        query.QueryType  = 0   # PropertyStandardQuery
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
    """
    通过 IOCTL_STORAGE_GET_DEVICE_NUMBER 获取盘符对应的物理磁盘号。
    用于判断哪些盘符属于同一物理磁盘（多分区设备）。
    对于 Dokan/WinFsp 虚拟文件系统，IOCTL 通常失败，返回 None。
    """
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


def get_dos_device(letter):
    """
    使用 QueryDosDeviceW 查询盘符对应的 NT 设备路径。
    不需要管理员权限，执行速度极快（纯内核调用）。

    不同类型的盘符返回不同的设备路径，例如：
      真实磁盘分区 → \\Device\\HarddiskVolume3
      VHD/VHDX    → \\Device\\HarddiskVolume10  （注册为真实卷）
      Dokan 虚拟盘 → \\Device\\Dokan_1{guid}
      WinFsp 虚拟盘 → \\Device\\WinFsp.Disk\\...
      ImDisk RAM盘 → \\Device\\ImDisk0
      网络驱动器   → \\Device\\LanmanRedirector\\...
    """
    letter = letter.rstrip(":\\").upper()
    k32 = ctypes.windll.kernel32
    buf = ctypes.create_unicode_buffer(1024)
    result = k32.QueryDosDeviceW(f"{letter}:", buf, 1024)
    if result:
        return buf.value
    return ""


def _is_virtual_device_path(device_path, drive_type):
    """
    启发式判断——通过 NT 设备路径判断是否为虚拟文件系统。
    仅对 DriveType=3（固定磁盘）应用此检测。

    原理：
    ─────────────────────────────────────────────────────────────
    真实磁盘分区的设备路径一定是 \\Device\\HarddiskVolumeN：
      - SATA/NVMe/USB 磁盘的分区都使用此路径
      - VHD/VHDX 挂载后也使用此路径（但它们已通过 BusType=14/15 识别）

    虚拟文件系统驱动使用自定义设备路径：
      - Dokan  → \\Device\\Dokan_1{...}
      - WinFsp → \\Device\\WinFsp.Disk\\...
      - ImDisk → \\Device\\ImDisk0
      - 其他   → 各种非标准路径

    因此对于报告为"固定磁盘"的盘符：
      如果设备路径不是 \\Device\\HarddiskVolume 开头 → 几乎可以确定是虚拟盘
    ─────────────────────────────────────────────────────────────

    参数:
        device_path: QueryDosDeviceW 返回的 NT 设备路径
        drive_type:  GetDriveTypeW 返回值

    返回: True 表示判定为虚拟设备
    """
    # 仅对"固定磁盘"(type=3) 应用此启发式检测
    if drive_type != 3 or not device_path:
        return False
    upper = device_path.upper()
    # 检查已知虚拟文件系统驱动的设备路径关键词
    for keyword in VIRTUAL_DEVICE_KEYWORDS:
        if keyword.upper() in upper:
            return True
    # 真实磁盘分区一定以 \Device\HarddiskVolume 开头
    # 如果一个"固定磁盘"的设备路径不匹配，几乎可以确定是虚拟/模拟的
    if not upper.startswith("\\DEVICE\\HARDDISKVOLUME"):
        return True
    return False


def get_all_partitions_on_disk(disk_number):
    """
    获取指定物理磁盘号上的所有盘符。
    遍历 A-Z 所有已挂载的盘符，通过 IOCTL 查询其磁盘号是否匹配。
    用于多分区 USB 设备的整体弹出。
    """
    if disk_number is None:
        return []
    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    partitions = []
    for i, ch in enumerate(string.ascii_uppercase):
        if bitmask & (1 << i):
            try:
                dn = get_disk_number_ioctl(ch)
                if dn == disk_number:
                    partitions.append(ch)
            except Exception:
                pass
    return sorted(partitions)


def get_offline_disks():
    """
    通过 PowerShell Get-Disk 获取所有脱机磁盘列表。
    脱机磁盘通常是上次使用 Set-Disk -IsOffline 弹出后遗留的状态。
    """
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
        for line in lines[1:]:  # 跳过 CSV 标题行
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
    """将脱机磁盘恢复联机（需要管理员权限）"""
    cmd = (
        f'powershell -NoProfile -Command "'
        f'Set-Disk -Number {disk_number} -IsOffline $false"'
    )
    rc, _, _ = run_cmd(cmd, timeout=15)
    return rc == 0


def eject_volume_api(letter):
    """
    通过 DeviceIoControl 的 IOCTL_STORAGE_EJECT_MEDIA 弹出卷。
    执行步骤：FlushFileBuffers → FSCTL_LOCK_VOLUME → FSCTL_DISMOUNT → IOCTL_EJECT
    返回: (成功布尔, 消息字符串, 警告列表)
    """
    letter = letter.rstrip(":\\")
    volume = f"\\\\.\\{letter}:"
    k32 = ctypes.windll.kernel32
    k32.CreateFileW.restype = ctypes.c_void_p
    SHARE_RW       = 0x1 | 0x2       # FILE_SHARE_READ | FILE_SHARE_WRITE
    OPEN_EXISTING  = 3
    FSCTL_LOCK     = 0x00090018       # 锁定卷（排他访问）
    FSCTL_DISMOUNT = 0x00090020       # 卸载文件系统
    IOCTL_EJECT    = 0x002D4808       # 弹出存储介质
    INVALID_HANDLE = ctypes.c_void_p(-1).value
    warnings = []
    write_access = False
    # 尝试以不同权限级别打开卷（从高到低）
    for access in [0xC0000000, 0x80000000, 0]:
        h = k32.CreateFileW(volume, access, SHARE_RW, None, OPEN_EXISTING, 0, None)
        if h is not None and h != INVALID_HANDLE:
            write_access = (access == 0xC0000000)
            break
    else:
        return False, "无法打开卷句柄", []
    br = wintypes.DWORD(0)
    # 步骤1: 刷新写缓冲区（确保数据安全）
    if write_access:
        if not k32.FlushFileBuffers(h):
            warnings.append("FlushFileBuffers 失败，写缓冲区可能未完全刷新")
    else:
        warnings.append("无写权限打开卷，FlushFileBuffers 可能无效")
        k32.FlushFileBuffers(h)
    # 步骤2: 锁定卷（获取排他访问权）
    locked = k32.DeviceIoControl(
        h, FSCTL_LOCK, None, 0, None, 0, ctypes.byref(br), None
    )
    if not locked:
        warnings.append("卷锁定失败（有程序占用），强制继续")
    # 步骤3: 卸载文件系统
    dismounted = k32.DeviceIoControl(
        h, FSCTL_DISMOUNT, None, 0, None, 0, ctypes.byref(br), None
    )
    if not dismounted:
        warnings.append("卸载文件系统失败")
    # 步骤4: 弹出存储介质
    ok = k32.DeviceIoControl(
        h, IOCTL_EJECT, None, 0, None, 0, ctypes.byref(br), None
    )
    k32.CloseHandle(h)
    msg = "API 弹出指令已发送" if ok else "API IOCTL 失败"
    return bool(ok), msg, warnings


# ══════════════════════════════════════════════════════════════
#   Restart Manager 文件占用检测
# ══════════════════════════════════════════════════════════════

class RM_UNIQUE_PROCESS(ctypes.Structure):
    """Restart Manager 用于唯一标识进程的结构体"""
    _fields_ = [
        ("dwProcessId", wintypes.DWORD),
        ("ProcessStartTime", wintypes.FILETIME),
    ]


class RM_PROCESS_INFO(ctypes.Structure):
    """Restart Manager 返回的进程信息结构体"""
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
    """
    使用 Windows Restart Manager API 检测哪些进程锁定了指定文件。
    这是 Windows 官方推荐的文件占用检测方式，比 handle.exe 更快更可靠。
    参数: paths — 文件路径字符串或字符串列表
    返回: [{pid, name, service}] 占用进程列表
    """
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
    # 启动 RM 会话
    ret = rm.RmStartSession(ctypes.byref(session_handle), 0, session_key)
    if ret != 0:
        return results
    try:
        # 注册要检测的文件
        n = len(paths)
        arr = (ctypes.c_wchar_p * n)(*paths)
        ret = rm.RmRegisterResources(
            session_handle.value, n, arr, 0, None, 0, None
        )
        if ret != 0:
            return results
        # 查询占用进程
        needed = wintypes.UINT(0)
        count = wintypes.UINT(0)
        reason = wintypes.DWORD(0)
        ret = rm.RmGetList(
            session_handle.value,
            ctypes.byref(needed), ctypes.byref(count),
            None, ctypes.byref(reason),
        )
        # ret=234 (ERROR_MORE_DATA) 表示有数据需要读取
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
    """
    递归收集目录内的文件列表（用于批量 RM 检测）。
    限制最大文件数和递归深度以避免扫描时间过长。
    """
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


# ══════════════════════════════════════════════════════════════
#   Shell 操作 / 进程检测
# ══════════════════════════════════════════════════════════════

class _SHFILEOPSTRUCTW(ctypes.Structure):
    """SHFileOperationW 的参数结构体（用于删除文件到回收站）"""
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
    """
    将文件/文件夹移到回收站（可从回收站恢复）。
    使用 SHFileOperationW 的 FO_DELETE + FOF_ALLOWUNDO 标志组合。
    """
    FO_DELETE          = 0x0003
    FOF_ALLOWUNDO      = 0x0040  # 允许撤销（放入回收站而非永久删除）
    FOF_NOCONFIRMATION = 0x0010  # 不弹出确认对话框
    op = _SHFILEOPSTRUCTW()
    op.hwnd = None
    op.wFunc = FO_DELETE
    op.pFrom = path + '\0'  # 必须双零结尾
    op.pTo = None
    op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION
    op.fAnyOperationsAborted = 0
    op.hNameMappings = None
    op.lpszProgressTitle = None
    result = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    return result == 0 and not op.fAnyOperationsAborted


def is_process_alive(pid):
    """
    检查指定 PID 的进程是否仍在运行。
    通过 OpenProcess + GetExitCodeProcess 判断，
    STILL_ACTIVE (259) 表示进程仍在运行。
    """
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


# ══════════════════════════════════════════════════════════════
#   UI 样式常量（推荐按钮外观）
# ══════════════════════════════════════════════════════════════

REC_BG       = "#dae8fc"    # 推荐按钮背景色（浅蓝）
REC_FG       = "#1a3a6b"    # 推荐按钮文字色（深蓝）
REC_ACTIVE   = "#b8d4f0"    # 推荐按钮激活色
STAR_COLOR   = "#c8a000"    # 星标颜色（金色）
STAR_HOVER   = "#ffe066"    # 星标悬停色（亮金色）
REC_HOVER_BG = "#3b7dd8"    # 推荐按钮悬停背景（中蓝）
REC_HOVER_FG = "#ffffff"    # 推荐按钮悬停文字（白色）


# ══════════════════════════════════════════════════════════════
#   主应用类
# ══════════════════════════════════════════════════════════════

class App:
    def __init__(self):
        self._is_admin = is_admin()
        self.root = tk.Tk()
        title_mode = "管理员模式" if self._is_admin else "普通模式"
        self.root.title(f"移动硬盘弹出工具 - {title_mode}")
        self.root.geometry("620x750")
        self.root.minsize(480, 550)
        self._set_icon(self.root)

        # ── 运行状态 ──
        self._busy = False                 # 是否有操作正在执行（防止重复操作）
        self._detecting = False            # 是否正在后台识别磁盘类型
        self._dnd_ok = False               # 拖放功能是否可用

        # ── 服务管理 ──
        self._all_stopped_services = {}    # 已停止的服务 {name: display}

        # ── 磁盘识别缓存 ──
        self._bus_cache = {}               # {letter: bus_type_string}  总线类型缓存
        self._disk_map = {}                # {disk_number: [letters]}   磁盘号→分区映射
        self._letter_to_disk = {}          # {letter: disk_number}      盘符→磁盘号映射
        self._labels = {}                  # {letter: label}  如 "USB硬盘"/"固定"/"虚拟盘"
        self._combo_to_primary = {}        # {下拉显示值: 主盘符字母}

        # ── 弹出操作状态 ──
        self._eject_disk_number = None     # 当前弹出的磁盘号
        self._eject_all_partitions = []    # 当前弹出涉及的所有分区字母列表
        self._eject_is_multi = False       # 是否为多分区弹出

        # ── 文件占用检测状态 ──
        self._file_lock_processes = []     # 检测到的占用进程列表
        self._file_lock_services = {}      # 检测到的占用服务 {name: display}
        self._file_lock_path = ""          # 检测的目标路径
        self._detection_is_restored = False  # 检测结果是否从提权状态恢复

        # ── 构建 UI 并启动 ──
        self.build_ui()
        self._restore_state_from_elevation()
        self.root.after(100, self._start_bus_detection)
        self.root.after(500, self._check_offline_on_start)
        self.root.mainloop()

    # ════════════════════════════════════════════════════════════
    #  窗口图标 / 管理员权限
    # ════════════════════════════════════════════════════════════

    def _set_icon(self, window):
        """尝试设置窗口图标（兼容 PyInstaller 打包，搜索多种文件名）"""
        try:
            for ico_name in ("diskout.ico", "DiskOut.ico", "Diskout.ico"):
                ico_path = resource_path(ico_name)
                if os.path.isfile(ico_path):
                    window.iconbitmap(ico_path)
                    return
        except Exception:
            pass

    def _require_admin(self, operation="此操作"):
        """
        检查是否有管理员权限，没有则弹出提示并提供提升选项。
        返回: True（已有权限），False（无权限且用户未选择提升）
        """
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

    # ════════════════════════════════════════════════════════════
    #  提权状态保存与恢复
    # ════════════════════════════════════════════════════════════

    def _save_state_for_elevation(self):
        """
        提权前保存当前操作状态到临时文件。
        保存内容包括：盘符选择、当前标签页、DEF开关、文件路径、检测结果等。
        """
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
                "file_lock_path": self._file_lock_path,
                "file_lock_processes": self._file_lock_processes,
                "file_lock_services": self._file_lock_services,
            }
            with open(STATE_FILE_PATH, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False)
        except Exception:
            pass

    def _restore_state_from_elevation(self):
        """
        提权后从临时文件恢复操作状态。
        如果状态文件超过 120 秒则视为过期并丢弃。
        如果包含检测结果，会验证进程是否仍存活并弹出恢复对话框。
        """
        try:
            if not os.path.exists(STATE_FILE_PATH):
                return
            # 超过 120 秒视为过期
            age = time.time() - os.path.getmtime(STATE_FILE_PATH)
            if age > 120:
                os.remove(STATE_FILE_PATH)
                return
            with open(STATE_FILE_PATH, "r", encoding="utf-8") as f:
                state = json.load(f)
            os.remove(STATE_FILE_PATH)
            restored = []
            # 恢复文件路径
            if state.get("file_path"):
                self.file_path_var.set(state["file_path"])
                restored.append(f"文件路径: {state['file_path']}")
            # 恢复 DEF 开关
            if state.get("show_def"):
                self.show_def_var.set(True)
                self.drive_frame.config(
                    text="盘符选择（D: 及之后 ⚠ 含本地硬盘）")
                restored.append("D/E/F 盘显示: 已启用")
            # 恢复盘符选择
            if state.get("drive"):
                drive_prefix = state["drive"].split()[0] if state["drive"] else ""
                if drive_prefix:
                    self.drive_var.set(state["drive"])
                    restored.append(f"盘符: {drive_prefix}")
            # 恢复标签页
            if state.get("tab") is not None:
                try:
                    self.notebook.select(state["tab"])
                except Exception:
                    pass
                tab_names = ["解除磁盘占用/弹出", "文件/文件夹占用", "进阶功能"]
                idx = state["tab"]
                if 0 <= idx < len(tab_names):
                    restored.append(f"标签页: {tab_names[idx]}")
            # 恢复检测结果
            has_detection = False
            if state.get("file_lock_path"):
                self._file_lock_path = state["file_lock_path"]
            if state.get("file_lock_processes"):
                self._file_lock_processes = state["file_lock_processes"]
                has_detection = True
            if state.get("file_lock_services"):
                self._file_lock_services = state["file_lock_services"]
                has_detection = True
            # 输出恢复信息到日志
            if restored:
                self.log_msg("[恢复] 已从提权前恢复以下状态：")
                for item in restored:
                    self.log_msg(f"  • {item}")
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
                # 验证进程是否仍存活（提权期间可能已退出）
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
                # 延迟弹出恢复对话框
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
        """提权后显示检测结果恢复对话框，提供三种操作选择"""
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
            dlg, text="✓ 已成功提升为管理员权限",
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

        ttk.Button(btn_frame, text="🔍 重新检测（推荐）",
                   command=do_redetect).pack(side="left", padx=(0, 6))
        ttk.Button(btn_frame, text="⚡ 直接解除占用",
                   command=do_kill_now).pack(side="left", padx=(0, 6))
        ttk.Button(btn_frame, text="稍后决定",
                   command=do_later).pack(side="right")
        dlg.protocol("WM_DELETE_WINDOW", do_later)
        # 居中对话框
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
        """清理提权状态临时文件"""
        try:
            if os.path.exists(STATE_FILE_PATH):
                os.remove(STATE_FILE_PATH)
        except Exception:
            pass

    def _restart_as_admin(self):
        """
        以管理员身份重新启动程序。
        先保存当前状态，然后调用 ShellExecuteW("runas") 执行提权启动。
        """
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
            if ret > 32:  # >32 表示启动成功
                self.root.destroy()
                sys.exit(0)
            else:
                self._cleanup_state_file()
                self.log_msg("[提示] 用户取消了权限提升，或提升失败")
        except Exception as e:
            self._cleanup_state_file()
            self.log_msg(f"[错误] 提升权限失败: {e}")

    def _request_admin_elevation(self):
        """用户主动点击「提升为管理员」按钮时调用"""
        msg = (
            "将以管理员身份重新启动程序。\n"
            "当前窗口将关闭，操作状态（含检测结果）将自动恢复。\n\n"
            "是否继续？"
        )
        if messagebox.askyesno("提升为管理员", msg):
            self._restart_as_admin()

    def _allow_drag_drop_admin(self):
        """
        管理员模式下放行 UIPI（用户界面权限隔离）拖放消息限制。
        Windows 默认阻止低权限进程（如资源管理器）向高权限进程拖放文件。
        需要对 WM_DROPFILES、WM_COPYDATA、WM_COPYGLOBALDATA 三个消息放行。
        """
        if not self._is_admin:
            return
        MSGFLT_ALLOW        = 1
        WM_DROPFILES        = 0x0233
        WM_COPYDATA         = 0x004A
        WM_COPYGLOBALDATA   = 0x0049
        user32 = ctypes.windll.user32
        # 全局消息过滤器
        try:
            _filter = user32.ChangeWindowMessageFilter
            _filter.argtypes = [wintypes.UINT, wintypes.DWORD]
            _filter.restype = wintypes.BOOL
            _filter(WM_DROPFILES, MSGFLT_ALLOW)
            _filter(WM_COPYDATA, MSGFLT_ALLOW)
            _filter(WM_COPYGLOBALDATA, MSGFLT_ALLOW)
        except Exception:
            pass
        # 窗口级消息过滤器（更精确）
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
            # Tkinter 的顶层窗口句柄可能和 winfo_id 不同
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
    #  推荐按钮（自定义外观，非 ttk 原生样式）
    # ════════════════════════════════════════════════════════════

    def _make_rec_btn(self, parent, command):
        """创建带星标的推荐操作按钮（蓝底金星，悬停反色）"""
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
        # 所有子控件列表（用于统一改色）
        ws = [frm, inner, line1, lbl_star, lbl_title, lbl_sub]

        def _hover_enter(e):
            for w in ws:
                w.configure(bg=REC_HOVER_BG)
            lbl_star.config(fg=STAR_HOVER)
            lbl_title.config(fg=REC_HOVER_FG)
            lbl_sub.config(fg="#ccc")

        def _hover_leave(e):
            # 检查鼠标是否仍在按钮区域内（防止子控件之间切换时闪烁）
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

    # ════════════════════════════════════════════════════════════
    #  构建 UI
    # ════════════════════════════════════════════════════════════

    def build_ui(self):
        """构建主界面所有控件"""
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
        style.configure("TCombobox",         font=UI_FONT, arrowsize=18)

        m = ttk.Frame(self.root, padding=8)
        m.pack(fill="both", expand=True)

        # ── 顶部状态行：版本号 + 权限状态 ──
        top_row = ttk.Frame(m)
        top_row.pack(fill="x", pady=(0, 2))
        ver_lbl = ttk.Label(
            top_row, text=f"v{APP_VERSION}",
            foreground="#909090", font=("Consolas", 10),
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

        # ── 盘符选择区域 ──
        self.drive_frame = ttk.LabelFrame(
            m, text="盘符选择（仅 G: 及之后）", padding=8
        )
        self.drive_frame.pack(fill="x", pady=(0, 4))

        row1 = ttk.Frame(self.drive_frame)
        row1.pack(fill="x")
        # 初始用快速扫描结果占位（后台识别完成后会更新为精确分组）
        drives = get_drives_fast('G')
        values = [f"{d[0]}  [{d[1]}]" for d in drives]
        self.drive_var = tk.StringVar(value=values[0] if values else "")
        # 加宽下拉菜单以适应分组显示（如 "H:, I:  [USB硬盘] 磁盘4"）
        self.combo = ttk.Combobox(
            row1, textvariable=self.drive_var,
            values=values, state="readonly",
            width=30, font=("Consolas", 11)
        )
        self.combo.pack(side="left", padx=(0, 8))
        self.combo.bind("<<ComboboxSelected>>", self._on_drive_selected)
        ttk.Button(row1, text="刷新", width=5, command=self.refresh).pack(side="left")
        self.status_lbl = ttk.Label(row1, text="", foreground="gray")
        self.status_lbl.pack(side="left", padx=10)
        if not drives:
            self.status_lbl.config(
                text="!! 未检测到 G: 及之后的盘符", foreground="red")
        else:
            self.status_lbl.config(
                text="正在识别磁盘类型...", foreground="blue")

        # ── DEF 盘开关（启用后可选择 D:/E:/F: 盘，通常为本地硬盘）──
        row2 = ttk.Frame(self.drive_frame)
        row2.pack(fill="x", pady=(6, 0))
        self.show_def_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            row2,
            text="启用 D: / E: / F: 盘（⚠ 通常为本地硬盘，谨慎操作）",
            variable=self.show_def_var,
            command=self._toggle_def,
        ).pack(side="left")

        # ── 功能标签页 ──
        nb = ttk.Notebook(m)
        self.notebook = nb
        gk = dict(sticky="nsew", padx=3, pady=3, ipady=2)  # 通用 grid 参数

        # ──── 标签页 1: 解除磁盘占用 / 弹出 ────
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

        # ──── 标签页 2: 文件/文件夹占用 ────
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
        ttk.Button(path_row, text="文件", width=5,
                   command=self.browse_file).pack(side="left", padx=1)
        ttk.Button(path_row, text="文件夹", width=6,
                   command=self.browse_folder).pack(side="left", padx=1)
        t2_btns = ttk.Frame(t2)
        t2_btns.pack(fill="x", pady=(0, 4))
        t2_btns.columnconfigure(0, weight=1)
        t2_btns.columnconfigure(1, weight=1)
        t2_btns.columnconfigure(2, weight=1)
        ttk.Button(t2_btns, text="检测占用",
                   command=self.detect_file_lock
                   ).grid(row=0, column=0, sticky="ew", padx=3, pady=2)
        ttk.Button(t2_btns, text="一键停止所有占用",
                   command=self.kill_all_file_lock
                   ).grid(row=0, column=1, sticky="ew", padx=3, pady=2)
        ttk.Button(t2_btns, text="恢复已停止的服务",
                   command=lambda: self._show_service_restore_dialog(
                       auto_popup=False)
                   ).grid(row=0, column=2, sticky="ew", padx=3, pady=2)
        # 拖放支持初始化
        if HAS_WINDND:
            try:
                self._allow_drag_drop_admin()
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

        # ──── 标签页 3: 进阶功能 ────
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

        # ── 执行日志区域 ──
        f4 = ttk.LabelFrame(m, text="执行日志", padding=4)
        f4.pack(fill="both", expand=True, pady=(4, 0))
        btn_bar = ttk.Frame(f4)
        btn_bar.pack(side="bottom", fill="x", pady=(2, 0))
        ttk.Button(btn_bar, text="清空日志",
                   command=lambda: self.log.delete("1.0", tk.END)
                   ).pack(anchor="e")
        self.log = scrolledtext.ScrolledText(
            f4, height=1, font=("Consolas", 10), wrap=tk.WORD
        )
        self.log.pack(fill="both", expand=True)

        # ── 延迟调整标签页高度 ──
        self.root.after(150, self._resize_notebook_to_current)

        # ── 初始化日志消息 ──
        mode_text = "管理员模式" if self._is_admin else "普通模式"
        self.log_msg(f"[OK] 工具已启动（{mode_text}）")
        if not self._is_admin:
            self.log_msg("[提示] 当前为普通模式，部分功能需要管理员权限，可点击右上角「提升为管理员」按钮获取完整功能")
        ds = ", ".join(f"{d[0]}[{d[1]}]" for d in drives) if drives else "无"
        self.log_msg(f"初始盘符：{ds}")
        self.log_msg("正在后台识别磁盘总线类型...")

    def _on_tab_changed(self, event=None):
        """标签页切换时自适应高度"""
        self.root.after(20, self._resize_notebook_to_current)

    def _resize_notebook_to_current(self):
        """调整 Notebook 高度以适应当前标签页内容"""
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
        """
        显示服务恢复对话框，可勾选要恢复的服务。
        auto_popup=True 时为自动弹出（弹出后提示），False 为用户主动点击。
        """
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
        # 服务勾选列表
        cb_frame = ttk.Frame(dlg, padding=(8, 0))
        cb_frame.pack(fill="x", padx=20)
        check_vars = {}
        for name, display in to_restore.items():
            rec_level, rec_text = SERVICE_RECOMMENDATIONS.get(name, ("", ""))
            var = tk.BooleanVar(value=True)  # 默认勾选
            check_vars[name] = var
            row = ttk.Frame(cb_frame)
            row.pack(fill="x", pady=(4, 0))
            cb = ttk.Checkbutton(
                row, text=f"{display} ({name})", variable=var)
            cb.pack(anchor="w")
            # 显示恢复建议信息
            if rec_level:
                if rec_level == "建议恢复":
                    hint_color = "#1a7f1a"
                    prefix = "→ 建议恢复"
                else:
                    hint_color = "#888888"
                    prefix = "→ 可选"
                ttk.Label(
                    row, text=f"    {prefix} — {rec_text}",
                    foreground=hint_color,
                    font=("Microsoft YaHei UI", 9),
                ).pack(anchor="w", padx=(20, 0))
        # 操作按钮
        btn_frame = ttk.Frame(dlg, padding=8)
        btn_frame.pack(fill="x", padx=20, pady=(10, 16))

        def select_all():
            for v in check_vars.values():
                v.set(True)

        def select_none():
            for v in check_vars.values():
                v.set(False)

        def do_restore():
            selected = {n: to_restore[n]
                        for n, v in check_vars.items() if v.get()}
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
        # 居中对话框
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
        """恢复用户选中的服务（后台线程执行）"""
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

    # ════════════════════════════════════════════════════════════
    #  后台磁盘识别（多层虚拟盘检测 + 分组日志 + 下拉合并）
    # ════════════════════════════════════════════════════════════

    def _start_bus_detection(self):
        """启动后台磁盘识别线程"""
        self._detecting = True
        threading.Thread(target=self._do_bus_detection, daemon=True).start()

    def _do_bus_detection(self):
        """
        后台线程：全盘扫描并识别所有磁盘的总线类型、分区分组、虚拟盘类型。

        扫描 A-Z 所有盘符用于日志显示（始终完整显示），
        仅 min_letter 及之后的盘符用于下拉菜单选择。

        虚拟盘检测三层启发式：
          A. NT 设备路径检查（QueryDosDeviceW）— 非 HarddiskVolume 即虚拟
          B. 无物理磁盘号 + Unknown 总线 — 底层无真实设备
          C. Get-Partition 交叉验证 — 不在分区列表中 = 无物理磁盘支撑
        """
        min_letter = 'D' if self.show_def_var.get() else 'G'

        # ── 1. 扫描所有盘符 (A-Z) ──
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        all_letters = []       # 所有存在的盘符字母（已排序）
        drive_types = {}       # {letter: GetDriveTypeW 返回值}
        for i, ch in enumerate(string.ascii_uppercase):
            if bitmask & (1 << i):
                all_letters.append(ch)
                drive_types[ch] = ctypes.windll.kernel32.GetDriveTypeW(f"{ch}:\\")

        if not all_letters:
            self.root.after(0, lambda: self._apply_bus_detection(
                log_line="(无盘符)", dropdown_values=[], combo_map={},
                disk_map={}, letter_to_disk={}, all_bus={}, labels={}))
            return

        # ── 2. 获取每个盘符的 IOCTL 总线类型、磁盘号、NT 设备路径 ──
        ioctl_bus = {}         # {letter: bus_type_string}  IOCTL 查询结果
        disk_map = {}          # {disk_number: [letters]}   磁盘号→分区映射
        letter_to_disk = {}    # {letter: disk_number}      盘符→磁盘号映射
        device_paths = {}      # {letter: NT设备路径}

        for ch in all_letters:
            # IOCTL 总线类型查询（速度极快，<1ms）
            ioctl_bus[ch] = get_bus_type_ioctl(ch)
            # IOCTL 磁盘号查询（用于多分区分组）
            dn = get_disk_number_ioctl(ch)
            if dn is not None:
                disk_map.setdefault(dn, []).append(ch)
                letter_to_disk[ch] = dn
            # NT 设备路径查询（用于虚拟文件系统检测）
            device_paths[ch] = get_dos_device(ch)

        # ── 3. PowerShell Get-Partition 交叉验证 ──
        # Get-Partition 只返回有真实物理磁盘支撑的分区
        # Dokan/WinFsp/ImDisk 等虚拟文件系统挂载的盘符不会出现在结果中
        ps_bus = get_drive_bus_types()

        # ── 4. 合并总线类型：IOCTL 优先，PowerShell 补充 ──
        all_bus = dict(ioctl_bus)
        for ch, bus in ps_bus.items():
            if all_bus.get(ch, "Unknown") == "Unknown" and bus != "Unknown":
                all_bus[ch] = bus

        # ── 5. 确定每个盘符的分类标签（多层虚拟盘启发式检测）──
        labels = {}            # {letter: "USB硬盘"/"固定"/"虚拟盘"/"可移动"/"网络"...}

        for ch in all_letters:
            dt = drive_types[ch]
            bus = all_bus.get(ch, "Unknown").upper()

            # ── 5a. USB 设备（优先级最高）──
            if bus in USB_BUS_TYPES:
                labels[ch] = "USB硬盘" if dt == 3 else "可移动"
                continue

            # ── 5b. 已知虚拟总线类型（IOCTL 报告 BusType=14/15/16）──
            if bus in VIRTUAL_BUS_TYPES:
                labels[ch] = "虚拟盘"
                continue

            # ── 5c. 固定磁盘的三层启发式虚拟盘检测 ──
            if dt == 3:
                is_virtual = False

                # 启发式 A: NT 设备路径检查
                # 真实磁盘分区的设备路径一定是 \Device\HarddiskVolumeN
                dev_path = device_paths.get(ch, "")
                if _is_virtual_device_path(dev_path, dt):
                    is_virtual = True

                # 启发式 B: 无物理磁盘号 + 总线类型未知
                if not is_virtual:
                    has_disk_num = ch in letter_to_disk
                    if not has_disk_num and bus in ("UNKNOWN", ""):
                        is_virtual = True

                # 启发式 C: 不在 Get-Partition 结果中（终极验证）
                if not is_virtual:
                    in_ps = ch in ps_bus
                    has_disk_num = ch in letter_to_disk
                    if not in_ps and not has_disk_num:
                        is_virtual = True
                    elif not in_ps and has_disk_num and bus in ("UNKNOWN", ""):
                        is_virtual = True

                if is_virtual:
                    labels[ch] = "虚拟盘"
                    all_bus[ch] = "Simulated"  # 标记为启发式识别的虚拟类型
                else:
                    labels[ch] = "固定"
                continue

            # ── 5d. 其他类型（可移动/网络/光驱/RAM 等）──
            labels[ch] = DRIVE_TYPE_MAP.get(dt, "未知")

        # ── 6. 构建统一分组日志行（包含 A-Z 所有盘符）──
        log_line = self._build_drive_log_line(
            all_letters, labels, letter_to_disk, disk_map, all_bus)

        # ── 7. 构建下拉菜单项（仅 min_letter 及之后，同磁盘合并）──
        dropdown_values, combo_map = self._build_dropdown_groups(
            all_letters, min_letter, labels, letter_to_disk, disk_map, all_bus)

        # ── 8. 传递到 UI 线程更新界面 ──
        self.root.after(0, lambda: self._apply_bus_detection(
            log_line=log_line,
            dropdown_values=dropdown_values,
            combo_map=combo_map,
            disk_map=disk_map,
            letter_to_disk=letter_to_disk,
            all_bus=all_bus,
            labels=labels,
        ))

    def _build_drive_log_line(self, all_letters, labels, letter_to_disk,
                               disk_map, all_bus):
        """
        构建盘符识别的统一日志行。

        格式示例：
          C:, D:[磁盘2(NVMe), 固定], E:[固定], G:[USB硬盘],
          H:, I:[磁盘4, USB硬盘], Z:[虚拟盘]

        规则：
        ─────────────────────────────────────────────────────────
        1. 相邻的同磁盘盘符分为一组，标注放在最后一个字母上
        2. 不相邻的同磁盘盘符各自标注磁盘号
        3. 单分区盘符（该磁盘只有一个分区）只标注类型，不加磁盘号
        4. 总线类型已包含在标签中时不重复显示
           例如 "USB硬盘" 已隐含 USB，不再加 (USB)
        ─────────────────────────────────────────────────────────
        """
        parts = []
        i = 0
        while i < len(all_letters):
            ch = all_letters[i]
            dn = letter_to_disk.get(ch)

            if dn is not None:
                # ── 收集在 all_letters 中连续出现的同磁盘盘符 ──
                group = [ch]
                j = i + 1
                while j < len(all_letters):
                    if letter_to_disk.get(all_letters[j]) == dn:
                        group.append(all_letters[j])
                        j += 1
                    else:
                        break
                i = j

                label = labels.get(group[0], "未知")
                bus = all_bus.get(group[0], "Unknown")

                # 判断是否需要显示总线信息
                bus_upper = bus.upper()
                show_bus = (
                    bus_upper not in ("UNKNOWN", "SIMULATED", "", "USB", "USB3")
                    and bus_upper not in label.upper()
                )
                bus_str = f"({bus})" if show_bus else ""

                n_total = len(disk_map.get(dn, []))  # 该磁盘的总分区数

                if len(group) == 1:
                    if n_total > 1:
                        # 该磁盘有多个分区但此处只出现一个（不相邻情况）
                        parts.append(f"{ch}:[磁盘{dn}{bus_str}, {label}]")
                    else:
                        # 该磁盘就一个分区
                        parts.append(f"{ch}:[{label}]")
                else:
                    # 多个相邻盘符属于同一磁盘：前面只写字母，最后一个带标注
                    prev_str = ", ".join(f"{l}:" for l in group[:-1])
                    last = group[-1]
                    parts.append(
                        f"{prev_str}, {last}:[磁盘{dn}{bus_str}, {label}]")
            else:
                # 无磁盘号的盘符（虚拟盘、网络盘等）
                parts.append(f"{ch}:[{labels.get(ch, '未知')}]")
                i += 1

        return ", ".join(parts)

    def _build_dropdown_groups(self, all_letters, min_letter, labels,
                                letter_to_disk, disk_map, all_bus):
        """
        构建下拉菜单的分组项。

        同一磁盘的所有分区合并为一个菜单项，例如：
          G:  [USB硬盘]
          H:, I:  [USB硬盘] 磁盘4
          T:  [虚拟盘]
          Z:  [虚拟盘]

        参数:
            all_letters:   所有已检测到的盘符字母（已排序）
            min_letter:    最小显示盘符（'D' 或 'G'）
            labels:        {letter: label} 分类标签
            letter_to_disk: {letter: disk_number} 盘符→磁盘号映射
            disk_map:      {disk_number: [letters]} 磁盘号→分区映射
            all_bus:       {letter: bus_type} 总线类型

        返回: (display_values_list, {display_string: primary_letter})
        """
        # 筛选 >= min_letter 的盘符
        filtered = [ch for ch in all_letters if ch >= min_letter]

        # 按磁盘分组（同一磁盘只出现一次）
        groups = []
        seen_disks = set()

        for ch in filtered:
            dn = letter_to_disk.get(ch)

            if dn is not None:
                if dn in seen_disks:
                    continue  # 该磁盘已被加入分组
                seen_disks.add(dn)
                # 获取该磁盘上所有 >= min_letter 的盘符
                disk_letters = sorted(
                    l for l in disk_map.get(dn, []) if l >= min_letter)
                if not disk_letters:
                    continue
                groups.append({
                    'letters': disk_letters,
                    'disk_number': dn,
                    'label': labels.get(disk_letters[0], '未知'),
                    'bus': all_bus.get(disk_letters[0], 'Unknown'),
                })
            else:
                # 无磁盘号的盘符（虚拟盘、网络盘等）各自独立一组
                groups.append({
                    'letters': [ch],
                    'disk_number': None,
                    'label': labels.get(ch, '未知'),
                    'bus': all_bus.get(ch, 'Unknown'),
                })

        # 按首个盘符字母排序
        groups.sort(key=lambda g: g['letters'][0])

        # 构建显示值和主盘符映射
        values = []
        combo_map = {}  # {display_string: 主盘符（第一个字母）}

        for g in groups:
            letters = g['letters']
            label = g['label']
            dn = g['disk_number']

            if len(letters) == 1:
                # 单分区或无磁盘号
                display = f"{letters[0]}:  [{label}]"
            else:
                # 多分区同磁盘
                letters_str = ", ".join(f"{l}:" for l in letters)
                display = f"{letters_str}  [{label}] 磁盘{dn}"

            values.append(display)
            combo_map[display] = letters[0]

        return values, combo_map

    def _apply_bus_detection(self, log_line, dropdown_values, combo_map,
                              disk_map, letter_to_disk, all_bus, labels):
        """在 UI 线程中应用后台识别结果，更新缓存和下拉菜单"""
        self._detecting = False

        # 更新所有缓存
        self._disk_map = disk_map
        self._letter_to_disk = letter_to_disk
        self._bus_cache = all_bus
        self._labels = labels
        self._combo_to_primary = combo_map

        # ── 更新下拉菜单（尝试保留用户之前的选择）──
        cur_primary = ""
        if self.drive_var.get():
            cur_primary = self._combo_to_primary.get(
                self.drive_var.get(),
                self.drive_var.get().split()[0].rstrip(':,').rstrip(':').upper()
            )
        self.combo["values"] = dropdown_values
        # 匹配之前选中的盘符
        hit = ""
        for v in dropdown_values:
            primary = combo_map.get(v, "")
            if primary == cur_primary:
                hit = v
                break
        if not hit and dropdown_values:
            hit = dropdown_values[0]
        self.drive_var.set(hit)

        # ── 输出统一日志行（仅输出结果，不输出启发式详情）──
        self.log_msg(f"[识别完成] 盘符：{log_line}")
        self.log_msg("")

        # ── 更新状态栏 ──
        if not dropdown_values:
            self.status_lbl.config(text="!! 未检测到可用盘符", foreground="red")
        else:
            self._on_drive_selected()

    # ════════════════════════════════════════════════════════════
    #  ★ v3.7: 从缓存重建下拉菜单（DEF 切换时使用）
    # ════════════════════════════════════════════════════════════

    def _rebuild_dropdown_from_cache(self):
        """
        ★ v3.7 新增：从已缓存的识别结果重建下拉菜单，不重新执行后台检测。
        用于 DEF 盘开关切换时，避免不必要的重复检测（检测需要约 2-3 秒）。

        前提条件：self._labels 等缓存已由后台识别填充。
        """
        min_letter = 'D' if self.show_def_var.get() else 'G'
        # 从缓存获取所有已识别的盘符
        all_letters = sorted(self._labels.keys())

        # 重建下拉菜单
        dropdown_values, combo_map = self._build_dropdown_groups(
            all_letters, min_letter, self._labels,
            self._letter_to_disk, self._disk_map, self._bus_cache)

        # 尝试保留之前的选择
        cur_primary = ""
        if self.drive_var.get():
            cur_primary = self._combo_to_primary.get(
                self.drive_var.get(),
                self.drive_var.get().split()[0].rstrip(':,').rstrip(':').upper()
            )

        # 更新映射和下拉框
        self._combo_to_primary = combo_map
        self.combo["values"] = dropdown_values

        # 恢复之前的选择
        hit = ""
        for v in dropdown_values:
            primary = combo_map.get(v, "")
            if primary == cur_primary:
                hit = v
                break
        if not hit and dropdown_values:
            hit = dropdown_values[0]
        self.drive_var.set(hit)

        # 更新状态栏
        if not dropdown_values:
            self.status_lbl.config(text="!! 未检测到可用盘符", foreground="red")
        else:
            self._on_drive_selected()

    # ════════════════════════════════════════════════════════════
    #  盘符选择 / 状态显示
    # ════════════════════════════════════════════════════════════

    def _get_bus_type(self, letter):
        """获取盘符的总线类型（优先从缓存读取，缓存未命中则实时查询）"""
        letter = letter.rstrip(":\\").upper()
        if letter in self._bus_cache:
            return self._bus_cache[letter]
        bus = get_bus_type_ioctl(letter)
        if bus == "Unknown":
            bus = get_drive_bus_type(letter)
        self._bus_cache[letter] = bus
        return bus

    def _toggle_def(self):
        """
        切换 D/E/F 盘显示开关。
        ★ v3.7: 如果已有缓存识别结果，直接从缓存重建下拉菜单，
        不再触发完整的后台重新检测。
        """
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

        # ★ v3.7: 如果有缓存的识别结果，直接重建下拉菜单（不重新检测）
        if self._labels:
            self._rebuild_dropdown_from_cache()
        else:
            # 尚未完成首次识别，需要执行完整刷新
            self.refresh()

    def _on_drive_selected(self, event=None):
        """下拉菜单选择变化时更新状态栏信息"""
        if self._busy:
            return
        v = self.drive_var.get()
        if not v:
            return
        if self._detecting:
            self.status_lbl.config(text="正在识别磁盘类型...", foreground="blue")
            return

        # 从映射获取主盘符
        primary = self._combo_to_primary.get(v)
        if primary is None:
            primary = v.split()[0].rstrip(':,').rstrip(':').upper()
        letter = primary.upper()

        label = self._labels.get(letter, "")

        # 同磁盘分区信息（用于状态栏显示）
        sibling_info = ""
        if letter in self._letter_to_disk:
            dn = self._letter_to_disk[letter]
            all_parts = sorted(self._disk_map.get(dn, []))
            if len(all_parts) > 1:
                max_show = 6
                if len(all_parts) <= max_show:
                    parts_str = ",".join(f"{p}:" for p in all_parts)
                else:
                    shown = all_parts[:max_show]
                    parts_str = (
                        ",".join(f"{p}:" for p in shown)
                        + f"…共{len(all_parts)}个"
                    )
                sibling_info = f" | 磁盘{dn} ({parts_str})"

        # ★ v3.7: 根据标签类型设置状态栏（虚拟盘使用通用描述）
        if label == "虚拟盘":
            self.status_lbl.config(
                text=f"⚠ 虚拟磁盘 谨慎操作{sibling_info}",
                foreground="#8B4513",
            )
        elif label == "固定":
            self.status_lbl.config(
                text=f"⚠ 固定硬盘 谨慎操作{sibling_info}", foreground="red")
        elif label == "网络":
            self.status_lbl.config(
                text="⚠ 网络硬盘 谨慎操作", foreground="#b34700")
        elif label == "光驱":
            self.status_lbl.config(text="光驱", foreground="gray")
        elif label == "USB硬盘":
            self.status_lbl.config(
                text=f"USB 移动硬盘{sibling_info}", foreground="green")
        elif label == "可移动":
            self.status_lbl.config(
                text=f"可移动设备{sibling_info}", foreground="green")
        else:
            self.status_lbl.config(
                text=sibling_info.lstrip(" |") if sibling_info else "",
                foreground="gray",
            )

    def _check_drive_safety(self, d):
        """
        弹出前的安全检查——对非 USB 设备显示相应的警告对话框。
        对虚拟磁盘、本地固定硬盘、网络驱动器、光驱分别给出不同级别的提示。
        返回: True（用户确认继续），False（用户取消）
        """
        letter = d[0]
        dt = get_drive_type_code(d)
        bus = self._get_bus_type(letter)
        is_usb = bus.upper() in USB_BUS_TYPES
        is_virtual = bus.upper() in VIRTUAL_BUS_TYPES

        # USB 固定磁盘：安全，直接放行
        if dt == 3 and is_usb:
            return True

        # ★ v3.7: 虚拟磁盘警告（使用通用描述，不假定具体工具）
        if (dt == 3 or dt == 2) and is_virtual:
            msg = (
                f"⚠ 安全提示\n\n"
                f"{d}\\ 被识别为【虚拟磁盘】，\n"
                f"可能是虚拟硬盘（VHD/VHDX）或第三方工具映射的虚拟存储。\n\n"
                f"弹出虚拟磁盘可能导致：\n"
                f"• 虚拟磁盘工具需要重新挂载\n"
                f"• 正在访问的文件或程序中断\n\n"
                f"是否仍要继续？"
            )
            return messagebox.askyesno("⚠ 安全提示", msg, icon="warning")

        # 本地固定硬盘（非 USB、非虚拟）：严重警告
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

        # 网络驱动器
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

        # 光驱
        elif dt == 5:
            msg = f"提示：{d}\\ 是光驱，弹出将打开光驱托盘。\n\n继续？"
            return messagebox.askyesno("提示", msg)

        return True

    def _check_offline_on_start(self):
        """启动时异步检查是否有脱机磁盘"""
        threading.Thread(target=self._do_check_offline_start, daemon=True).start()

    def _do_check_offline_start(self):
        """后台检查脱机磁盘并在日志中提示用户"""
        offline = get_offline_disks()
        if offline:
            usb_offline = [d for d in offline if d["bus"] in ("USB", "USB3")]
            if usb_offline:
                info_lines = []
                for d in usb_offline:
                    info_lines.append(
                        f"  磁盘 {d['number']}: {d['name']} "
                        f"({d['size_gb']:.1f} GB, {d['bus']})")
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
                    f"（非USB），可能为正常状态。\n")

    # ════════════════════════════════════════════════════════════
    #  盘符获取 / 刷新
    # ════════════════════════════════════════════════════════════

    def get_drive(self):
        """
        获取当前选择的主盘符（如 "G:"）。
        支持从分组下拉菜单中提取主盘符（取多分区组的第一个字母）。
        """
        v = self.drive_var.get()
        if not v:
            messagebox.showwarning("提示", "请先选择一个盘符")
            return None
        # 从映射中获取主盘符，或从显示值解析
        primary = self._combo_to_primary.get(v)
        if primary is None:
            primary = v.split()[0].rstrip(':,').rstrip(':').upper()
        d = f"{primary}:"
        if not drive_exists(d):
            messagebox.showwarning("提示", f"{d}\\ 不可访问")
            return None
        return d

    def refresh(self):
        """刷新盘符列表并重新执行完整的后台识别"""
        min_letter = 'D' if self.show_def_var.get() else 'G'
        # 先用快速扫描占位
        drives = get_drives_fast(min_letter)
        values = [f"{d[0]}  [{d[1]}]" for d in drives]
        self.combo["values"] = values
        cur = self.drive_var.get().split()[0].rstrip(':,') if self.drive_var.get() else ""
        hit = next((v for v in values if v.startswith(cur)),
                   values[0] if values else "")
        self.drive_var.set(hit)
        ds = ", ".join(f"{d[0]}[{d[1]}]" for d in drives) if drives else "无"
        self.log_msg(f"[刷新] 快速扫描：{ds}（正在识别类型...）")
        if not drives:
            self.status_lbl.config(text="!! 未检测到可用盘符", foreground="red")
        else:
            self.status_lbl.config(text="正在识别磁盘类型...", foreground="blue")
        # 启动后台完整识别
        self._start_bus_detection()
        # 同时检查脱机磁盘
        threading.Thread(target=self._do_check_offline_start, daemon=True).start()

    # ════════════════════════════════════════════════════════════
    #  日志 / 命令执行 / 线程管理
    # ════════════════════════════════════════════════════════════

    def log_msg(self, msg):
        """向日志区域追加一行消息"""
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.root.update_idletasks()

    def exec_cmd(self, cmd, timeout=120):
        """执行命令并将输出写入日志，返回 (返回码, 标准输出, 标准错误)"""
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
        """在后台线程中执行操作（防止 UI 冻结，同时阻止重复操作）"""
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

    # ════════════════════════════════════════════════════════════
    #  脱机磁盘恢复
    # ════════════════════════════════════════════════════════════

    def recover_offline(self):
        """恢复脱机磁盘（需要管理员权限）"""
        if not self._require_admin("恢复脱机磁盘"):
            return
        self.run_in_thread(self._do_recover_offline)

    def _do_recover_offline(self):
        """后台执行脱机磁盘恢复"""
        self.log_msg("\n--- 检测脱机磁盘 ---")
        offline = get_offline_disks()
        if not offline:
            self.log_msg("  未检测到脱机磁盘，一切正常。\n")
            return
        self.log_msg(f"  发现 {len(offline)} 个脱机磁盘：")
        for d in offline:
            self.log_msg(
                f"    磁盘 {d['number']}: {d['name']} "
                f"({d['size_gb']:.1f} GB, 总线: {d['bus']})")
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

    # ════════════════════════════════════════════════════════════
    #  磁盘号获取 / USB 安全移除
    # ════════════════════════════════════════════════════════════

    def _get_disk_number(self, letter):
        """
        获取盘符对应的物理磁盘号。
        查询优先级：缓存 → IOCTL → PowerShell
        """
        letter = letter.rstrip(":\\").upper()
        # 优先从缓存获取
        if letter in self._letter_to_disk:
            return self._letter_to_disk[letter]
        # IOCTL 查询
        num = get_disk_number_ioctl(letter)
        if num is not None:
            return num
        # 回退到 PowerShell
        rc, out, _ = run_cmd(
            f'powershell -NoProfile -Command "'
            f"(Get-Partition -DriveLetter {letter} -ErrorAction Stop).DiskNumber"
            f'"', timeout=10)
        if rc == 0 and out.strip().isdigit():
            return int(out.strip())
        return None

    def _usb_safe_remove(self, disk_number):
        """
        通过 CfgMgr32.dll 的 CM_Request_Device_Eject 执行 USB 硬件级安全移除。
        这会向 USB 控制器发送停止设备命令，使硬盘停止转动。
        """
        ps_file = os.path.join(os.environ.get("TEMP", "."), "_usb_eject.ps1")
        try:
            with open(ps_file, "w", encoding="utf-8-sig") as f:
                f.write(USB_EJECT_PS1)
            cmd = (
                f'powershell -NoProfile -ExecutionPolicy Bypass '
                f'-File "{ps_file}" -DiskNumber {disk_number}')
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
        """
        刷新卷的写缓冲区（防止数据丢失）。
        在弹出前调用，确保所有待写数据已物理写入磁盘。
        """
        letter = d.rstrip(":\\")
        volume = f"\\\\.\\{letter}:"
        k32 = ctypes.windll.kernel32
        k32.CreateFileW.restype = ctypes.c_void_p
        INVALID_HANDLE = ctypes.c_void_p(-1).value
        SHARE_RW = 0x3
        OPEN_EXISTING = 3
        # 尝试以写权限打开（优先），失败则以读权限
        h = k32.CreateFileW(volume, 0xC0000000, SHARE_RW, None, OPEN_EXISTING, 0, None)
        if h is None or h == INVALID_HANDLE:
            h = k32.CreateFileW(volume, 0x80000000, SHARE_RW, None, OPEN_EXISTING, 0, None)
        if h is None or h == INVALID_HANDLE:
            self.log_msg(f"    [注意] 无法打开 {letter}: 卷句柄，跳过刷缓冲区")
            return
        ok = k32.FlushFileBuffers(h)
        k32.CloseHandle(h)
        if ok:
            self.log_msg(f"    [OK] {letter}: 缓冲区已刷新")
        else:
            self.log_msg(f"    [注意] {letter}: FlushFileBuffers 返回失败，"
                         f"缓冲区可能未完全刷新")

    # ════════════════════════════════════════════════════════════
    #  弹出逻辑（五种方法逐一尝试）
    # ════════════════════════════════════════════════════════════

    def _try_eject(self, d):
        """
        尝试弹出磁盘（五种方法逐一尝试，任一成功即停止）：
          方法1: CM_Request_Device_Eject — USB 硬件级安全移除，能停转硬盘
          方法2: DeviceIoControl IOCTL_STORAGE_EJECT_MEDIA — 内核级弹出
          方法3: Shell.Application Eject — 调用资源管理器弹出
          方法4: Set-Disk -IsOffline — 设为脱机 + USB 安全移除
          方法5: diskpart remove all dismount — 最后手段

        对于多分区设备，所有分区都会被统一弹出。
        """
        letter = d[0]
        self.log_msg("\n  获取磁盘信息...")
        disk_number = self._get_disk_number(letter)
        all_partitions = []
        sibling_letters = []
        is_multi = False
        if disk_number is not None:
            self.log_msg(f"    物理磁盘号: {disk_number}")
            all_partitions = get_all_partitions_on_disk(disk_number)
            sibling_letters = [
                p for p in all_partitions if p.upper() != letter.upper()]
            is_multi = len(all_partitions) > 1
            if is_multi:
                parts_str = ", ".join(f"{p}:" for p in all_partitions)
                self.log_msg(
                    f"    [多分区] 磁盘 {disk_number} 包含 "
                    f"{len(all_partitions)} 个分区: {parts_str}")
                self.log_msg(
                    f"    [多分区] 将对整个设备执行弹出（影响所有分区）")
        else:
            self.log_msg("    [!] 无法获取磁盘号，USB 安全移除可能不可用")

        # 保存弹出状态供后续使用（成功/失败提示需要这些信息）
        self._eject_disk_number = disk_number
        self._eject_all_partitions = list(all_partitions)
        self._eject_is_multi = is_multi

        # 辅助函数：检查所有分区盘符是否已消失
        def all_gone():
            for p in all_partitions:
                if drive_exists(f"{p}:"):
                    return False
            return True

        # 辅助函数：获取仍在挂载的分区列表
        def get_remaining():
            return [f"{p}:" for p in all_partitions if drive_exists(f"{p}:")]

        # 辅助函数：检查目标盘符是否已消失
        def target_gone():
            return not drive_exists(d)

        # 多分区时先刷新所有分区的写缓冲区
        if is_multi:
            self.log_msg("\n  刷新磁盘上所有分区的写缓冲区...")
            for p in all_partitions:
                if drive_exists(f"{p}:"):
                    self._flush_volume(f"{p}:")

        # ── 方法1: USB 安全移除（设备级，最理想的弹出方式）──
        if disk_number is not None:
            self.log_msg("\n  方法1: USB 安全移除 (CM_Request_Device_Eject) ...")
            usb_ok = self._usb_safe_remove(disk_number)
            if usb_ok:
                time.sleep(2)
                if is_multi:
                    if all_gone():
                        self.log_msg("    所有分区盘符已消失，硬盘已安全移除并停转!")
                        return True
                    else:
                        remaining = get_remaining()
                        self.log_msg(
                            f"    弹出指令已发送但部分盘符仍在"
                            f"（剩余: {', '.join(remaining)}），继续...")
                else:
                    if target_gone():
                        self.log_msg("    盘符已消失，硬盘已安全移除并停转!")
                        return True
                    self.log_msg("    弹出指令成功但盘符仍在，继续...")
            else:
                self.log_msg("    失败（可能有程序占用），尝试下一方法...")

        # ── 方法2: DeviceIoControl API（卷级弹出）──
        if is_multi:
            self.log_msg("\n  方法2: DeviceIoControl API（逐一弹出所有分区）...")
            for p in all_partitions:
                if drive_exists(f"{p}:"):
                    self.log_msg(f"    弹出 {p}: ...")
                    ok_v, msg_v, warnings_v = eject_volume_api(f"{p}:")
                    for w in warnings_v:
                        self.log_msg(f"      [注意] {w}")
                    self.log_msg(f"      {msg_v}")
        else:
            self.log_msg("\n  方法2: DeviceIoControl API ...")
            ok_v, msg_v, warnings_v = eject_volume_api(d)
            if warnings_v:
                for w in warnings_v:
                    self.log_msg(f"    [注意] {w}")
            self.log_msg(f"    {msg_v}")

        time.sleep(2)

        if is_multi:
            if all_gone():
                self.log_msg("    所有分区盘符已消失!")
                if disk_number is not None:
                    self.log_msg("    追加 USB 安全移除以停转硬盘...")
                    self._usb_safe_remove(disk_number)
                return True
            if disk_number is not None:
                remaining = get_remaining()
                self.log_msg(
                    f"    部分分区仍在（{', '.join(remaining)}），"
                    f"重试设备级弹出...")
                self._usb_safe_remove(disk_number)
                time.sleep(2)
                if all_gone():
                    self.log_msg("    所有分区盘符已消失，硬盘已安全移除!")
                    return True
            self.log_msg("    盘符仍在，尝试下一方法...")
        else:
            if target_gone():
                self.log_msg("    盘符已消失!")
                if disk_number is not None:
                    self.log_msg("    追加 USB 安全移除以停转硬盘...")
                    self._usb_safe_remove(disk_number)
                return True
            self.log_msg("    盘符仍在，尝试下一方法...")

        # ── 方法3: Shell.Application Eject（资源管理器级弹出）──
        self.log_msg("\n  方法3: Shell.Application Eject ...")
        cmd = (
            'powershell -NoProfile -Command "'
            "(New-Object -ComObject Shell.Application)"
            ".NameSpace(17).ParseName('" + letter + ":').InvokeVerb('Eject')"
            '"')
        self.exec_cmd(cmd, timeout=15)
        time.sleep(3)
        if is_multi:
            if all_gone():
                self.log_msg("    所有分区盘符已消失!")
                if disk_number is not None:
                    self.log_msg("    追加 USB 安全移除以停转硬盘...")
                    self._usb_safe_remove(disk_number)
                return True
        else:
            if target_gone():
                self.log_msg("    盘符已消失!")
                if disk_number is not None:
                    self.log_msg("    追加 USB 安全移除以停转硬盘...")
                    self._usb_safe_remove(disk_number)
                return True
        self.log_msg("    盘符仍在，尝试下一方法...")

        # ── 方法4: Set-Disk -IsOffline + USB 安全移除（需管理员）──
        if self._is_admin:
            self.log_msg("\n  方法4: Set-Disk -IsOffline + USB 安全移除 ...")
            self.log_msg(
                "    [注意] 此方法如果 USB 移除失败，"
                "下次插入可能需要手动恢复联机")
            if is_multi:
                self.log_msg("    刷新磁盘上所有分区的缓冲区...")
                for p in all_partitions:
                    if drive_exists(f"{p}:"):
                        self._flush_volume(f"{p}:")
            else:
                self.log_msg("    刷新卷缓冲区...")
                self._flush_volume(d)
            cmd = (
                'powershell -NoProfile -Command "'
                "$p = Get-Partition -DriveLetter " + letter
                + " -ErrorAction Stop; "
                "$dk = $p | Get-Disk; "
                'Set-Disk -Number $dk.Number -IsOffline $true"')
            self.exec_cmd(cmd, timeout=15)
            time.sleep(2)
            if target_gone():
                self.log_msg("    盘符已消失!")
                if disk_number is not None:
                    self.log_msg(
                        "    追加 USB 安全移除以停转硬盘并清除脱机标记...")
                    usb_ok = self._usb_safe_remove(disk_number)
                    if usb_ok:
                        self.log_msg("    硬盘已安全移除并停转!")
                    else:
                        self.log_msg("    [!] USB 安全移除失败")
                        self.log_msg(
                            "    [!] 数据已安全（文件系统已脱机），"
                            "但硬盘可能仍在转动")
                        self.log_msg(
                            "    [!] 建议等待约 5 秒再拔出 USB 线缆")
                        self.log_msg(
                            "    [!] 下次插入如无盘符，请点击【恢复脱机磁盘】")
                return True
            self.log_msg("    盘符仍在，尝试下一方法...")
        else:
            self.log_msg("\n  方法4: 跳过（需要管理员权限）")

        # ── 方法5: diskpart remove all dismount（最后手段，需管理员）──
        if self._is_admin:
            self.log_msg("\n  方法5: diskpart ...")
            if is_multi:
                self.log_msg("    刷新磁盘上所有分区的缓冲区...")
                for p in all_partitions:
                    if drive_exists(f"{p}:"):
                        self._flush_volume(f"{p}:")
                for p in all_partitions:
                    if drive_exists(f"{p}:"):
                        self.log_msg(f"    对 {p}: 执行 diskpart remove ...")
                        tmp = os.path.join(
                            os.environ.get("TEMP", "."), "_eject.txt")
                        with open(tmp, "w") as f:
                            f.write(
                                f"select volume {p}\nremove all dismount\n")
                        self.exec_cmd(f'diskpart /s "{tmp}"', timeout=30)
                        try:
                            os.remove(tmp)
                        except OSError:
                            pass
            else:
                self.log_msg("    刷新卷缓冲区...")
                self._flush_volume(d)
                tmp = os.path.join(
                    os.environ.get("TEMP", "."), "_eject.txt")
                with open(tmp, "w") as f:
                    f.write(
                        f"select volume {letter}\nremove all dismount\n")
                self.exec_cmd(f'diskpart /s "{tmp}"', timeout=30)
                try:
                    os.remove(tmp)
                except OSError:
                    pass

            time.sleep(2)

            if is_multi:
                if all_gone():
                    self.log_msg("    所有分区盘符已消失!")
                    if disk_number is not None:
                        self.log_msg("    追加 USB 安全移除以停转硬盘...")
                        self._usb_safe_remove(disk_number)
                    return True
                if disk_number is not None:
                    remaining = get_remaining()
                    if remaining:
                        self.log_msg(
                            f"    部分分区仍在（{', '.join(remaining)}），"
                            f"重试USB安全移除...")
                        self._usb_safe_remove(disk_number)
                        time.sleep(2)
                        if all_gone():
                            self.log_msg(
                                "    所有分区盘符已消失，硬盘已安全移除!")
                            return True
            else:
                if target_gone():
                    self.log_msg("    盘符已消失!")
                    if disk_number is not None:
                        self.log_msg("    追加 USB 安全移除以停转硬盘...")
                        self._usb_safe_remove(disk_number)
                    return True
        else:
            self.log_msg("\n  方法5: 跳过（需要管理员权限）")

        # 所有方法都失败
        if is_multi:
            remaining = get_remaining()
            if remaining:
                self.log_msg(
                    f"\n  [多分区] 仍在挂载的分区: {', '.join(remaining)}")
                self.log_msg(
                    "  [!!] 警告：请勿拔出硬盘！上述分区可能仍有数据读写。")

        return False

    # ════════════════════════════════════════════════════════════
    #  检测占用 / 停止服务 / 弹出操作
    # ════════════════════════════════════════════════════════════

    def detect(self):
        """检测选中盘符的占用情况"""
        d = self.get_drive()
        if not d:
            return
        self.run_in_thread(lambda: self._detect(d))

    def _detect(self, d):
        """后台执行磁盘占用检测"""
        self.log_msg(f"\n{'='*50}")
        self.log_msg(f"  检测占用 {d}\\ 的进程和服务")
        self.log_msg(f"{'='*50}")
        if not self._is_admin:
            self.log_msg("  [注意] 当前为普通模式，部分检测功能可能受限")
        bus = self._get_bus_type(d[0])
        dt = get_drive_type_code(d)
        # 使用缓存标签
        type_name = self._labels.get(d[0], DRIVE_TYPE_MAP.get(dt, "未知"))
        self.log_msg(f"\n  磁盘类型: {type_name}  总线: {bus}")

        # 显示多分区信息
        disk_number = self._get_disk_number(d[0])
        if disk_number is not None:
            all_parts = get_all_partitions_on_disk(disk_number)
            if len(all_parts) > 1:
                parts_str = ", ".join(f"{p}:" for p in all_parts)
                self.log_msg(
                    f"  [多分区] 磁盘 {disk_number} 包含 "
                    f"{len(all_parts)} 个分区: {parts_str}")

        # [1] 在该盘上运行的进程
        self.log_msg("\n[1] 在该盘上运行的进程：")
        cmd = (
            'powershell -NoProfile -Command "'
            "Get-Process | Where-Object { $_.Path -like '"
            + d + "\\*' } | "
            'Format-Table Id,Name,Path -AutoSize | Out-String -Width 300"')
        rc, out, _ = self.exec_cmd(cmd)
        if not out.strip():
            self.log_msg("    （无）")

        # [2] 加载了该盘文件的进程
        self.log_msg("\n[2] 加载了该盘文件的进程：")
        ps = (
            'powershell -NoProfile -Command "Get-Process | ForEach-Object { $p=$_; try { '
            "$p.Modules | Where-Object { $_.FileName -like '"
            + d + "\\*' } | "
            "ForEach-Object { Write-Output ('PID={0}  {1}  {2}' -f "
            "$p.Id,$p.Name,$_.FileName) }"
            ' } catch {} }"')
        rc, out, _ = self.exec_cmd(ps)
        if not out.strip():
            self.log_msg("    （无）")

        # [3] 常见占用服务状态
        self.log_msg("\n[3] 常见占用服务状态：")
        for name, display in SERVICES:
            st = svc_status(name)
            if st == "running":
                self.log_msg(f"    * {display} ({name})  ->  运行中 !!")
            elif st == "stopped":
                self.log_msg(f"    o {display} ({name})  ->  已停止")
            else:
                self.log_msg(f"    - {display} ({name})  ->  未安装")

        # [4] openfiles 查询
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

    def stop_svc(self):
        """一键停止占用服务"""
        if not self._require_admin("停止系统服务"):
            return
        self.run_in_thread(self._stop_svc)

    def _stop_svc(self):
        """后台执行停止服务"""
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

    def smart_eject(self):
        """安全弹出（推荐）：停止服务 → 弹出 → 恢复服务"""
        d = self.get_drive()
        if not d:
            return
        if not self._require_admin("安全弹出（需停止服务）"):
            return
        if not self._check_drive_safety(d):
            self.log_msg(f"[取消] 用户取消了对 {d} 的弹出操作\n")
            return
        letter = d[0]
        disk_number = self._get_disk_number(letter)
        multi_partition_warning = ""
        if disk_number is not None:
            all_parts = get_all_partitions_on_disk(disk_number)
            if len(all_parts) > 1:
                parts_str = ", ".join(f"{p}:" for p in all_parts)
                multi_partition_warning = (
                    f"\n⚠ 该设备（磁盘 {disk_number}）包含 "
                    f"{len(all_parts)} 个分区：{parts_str}\n"
                    f"弹出操作将移除整个设备上的所有分区！\n")
        msg = (
            f"将执行以下步骤：\n\n"
            f"1. 停止常见占用服务\n"
            f"2. 多种方式尝试弹出 {d}（含 USB 硬件级安全移除）\n"
            f"3. 恢复服务\n"
            f"{multi_partition_warning}\n继续？")
        if not messagebox.askyesno("安全弹出", msg):
            return
        self.run_in_thread(lambda: self._smart_eject(d))

    def _smart_eject(self, d):
        """后台执行安全弹出全流程"""
        self.log_msg(f"\n{'='*50}")
        self.log_msg(f"  安全弹出 {d}")
        self.log_msg(f"{'='*50}")

        # 步骤1：停止占用服务
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

        # 步骤2：弹出硬盘
        self.log_msg("\n步骤 2：弹出硬盘（逐一尝试多种方法）...")
        ok = self._try_eject(d)
        if ok:
            if self._eject_is_multi:
                parts_str = ", ".join(
                    f"{p}:" for p in self._eject_all_partitions)
                self.log_msg(
                    f"\n[OK] 磁盘 {self._eject_disk_number}"
                    f"（{parts_str}）已成功弹出！")
                self.log_msg("     所有分区已安全移除，可以拔出硬盘。")
            else:
                self.log_msg(f"\n[OK] {d} 已成功弹出！可以安全拔出硬盘。")
        else:
            if self._eject_is_multi:
                remaining = [
                    f"{p}:" for p in self._eject_all_partitions
                    if drive_exists(f"{p}:")]
                if remaining:
                    self.log_msg(
                        f"\n[!!] 弹出失败。仍在挂载的分区: "
                        f"{', '.join(remaining)}")
                    self.log_msg(
                        "     请勿拔出硬盘！请关闭占用这些分区的程序后重试。")
                else:
                    self.log_msg(f"\n[!!] {d} 弹出状态不确定，请检查。")
            else:
                self.log_msg(f"\n[!!] {d} 仍然存在，所有弹出方法均失败。")
                self.log_msg("     请关闭该盘上所有打开的文件/窗口后重试，")
                self.log_msg("     或使用【检测】功能查看哪些进程在占用。")

        # 步骤3：恢复服务
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
        """强制弹出（跳过停止服务）"""
        d = self.get_drive()
        if not d:
            return
        if not self._check_drive_safety(d):
            self.log_msg(f"[取消] 用户取消了对 {d} 的弹出操作\n")
            return
        letter = d[0]
        disk_number = self._get_disk_number(letter)
        multi_warning = ""
        if disk_number is not None:
            all_parts = get_all_partitions_on_disk(disk_number)
            if len(all_parts) > 1:
                parts_str = ", ".join(f"{p}:" for p in all_parts)
                multi_warning = (
                    f"\n\n⚠ 该设备（磁盘 {disk_number}）包含 "
                    f"{len(all_parts)} 个分区：{parts_str}\n"
                    f"弹出将移除所有分区！")
        if not self._is_admin:
            msg = (
                f"当前为普通模式，部分弹出方法（diskpart、Set-Disk 等）\n"
                f"将不可用，但仍可尝试 USB 安全移除等方法。\n\n"
                f"尝试弹出 {d}？\n\n"
                f"提示：如需完整功能，请点击右上角「提升为管理员」。"
                f"{multi_warning}")
        else:
            msg = f"跳过停止服务，直接弹出 {d}？{multi_warning}"
        if not messagebox.askyesno("强制弹出", msg):
            return
        self.run_in_thread(lambda: self._force_eject(d))

    def _force_eject(self, d):
        """后台执行强制弹出"""
        self.log_msg(f"\n--- 强制弹出 {d} ---")
        if not self._is_admin:
            self.log_msg("  [注意] 当前为普通模式，方法4/5（需管理员）将被跳过")
        ok = self._try_eject(d)
        if ok:
            if self._eject_is_multi:
                parts_str = ", ".join(
                    f"{p}:" for p in self._eject_all_partitions)
                self.log_msg(
                    f"\n[OK] 磁盘 {self._eject_disk_number}"
                    f"（{parts_str}）已弹出！")
                self.log_msg("     所有分区已安全移除，可以拔出硬盘。\n")
            else:
                self.log_msg(f"\n[OK] {d} 已弹出！可以安全拔出硬盘。\n")
            self.log_msg("正在刷新盘符列表...")
            time.sleep(1)
            self.root.after(0, self.refresh)
        else:
            if self._eject_is_multi:
                remaining = [
                    f"{p}:" for p in self._eject_all_partitions
                    if drive_exists(f"{p}:")]
                if remaining:
                    self.log_msg(
                        f"\n[!!] 弹出失败。仍在挂载的分区: "
                        f"{', '.join(remaining)}")
                    self.log_msg("     请勿拔出硬盘！")
                else:
                    self.log_msg(f"\n[!!] {d} 弹出状态不确定，请检查。")
            else:
                self.log_msg(f"\n[!!] {d} 仍然存在，弹出失败。")
            if not self._is_admin:
                self.log_msg("     建议以管理员身份运行后重试。")
            else:
                self.log_msg("     请关闭占用该盘的程序后重试。")
            self.log_msg("")

    # ════════════════════════════════════════════════════════════
    #  进阶功能：删除系统文件夹 / 权限管理
    # ════════════════════════════════════════════════════════════

    def _del_folder(self, d, name, takeown=False):
        """删除指定盘符上的系统文件夹（可选先获取所有权）"""
        path = f"{d}\\{name}"
        if not os.path.exists(path):
            self.log_msg(f"  {path} 不存在，跳过")
            return
        self.log_msg(f"\n--- 删除 {path} ---")
        if takeown:
            self.exec_cmd(f'takeown /f "{path}" /r /d y', timeout=180)
            self.exec_cmd(
                f'icacls "{path}" /grant administrators:F /t', timeout=180)
        self.exec_cmd(f'cmd /c rd /s /q "{path}"')
        if os.path.exists(path):
            self.log_msg(f"  [!!] {path} 可能未完全删除")
        else:
            self.log_msg(f"  [OK] {path} 已删除")

    def del_svi(self):
        """删除 System Volume Information 文件夹"""
        if not self._require_admin("删除系统文件夹"):
            return
        d = self.get_drive()
        if not d:
            return
        if not messagebox.askyesno(
            "确认", f"删除 {d}\\System Volume Information？"):
            return
        self.run_in_thread(lambda: self._do_del_svi(d))

    def _do_del_svi(self, d):
        self._del_folder(d, "System Volume Information", True)
        self.log_msg("")

    def del_rec(self):
        """删除 $RECYCLE.BIN 文件夹"""
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
        """一键删除两个系统文件夹"""
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

    def deny_write(self):
        """禁止 SYSTEM 账户对目标磁盘的写入权限"""
        if not self._require_admin("修改磁盘权限"):
            return
        d = self.get_drive()
        if not d:
            return
        msg = (
            f"禁止 SYSTEM 写入 {d}\\ ？\n\n"
            f"效果：系统无法在该盘自动创建文件夹\n"
            f"恢复：随时点击【恢复】按钮")
        if not messagebox.askyesno("确认", msg):
            return
        self.run_in_thread(lambda: self._deny(d))

    def _deny(self, d):
        self.log_msg(f"\n--- 禁止 SYSTEM 写入 {d}\\ ---")
        self.exec_cmd(f'icacls {d}\\ /deny "SYSTEM:(WD)" /T /C')
        self.log_msg("[OK] 已禁止\n")

    def allow_write(self):
        """恢复 SYSTEM 账户对目标磁盘的写入权限"""
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
    #  文件/文件夹占用检测
    # ════════════════════════════════════════════════════════════

    def _on_file_drop(self, files):
        """拖放回调：接收拖入的文件/文件夹路径并填入路径框"""
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
            self.notebook.select(1)  # 自动切换到文件占用标签页
        except Exception:
            pass
        self.log_msg(f"[拖放] 已接收路径: {path}")

    def browse_file(self):
        """打开文件选择对话框"""
        path = filedialog.askopenfilename(title="选择要检测占用的文件")
        if path:
            self.file_path_var.set(os.path.normpath(path))

    def browse_folder(self):
        """打开文件夹选择对话框"""
        path = filedialog.askdirectory(title="选择要检测占用的文件夹")
        if path:
            self.file_path_var.set(os.path.normpath(path))

    @staticmethod
    def _drive_letter_of(path):
        """从路径中提取盘符字母（如 C），用于后续按盘符查询占用"""
        normed = os.path.normpath(path)
        if len(normed) >= 2 and normed[1] == ':':
            return normed[0].upper()
        return None

    def detect_file_lock(self):
        """检测文件/文件夹的占用进程和服务"""
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
        """后台执行文件占用检测（四步检测）"""
        is_dir = os.path.isdir(path)
        self._file_lock_path = path
        self.log_msg(f"\n{'='*50}")
        self.log_msg(f"  检测占用: {path}")
        self.log_msg(f"  类型: {'文件夹' if is_dir else '文件'}")
        self.log_msg(f"{'='*50}")
        all_procs = {}  # {pid: {pid, name, detail}} 去重收集

        # ── [1] Restart Manager API 检测 ──
        self.log_msg("\n[1] Restart Manager API 检测:")
        if is_dir:
            self.log_msg("    收集目录内文件（最多300个，深度3层）...")
            files = collect_files_in_dir(path, max_files=300, max_depth=3)
            self.log_msg(f"    收集到 {len(files)} 个文件")
            if files:
                # 分批处理避免 RM 资源限制
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
                    f"{info['name']:<24}  {info['detail']}")
        else:
            self.log_msg("    （未检测到）")

        # ── [2] PowerShell 进程路径/模块检测 ──
        if is_dir:
            self.log_msg("\n[2] PowerShell 进程路径/模块检测:")
            escaped = path.replace("'", "''").rstrip('\\')
            ps_cmd = (
                'powershell -NoProfile -Command "Get-Process | ForEach-Object { '
                '$p=$_; $found=$false; '
                "if ($p.Path -and ($p.Path -like '"
                + escaped + "\\*')) { $found=$true }; "
                'if (-not $found) { try { '
                "$p.Modules | ForEach-Object { "
                "if ($_.FileName -like '" + escaped + "\\*') { $found=$true } } "
                '} catch {} }; '
                'if ($found) { '
                'Write-Output ("{0}|{1}|{2}" -f $p.Id,$p.Name,$p.Path) '
                '} }"')
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
                                f"    PID={pid:<6}  {pname:<24}  {ppath}")
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
                "if ($p.Path -and ($p.Path -eq '" + escaped
                + "')) { $tag='EXE' }; "
                "if (-not $tag) { try { "
                "$p.Modules | ForEach-Object { "
                "if ($_.FileName -eq '" + escaped
                + "') { $tag='MOD' } } "
                "} catch {} }; "
                "if ($tag) { "
                "Write-Output ('{0}|{1}|{2}|{3}' -f "
                "$tag,$p.Id,$p.Name,$p.Path) "
                '} }"')
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
                                f"    PID={pid:<6}  {pname:<24}  {detail}")
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

        # ── [3] 常见占用服务状态 ──
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

        # ── [4] openfiles 查询 ──
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
                        "    不可用（需先运行 openfiles /local on 并重启）")

        # ── 汇总结果 ──
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
        """一键停止所有占用（服务+进程）"""
        has_proc = bool(self._file_lock_processes)
        has_svc = bool(self._file_lock_services)
        if not has_proc and not has_svc:
            messagebox.showinfo(
                "提示",
                "没有检测到占用进程或运行中的服务。\n请先点击【检测占用】。")
            return
        # 如果是从提权恢复的结果，先验证有效性
        if self._detection_is_restored:
            alive = [p for p in self._file_lock_processes
                     if is_process_alive(p["pid"])]
            self._file_lock_processes = alive
            has_proc = bool(alive)
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
                    "如仍有问题，请重新点击【检测占用】。")
                self._detection_is_restored = False
                return
        # 停止服务需要管理员权限
        if has_svc and not self._require_admin("停止占用服务并结束进程"):
            return
        # 构建确认对话框内容
        lines = []
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
                    f"  ... 还有 {len(self._file_lock_processes)-15} 个")
            lines.append("")
        lines.append("⚠ 未保存的数据可能丢失！确定继续？")
        msg = "\n".join(lines)
        if not messagebox.askyesno("确认停止所有占用", msg, icon="warning"):
            return
        self._detection_is_restored = False
        self.run_in_thread(self._do_kill_all_file_lock)

    def _do_kill_all_file_lock(self):
        """后台执行：停止服务 + 结束进程"""
        self.log_msg(f"\n{'='*50}")
        self.log_msg("  一键停止所有占用")
        self.log_msg(f"{'='*50}")
        # 步骤1：停止占用服务
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
        # 步骤2：结束占用进程
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
        # 清空检测状态
        self._file_lock_processes = []
        self._file_lock_services = {}
        self.log_msg(f"\n[OK] 操作完成：停止 {svc_stopped} 个服务"
                     f"，结束 {killed} 个进程\n")
        # 提示用户是否删除文件
        saved_path = self._file_lock_path
        if saved_path and os.path.exists(saved_path):
            self.root.after(
                300,
                lambda p=saved_path: self._prompt_delete_after_unlock(p))
        elif self._all_stopped_services:
            self.root.after(
                0,
                lambda: self._show_service_restore_dialog(auto_popup=True))

    def _prompt_delete_after_unlock(self, path):
        """占用解除后提示用户是否将文件移到回收站"""
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
            f"• 【否】→ 保留不删除")
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


# ══════════════════════════════════════════════════════════════
#   程序入口
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    try:
        App()
    except Exception as e:
        messagebox.showerror("启动失败", str(e))