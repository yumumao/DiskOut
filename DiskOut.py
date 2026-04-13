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

SERVICES = [
    ("WSearch",       "Windows Search"),
    ("SysMain",       "SysMain"),
    ("VSS",           "Volume Shadow Copy"),
    ("defragsvc",     "Optimize Drives"),
    ("WMPNetworkSvc", "WMP Network Sharing"),
    ("StorSvc",       "Storage Service"),
]

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


def get_drives():
    drives = []
    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    type_map = {2: "可移动", 3: "固定", 4: "网络", 5: "光驱", 6: "RAM"}
    for i, ch in enumerate(string.ascii_uppercase):
        if ch < "G":
            continue
        if bitmask & (1 << i):
            path = f"{ch}:\\"
            dt = ctypes.windll.kernel32.GetDriveTypeW(path)
            label = type_map.get(dt, "未知")
            drives.append((f"{ch}:", label))
    return drives


def get_offline_disks():
    """获取所有脱机状态的磁盘"""
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
                        "number": num,
                        "name": name,
                        "size_gb": size_gb,
                        "bus": bus,
                    })
                except (ValueError, IndexError):
                    pass
    return disks


def set_disk_online(disk_number):
    """将脱机磁盘恢复联机"""
    cmd = (
        f'powershell -Command "'
        f'Set-Disk -Number {disk_number} -IsOffline $false"'
    )
    rc, out, err = run_cmd(cmd, timeout=15)
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


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("移动硬盘清理工具 - 管理员模式")
        self.root.geometry("680x780")
        self.root.minsize(600, 680)
        self._busy = False
        self._svc_was_running = {}
        self.build_ui()
        self.root.after(500, self._check_offline_on_start)
        self.root.mainloop()

    def build_ui(self):
        m = ttk.Frame(self.root, padding=8)
        m.pack(fill="both", expand=True)

        f0 = ttk.LabelFrame(m, text="盘符选择（仅 G: 及之后）", padding=8)
        f0.pack(fill="x", pady=(0, 4))

        drives = get_drives()
        values = [f"{d[0]}  [{d[1]}]" for d in drives]
        self.drive_var = tk.StringVar(value=values[0] if values else "")
        self.combo = ttk.Combobox(
            f0, textvariable=self.drive_var,
            values=values, state="readonly",
            width=18, font=("Consolas", 11)
        )
        self.combo.pack(side="left", padx=(0, 8))
        ttk.Button(f0, text="刷新盘符", command=self.refresh).pack(side="left")
        self.status_lbl = ttk.Label(f0, text="", foreground="gray")
        self.status_lbl.pack(side="left", padx=10)
        if not drives:
            self.status_lbl.config(
                text="!! 未检测到 G: 及之后的盘符", foreground="red"
            )

        nb = ttk.Notebook(m)
        nb.pack(fill="x", pady=4)

        # Tab 1: 解除占用 / 弹出
        t1 = ttk.Frame(nb, padding=6)
        nb.add(t1, text=" 解除占用 / 弹出 ")

        ttk.Button(t1, text="[检测] 检测占用该盘的进程和服务",
                   command=self.detect).pack(fill="x", pady=2)
        ttk.Button(t1, text="[停止] 一键停止常见占用服务",
                   command=self.stop_svc).pack(fill="x", pady=2)
        ttk.Button(t1, text="[恢复] 恢复已停止的服务",
                   command=self.start_svc).pack(fill="x", pady=2)

        ttk.Separator(t1).pack(fill="x", pady=6)

        ttk.Button(t1, text="[安全弹出] 停止服务 + 弹出硬盘（推荐）",
                   command=self.smart_eject).pack(fill="x", pady=2)
        ttk.Button(t1, text="[强制弹出] 直接弹出硬盘",
                   command=self.force_eject).pack(fill="x", pady=2)

        ttk.Separator(t1).pack(fill="x", pady=6)

        self.online_btn = ttk.Button(
            t1, text="[恢复脱机磁盘] 检测并恢复被标记为脱机的磁盘",
            command=self.recover_offline
        )
        self.online_btn.pack(fill="x", pady=2)

        # Tab 2: 删除系统文件夹
        t2 = ttk.Frame(nb, padding=6)
        nb.add(t2, text=" 删除系统文件夹 ")

        ttk.Button(t2, text="[删除] System Volume Information",
                   command=self.del_svi).pack(fill="x", pady=2)
        ttk.Button(t2, text="[删除] $RECYCLE.BIN",
                   command=self.del_rec).pack(fill="x", pady=2)
        ttk.Button(t2, text="[一键删除] 以上两个文件夹",
                   command=self.del_both).pack(fill="x", pady=2)

        # Tab 3: SYSTEM 写入权限
        t3 = ttk.Frame(nb, padding=6)
        nb.add(t3, text=" SYSTEM 写入权限 ")

        ttk.Button(t3, text="[禁止] 禁止 SYSTEM 写入",
                   command=self.deny_write).pack(fill="x", pady=2)
        ttk.Button(t3, text="[恢复] 恢复 SYSTEM 写入",
                   command=self.allow_write).pack(fill="x", pady=2)

        hint_text = (
            "提示：禁止写入后，系统服务将无法在该盘创建任何文件。"
            "\n如需恢复，请在拔盘前点击【恢复】按钮。"
        )
        ttk.Label(t3, foreground="gray", text=hint_text).pack(
            anchor="w", pady=(8, 0)
        )

        # 日志
        f4 = ttk.LabelFrame(m, text="执行日志", padding=4)
        f4.pack(fill="both", expand=True, pady=(4, 0))

        self.log = scrolledtext.ScrolledText(
            f4, height=14, font=("Consolas", 9), wrap=tk.WORD
        )
        self.log.pack(fill="both", expand=True)
        ttk.Button(f4, text="清空日志",
                   command=lambda: self.log.delete("1.0", tk.END)
                   ).pack(anchor="e", pady=(2, 0))

        self.log_msg("[OK] 工具已启动（管理员模式）")
        ds = ", ".join(f"{d[0]}[{d[1]}]" for d in drives) if drives else "无"
        self.log_msg(f"可用盘符：{ds}\n")

    # ========== 启动时检查脱机磁盘 ==========

    def _check_offline_on_start(self):
        threading.Thread(target=self._do_check_offline_start, daemon=True).start()

    def _do_check_offline_start(self):
        offline = get_offline_disks()
        if offline:
            usb_offline = [d for d in offline if d["bus"] in ("USB", "USB3")]
            all_offline = offline
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
            elif all_offline:
                self.log_msg(f"[信息] 检测到 {len(all_offline)} 个脱机磁盘（非USB），可能为正常状态。\n")

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
        drives = get_drives()
        values = [f"{d[0]}  [{d[1]}]" for d in drives]
        self.combo["values"] = values
        cur = self.drive_var.get().split()[0] if self.drive_var.get() else ""
        hit = next((v for v in values if v.startswith(cur)),
                   values[0] if values else "")
        self.drive_var.set(hit)
        ds = ", ".join(f"{d[0]}[{d[1]}]" for d in drives) if drives else "无"
        self.log_msg(f"[刷新] 盘符：{ds}")
        if not drives:
            self.status_lbl.config(text="!! 未检测到可用盘符", foreground="red")
        else:
            self.status_lbl.config(text="", foreground="gray")

        # 同时检查脱机磁盘
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
                self.root.after(0, lambda: self.status_lbl.config(
                    text="", foreground="gray"))

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
            f'powershell -Command "(Get-Partition -DriveLetter {letter} -ErrorAction Stop).DiskNumber"',
            timeout=10
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

        # ── 方法 1：USB 安全移除（最佳：盘符消失 + 硬盘停转）──
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

        # ── 方法 2：DeviceIoControl API ──
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

        # ── 方法 3：Shell.Application Eject ──
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

        # ── 方法 4：Set-Disk -IsOffline + USB 安全移除 ──
        # 注意：IsOffline 会被 Windows 记住，必须追加 USB 移除来避免后遗症
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

        # ── 方法 5：diskpart ──
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