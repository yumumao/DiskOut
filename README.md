# DiskOut — 安全弹盘工具

**DiskOut** 是一款 Windows 下的移动硬盘清理与安全弹出工具，支持 USB 硬件级安全移除（停转硬盘），
解决 Windows 系统"该设备正在使用中"无法弹出的问题。

**DiskOut** is a Windows utility for cleaning and safely ejecting removable hard drives.
It supports USB hardware-level safe removal (spinning down the drive) and solves the
common "This device is currently in use" ejection failure.

---

## ✨ 功能 / Features

### 中文

- **多方法安全弹出**：依次尝试 USB 硬件弹出、API 弹出、Shell 弹出、磁盘脱机、diskpart 共 5 种方法
- **USB 硬件级移除**：调用 `CM_Request_Device_Eject`，让硬盘真正停转断电，等同于系统托盘"安全删除硬件"
- **一键停止占用服务**：自动停止 Windows Search、SysMain、VSS 等常见占用服务，弹出后自动恢复
- **检测占用进程**：扫描哪些进程/服务正在访问目标盘符
- **删除系统垃圾文件夹**：清除 `System Volume Information` 和 `$RECYCLE.BIN`
- **禁止 SYSTEM 写入**：阻止系统服务在移动硬盘上自动创建文件
- **脱机磁盘恢复**：自动检测并修复因异常弹出导致的磁盘脱机问题
- **管理员自动提权**：启动时自动请求管理员权限

### English

- **Multi-method safe ejection**: Tries 5 methods in sequence — USB hardware eject, API eject, Shell eject, disk offline, and diskpart
- **USB hardware-level removal**: Calls `CM_Request_Device_Eject` to spin down and power off the drive, equivalent to "Safely Remove Hardware" in the system tray
- **One-click service stopping**: Automatically stops Windows Search, SysMain, VSS, and other common locking services, then restores them after ejection
- **Lock detection**: Scans which processes/services are accessing the target drive
- **System folder cleanup**: Removes `System Volume Information` and `$RECYCLE.BIN`
- **Block SYSTEM writes**: Prevents system services from automatically creating files on the removable drive
- **Offline disk recovery**: Automatically detects and fixes disks stuck in offline state due to abnormal ejection
- **Auto admin elevation**: Automatically requests administrator privileges on startup

---

## 📋 系统要求 / Requirements

| 项目 / Item         | 要求 / Requirement             |
|---------------------|-------------------------------|
| 操作系统 / OS       | Windows 10 / 11               |
| 权限 / Privileges   | 管理员 / Administrator         |
| 运行时 / Runtime    | 无（EXE 已内置）/ None (EXE is self-contained) |

如需从源码运行 / To run from source:

| 项目 / Item         | 要求 / Requirement             |
|---------------------|-------------------------------|
| Python              | 3.8+                          |
| 依赖 / Dependencies | 仅标准库 / Standard library only |

---

## 🚀 使用方法 / Usage

### 直接运行 EXE / Run EXE directly

1. 双击 `DiskOut.exe`
2. 在弹出的 UAC 提示中点击"是"
3. 选择要操作的盘符（仅显示 G: 及之后的盘符，避免误操作系统盘）
4. 选择需要的操作

---

1. Double-click `DiskOut.exe`
2. Click "Yes" on the UAC prompt
3. Select the target drive letter (only G: and later are shown to avoid accidental system drive operations)
4. Choose the desired operation

### 从源码运行 / Run from source

```bash
# 以管理员身份运行 / Run as administrator
python diskout.py
```

---

## 🔨 从源码打包 / Build from Source

### 方法一：使用打包脚本 / Method 1: Use the build script

```
将以下文件放在同一目录 / Place these files in the same directory:
├── diskout.py       (主程序 / main program)
├── build.bat        (打包脚本 / build script)
└── gen_icon.py      (图标生成 / icon generator, optional)

双击 build.bat 即可 / Double-click build.bat
```

输出 / Output: `dist\DiskOut.exe`

### 方法二：手动打包 / Method 2: Manual build

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --uac-admin --name DiskOut diskout.py
```

### 可选：生成图标 / Optional: Generate icon

```bash
pip install Pillow
python gen_icon.py
pyinstaller --onefile --windowed --uac-admin --name DiskOut --icon=diskout.ico diskout.py
```

---

## 📖 操作说明 / Operation Guide

### 解除占用 / 弹出 (Unlock / Eject)

| 按钮 / Button | 说明 / Description |
|---|---|
| **检测** | 扫描占用该盘的进程和服务 / Scan processes and services locking the drive |
| **停止服务** | 停止 6 个常见占用服务 / Stop 6 common locking services |
| **恢复服务** | 恢复已停止的服务 / Restore stopped services |
| **安全弹出 ⭐** | 停止服务 → 弹出 → 恢复服务（推荐）/ Stop services → Eject → Restore (recommended) |
| **强制弹出** | 跳过停止服务，直接尝试弹出 / Skip service stopping, eject directly |
| **恢复脱机磁盘** | 修复因脱机标记导致无盘符的磁盘 / Fix disks with no drive letter due to offline flag |

### 删除系统文件夹 (Delete System Folders)

删除 Windows 自动创建的隐藏文件夹，释放空间并保持移动硬盘整洁。

Delete hidden folders automatically created by Windows to free space and keep the drive clean.

### SYSTEM 写入权限 (SYSTEM Write Permission)

禁止后，Windows 系统服务将无法在该盘创建 `System Volume Information` 等文件夹。
拔盘前建议恢复权限，以免影响其他用途。

When denied, Windows system services cannot create folders like `System Volume Information` on the drive.
It is recommended to restore permissions before unplugging to avoid affecting other uses.

---

## ⚡ 弹出方法说明 / Ejection Methods

DiskOut 会按优先级依次尝试以下方法，成功即停止：

DiskOut tries the following methods in priority order, stopping on first success:

| 优先级 / Priority | 方法 / Method | 说明 / Description |
|---|---|---|
| 1 | CM_Request_Device_Eject | USB 硬件级安全移除，硬盘停转 / USB hardware-level safe removal, drive spins down |
| 2 | DeviceIoControl API | Windows API 弹出卷 / Windows API volume ejection |
| 3 | Shell.Application Eject | 模拟资源管理器弹出操作 / Simulates Explorer eject action |
| 4 | Set-Disk -IsOffline | 将磁盘标记为脱机（有后遗症，最后手段）/ Marks disk offline (has side effects, last resort) |
| 5 | diskpart | 使用 diskpart 卸载卷 / Uses diskpart to dismount volume |

---

## ⚠️ 注意事项 / Notes

- 本工具仅显示 **G:** 及之后的盘符，以防止误操作系统盘或常用数据盘。
  This tool only shows drive letters **G:** and later to prevent accidental operations on system or common data drives.

- **安全弹出**（推荐按钮）会自动停止服务、弹出、再恢复服务，是最安全的一键操作。
  **Safe Eject** (recommended button) automatically stops services, ejects, then restores services — the safest one-click operation.

- 如果硬盘弹出后重新插入没有盘符，点击**恢复脱机磁盘**按钮即可修复。
  If the drive has no letter after re-insertion, click the **Recover Offline Disk** button to fix it.

- 禁止 SYSTEM 写入后，某些依赖该权限的功能（如系统还原点）将无法在该盘使用。
  After denying SYSTEM write access, certain features (like System Restore points) will not work on the drive.

---

## 📁 文件结构 / File Structure

```
DiskOut/
├── diskout.py       # 主程序源码 / Main program source
├── build.bat        # 一键打包脚本 / One-click build script
├── gen_icon.py      # 图标生成脚本 / Icon generator script
├── README.md        # 本文件 / This file
└── dist/
    └── DiskOut.exe  # 打包后的可执行文件 / Built executable
```

---

## 📄 License

MIT License

Copyright (c) 2025

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.