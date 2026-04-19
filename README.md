# DiskOut — 安全弹盘工具

**DiskOut** 是一款 Windows 下的移动硬盘清理与安全弹出工具，支持 USB 硬件级安全移除（停转硬盘），
解决 Windows 系统"该设备正在使用中"无法弹出的问题。

**DiskOut** is a Windows utility for cleaning and safely ejecting removable hard drives.
It supports USB hardware-level safe removal (spinning down the drive) and solves the
common "This device is currently in use" ejection failure.

> 💡 **相比常见的弹出工具**，DiskOut 内置 5 种弹出方法逐一尝试（USB 硬件级移除 → API → Shell → 磁盘脱机 → diskpart），
> 覆盖了从用户态到系统底层的多条路径，成功率显著高于仅使用单一方法的工具。
> 所有操作均通过 Windows 标准 API 和系统命令实现，**不安装驱动、不注入进程、不修改注册表**，
> 弹出后自动恢复被停止的服务，不会对系统产生持久性影响。

> 💡 **Compared to common ejection tools**, DiskOut has 5 built-in ejection methods tried in sequence
> (USB hardware removal → API → Shell → disk offline → diskpart), covering multiple paths from user mode
> to the system level, resulting in a significantly higher success rate than single-method tools.
> All operations use standard Windows APIs and system commands — **no drivers installed, no process injection,
> no registry modifications**. Stopped services are automatically restored after ejection, leaving no
> persistent impact on the system.

---

## ✨ 功能 / Features

### 中文

- **多方法安全弹出**：依次尝试 USB 硬件弹出、API 弹出、Shell 弹出、磁盘脱机、diskpart 共 5 种方法
- **USB 硬件级移除**：调用 `CM_Request_Device_Eject`，让硬盘真正停转断电，等同于系统托盘"安全删除硬件"
- **一键停止占用服务**：自动停止 Windows Search、SysMain、VSS 等常见占用服务，弹出后自动恢复
- **检测占用进程**：扫描哪些进程/服务正在访问目标盘符
- **文件/文件夹占用检测**：使用 Restart Manager API 精确检测任意文件或文件夹的占用进程，并支持一键结束
- **删除系统垃圾文件夹**：清除 `System Volume Information` 和 `$RECYCLE.BIN`
- **禁止 SYSTEM 写入**：阻止系统服务在移动硬盘上自动创建文件
- **脱机磁盘恢复**：自动检测并修复因异常弹出导致的磁盘脱机问题
- **普通模式启动，按需提权**：默认以普通用户权限运行，仅在需要管理员权限的操作时提示提升，也可随时手动提升为管理员

### English

- **Multi-method safe ejection**: Tries 5 methods in sequence — USB hardware eject, API eject, Shell eject, disk offline, and diskpart
- **USB hardware-level removal**: Calls `CM_Request_Device_Eject` to spin down and power off the drive, equivalent to "Safely Remove Hardware" in the system tray
- **One-click service stopping**: Automatically stops Windows Search, SysMain, VSS, and other common locking services, then restores them after ejection
- **Lock detection**: Scans which processes/services are accessing the target drive
- **File/folder lock detection**: Uses the Restart Manager API to precisely detect processes locking any file or folder, with one-click termination
- **System folder cleanup**: Removes `System Volume Information` and `$RECYCLE.BIN`
- **Block SYSTEM writes**: Prevents system services from automatically creating files on the removable drive
- **Offline disk recovery**: Automatically detects and fixes disks stuck in offline state due to abnormal ejection
- **Normal mode by default, elevate on demand**: Runs with standard user privileges by default; prompts for admin elevation only when needed, with an option to manually elevate at any time

---

## 📋 系统要求 / Requirements

| 项目 / Item         | 要求 / Requirement             |
|---------------------|-------------------------------|
| 操作系统 / OS       | Windows 10 / 11               |
| 权限 / Privileges   | 普通用户即可启动，部分功能需管理员 / Standard user to launch, some features require Administrator |
| 运行时 / Runtime    | 无（EXE 已内置）/ None (EXE is self-contained) |

如需从源码运行 / To run from source:

| 项目 / Item         | 要求 / Requirement             |
|---------------------|-------------------------------|
| Python              | 3.8+                          |
| 依赖 / Dependencies | 仅标准库 / Standard library only |
| 可选 / Optional     | `windnd`（拖放支持）/ `windnd` (drag & drop support) |

---

## 🚀 使用方法 / Usage

### 直接运行 EXE / Run EXE directly

1. 双击 `DiskOut.exe`（无需管理员权限即可启动）
2. 选择要操作的盘符（仅显示 G: 及之后的盘符，避免误操作系统盘）
3. 选择需要的操作
4. 如果所选操作需要管理员权限，程序会自动提示是否提升权限
5. 也可随时点击右上角「⬆ 提升为管理员」按钮手动获取完整功能

---

1. Double-click `DiskOut.exe` (no admin privileges required to launch)
2. Select the target drive letter (only G: and later are shown to avoid accidental system drive operations)
3. Choose the desired operation
4. If the selected operation requires admin privileges, the program will prompt for elevation
5. You can also click the "⬆ Elevate to Admin" button in the top-right corner at any time for full functionality

### 从源码运行 / Run from source

```bash
# 直接运行（普通用户即可）/ Run directly (standard user is fine)
python diskout.py

# 安装拖放支持（可选）/ Install drag & drop support (optional)
pip install windnd
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
pyinstaller --onefile --windowed --name DiskOut diskout.py
```

> **注意 / Note**: v3.0 起不再需要 `--uac-admin` 参数，程序会在需要时自行请求提权。
> Since v3.0, the `--uac-admin` flag is no longer needed — the program requests elevation on demand.

### 可选：生成图标 / Optional: Generate icon

```bash
pip install Pillow
python gen_icon.py
pyinstaller --onefile --windowed --name DiskOut --icon=diskout.ico diskout.py
```

---

## 📖 操作说明 / Operation Guide

### Tab 1：解除占用 / 弹出 (Unlock / Eject)

| 按钮 / Button | 权限 / Privilege | 说明 / Description |
|---|---|---|
| **检测占用进程和服务** | 普通 / Standard | 扫描占用该盘的进程和服务 / Scan processes and services locking the drive |
| **一键停止占用服务** | 管理员 / Admin | 停止 6 个常见占用服务 / Stop 6 common locking services |
| **恢复已停止的服务** | 管理员 / Admin | 恢复已停止的服务 / Restore stopped services |
| **恢复脱机磁盘** | 管理员 / Admin | 修复因脱机标记导致无盘符的磁盘 / Fix disks with no drive letter due to offline flag |
| **★ 安全弹出（推荐）** | 管理员 / Admin | 停止服务 → 弹出 → 恢复服务（推荐）/ Stop services → Eject → Restore (recommended) |
| **强制弹出** | 普通可用，管理员更完整 / Standard works, Admin for full methods | 跳过停止服务，直接尝试弹出 / Skip service stopping, eject directly |

### Tab 2：文件/文件夹占用 (File/Folder Lock Detection)

| 按钮 / Button | 权限 / Privilege | 说明 / Description |
|---|---|---|
| **检测占用** | 普通 / Standard | 使用 Restart Manager API 检测指定文件/文件夹的占用进程 / Detect locking processes using Restart Manager API |
| **一键停止所有占用** | 涉及服务时需管理员 / Admin when services involved | 结束占用进程并停止相关服务 / Kill locking processes and stop related services |
| **恢复已停止的服务** | 管理员 / Admin | 恢复已停止的服务 / Restore stopped services |

> 支持拖放文件/文件夹到窗口（需安装 `windnd`：`pip install windnd`）
> Supports drag & drop files/folders into the window (requires `windnd`: `pip install windnd`)

### Tab 3：进阶功能 (Advanced Features)

| 按钮 / Button | 权限 / Privilege | 说明 / Description |
|---|---|---|
| **删除 System Volume Information** | 管理员 / Admin | 删除系统卷信息文件夹 / Delete System Volume Information folder |
| **删除 $RECYCLE.BIN** | 管理员 / Admin | 删除回收站文件夹 / Delete Recycle Bin folder |
| **一键删除以上两个文件夹** | 管理员 / Admin | 同时删除两个系统文件夹 / Delete both system folders |
| **禁止 SYSTEM 写入** | 管理员 / Admin | 阻止系统服务在该盘创建文件 / Prevent system services from creating files |
| **恢复 SYSTEM 写入** | 管理员 / Admin | 恢复系统写入权限 / Restore system write permissions |

---

## 🔑 权限模式说明 / Permission Modes

DiskOut v3.0 起采用**按需提权**模式：

Since v3.0, DiskOut uses an **elevate-on-demand** model:

| 模式 / Mode | 标题栏 / Title Bar | 说明 / Description |
|---|---|---|
| 普通模式 / Standard Mode | `移动硬盘清理工具 - 普通模式` | 检测、强制弹出等基础功能可用 / Detection, force eject and other basic features available |
| 管理员模式 / Admin Mode | `移动硬盘清理工具 - 管理员模式` | 所有功能完整可用 / All features fully available |

**普通模式下的限制 / Limitations in Standard Mode:**
- 无法停止/恢复系统服务
- 弹出方法 4（Set-Disk -IsOffline）和方法 5（diskpart）不可用
- 无法删除系统文件夹或修改磁盘权限
- 无法恢复脱机磁盘
- `openfiles` 查询不可用

**提升方式 / How to elevate:**
- 点击需要管理员权限的按钮时，程序会自动弹窗提示提升
- 也可随时点击右上角「⬆ 提升为管理员」按钮主动提升
- 提升后程序会以管理员身份重新启动（当前窗口关闭）

---

## ⚡ 弹出方法说明 / Ejection Methods

DiskOut 会按优先级依次尝试以下方法，成功即停止：

DiskOut tries the following methods in priority order, stopping on first success:

| 优先级 / Priority | 方法 / Method | 权限 / Privilege | 说明 / Description |
|---|---|---|---|
| 1 | CM_Request_Device_Eject | 普通 / Standard | USB 硬件级安全移除，硬盘停转 / USB hardware-level safe removal, drive spins down |
| 2 | DeviceIoControl API | 普通 / Standard | Windows API 弹出卷 / Windows API volume ejection |
| 3 | Shell.Application Eject | 普通 / Standard | 模拟资源管理器弹出操作 / Simulates Explorer eject action |
| 4 | Set-Disk -IsOffline | 管理员 / Admin | 将磁盘标记为脱机（有后遗症，最后手段）/ Marks disk offline (has side effects, last resort) |
| 5 | diskpart | 管理员 / Admin | 使用 diskpart 卸载卷 / Uses diskpart to dismount volume |

> 普通模式下方法 4 和 5 会被自动跳过，不会报错。
> In standard mode, methods 4 and 5 are automatically skipped without errors.

---

## 🛡️ 安全性说明 / Security

- **无驱动安装**：不安装任何内核驱动，所有操作均在用户态完成
- **无进程注入**：不向其他进程注入代码或 DLL
- **无注册表修改**：不写入或修改 Windows 注册表
- **无网络访问**：工具完全离线运行，不发送任何数据
- **服务自动恢复**：被停止的系统服务在弹出后自动恢复，不会遗留系统变更
- **源码公开**：单文件 Python 源码，可随时审查

---

- **No driver installation**: No kernel drivers installed; all operations run in user mode
- **No process injection**: No code or DLL injection into other processes
- **No registry modification**: Does not write to or modify the Windows registry
- **No network access**: The tool runs completely offline and sends no data
- **Automatic service restoration**: Stopped system services are automatically restored after ejection, leaving no lingering system changes
- **Open source**: Single-file Python source code, fully auditable at any time

---

## ⚠️ 注意事项 / Notes

- 本工具仅显示 **G:** 及之后的盘符，以防止误操作系统盘或常用数据盘（可手动启用 D:/E:/F:）。
  This tool only shows drive letters **G:** and later to prevent accidental operations on system or common data drives (D:/E:/F: can be manually enabled).

- **安全弹出**（推荐按钮）会自动停止服务、弹出、再恢复服务，是最安全的一键操作（需管理员权限）。
  **Safe Eject** (recommended button) automatically stops services, ejects, then restores services — the safest one-click operation (requires admin).

- **强制弹出**在普通模式下也能使用前 3 种弹出方法，适合不想提权的场景。
  **Force Eject** can use the first 3 ejection methods in standard mode, suitable when you prefer not to elevate.

- 如果硬盘弹出后重新插入没有盘符，点击**恢复脱机磁盘**按钮即可修复。
  If the drive has no letter after re-insertion, click the **Recover Offline Disk** button to fix it.

- 禁止 SYSTEM 写入后，某些依赖该权限的功能（如系统还原点）将无法在该盘使用。
  After denying SYSTEM write access, certain features (like System Restore points) will not work on the drive.

- v3.0 起程序不再强制要求管理员权限启动，打包时也无需添加 `--uac-admin` 参数。
  Since v3.0, the program no longer requires admin privileges to launch, and `--uac-admin` is no longer needed when building.

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