@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
echo.
echo  ╔══════════════════════════════════════╗
echo  ║     DiskOut 安全弹盘工具 - 打包      ║
echo  ╚══════════════════════════════════════╝
echo.

REM ── 检测 Python ──
set "PYTHON_CMD="

py --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py"
    goto :PYTHON_FOUND
)

python --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    goto :PYTHON_FOUND
)

python3 --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=python3"
    goto :PYTHON_FOUND
)

echo [错误] 未检测到 Python (已尝试 py / python / python3)
echo.
set /p PYTHON_CMD="请输入你的 Python 命令或完整路径: "
if "%PYTHON_CMD%"=="" (
    echo [错误] 未输入任何内容，退出
    pause
    exit /b 1
)
%PYTHON_CMD% --version >nul 2>&1
if errorlevel 1 (
    echo [错误] "%PYTHON_CMD%" 无法执行
    pause
    exit /b 1
)

:PYTHON_FOUND
echo [OK] 使用 Python 命令: %PYTHON_CMD%
for /f "tokens=*" %%i in ('%PYTHON_CMD% --version 2^>^&1') do echo     %%i

REM ── 检测源码 ──
if not exist "diskout.py" (
    echo [错误] 当前目录下未找到 diskout.py
    pause
    exit /b 1
)
echo [OK] diskout.py 已找到

REM ── 创建干净的虚拟环境 ──
echo.
echo [1/5] 创建干净的打包虚拟环境 ...
if exist ".build_venv" (
    echo [INFO] 删除旧的虚拟环境 ...
    rd /s /q .build_venv
)
%PYTHON_CMD% -m venv .build_venv
if errorlevel 1 (
    echo [错误] 虚拟环境创建失败
    pause
    exit /b 1
)

set "VENV_PYTHON=.build_venv\Scripts\python.exe"
echo [OK] 虚拟环境已创建

REM ── 在 venv 中安装最小依赖 ──
echo.
echo [2/5] 安装最小依赖 (PyInstaller + windnd) ...
%VENV_PYTHON% -m pip install --upgrade pip -q
%VENV_PYTHON% -m pip install pyinstaller -q
%VENV_PYTHON% -m pip install windnd -q
if errorlevel 1 (
    echo [警告] windnd 安装失败，打包后拖放功能将不可用
) else (
    echo [OK] windnd 安装成功
)
echo [OK] 依赖安装完成

REM ── 准备图标 ──
echo.
echo [3/5] 准备图标 ...
set "ICON_OPT="
set "ADD_DATA_OPT="

REM 优先使用现成的 ico
if exist "diskout.ico" (
    echo [OK] 检测到 diskout.ico
    set "ICON_OPT=--icon=diskout.ico"
    set "ADD_DATA_OPT=--add-data "diskout.ico;.""
    goto :ICON_DONE
)

REM 没有 ico 但有生成脚本，则尝试生成
if not exist "gen_icon.py" (
    echo [跳过] 无图标文件也无 gen_icon.py，使用默认图标
    goto :ICON_DONE
)

echo [INFO] 检测 Pillow ...
%VENV_PYTHON% -c "import PIL" >nul 2>&1
if errorlevel 1 (
    echo [INFO] 安装 Pillow ...
    %VENV_PYTHON% -m pip install Pillow -q
)

echo [INFO] 运行 gen_icon.py ...
%VENV_PYTHON% gen_icon.py

if exist "diskout.ico" (
    echo [OK] 图标生成成功
    set "ICON_OPT=--icon=DiskOut.ico"
    set "ADD_DATA_OPT=--add-data "DiskOut.ico;.""
) else (
    echo [跳过] 图标生成失败，使用默认图标
)

:ICON_DONE

REM ── 验证图标格式 ──
if defined ICON_OPT (
    echo [INFO] 验证图标文件 ...
    %VENV_PYTHON% -c "f=open('diskout.ico','rb');h=f.read(4);f.close();assert h[:4]==b'\x00\x00\x01\x00','Not a valid ICO'" >nul 2>&1
    if errorlevel 1 (
        echo [警告] diskout.ico 不是有效的 ICO 格式 (可能是 PNG 改后缀)
        echo [INFO] 尝试自动转换 ...
        %VENV_PYTHON% -c "import PIL" >nul 2>&1
        if errorlevel 1 (
            %VENV_PYTHON% -m pip install Pillow -q
        )
        %VENV_PYTHON% -c "from PIL import Image;img=Image.open('diskout.ico');img.save('diskout.ico',format='ICO',sizes=[(16,16),(32,32),(48,48),(256,256)])" >nul 2>&1
        if errorlevel 1 (
            echo [警告] 转换失败，放弃图标，使用默认
            set "ICON_OPT="
            set "ADD_DATA_OPT="
        ) else (
            echo [OK] 图标已转换为有效 ICO 格式
        )
    ) else (
        echo [OK] 图标格式验证通过
    )
)

REM ── 打包 ──
echo.
echo [4/5] 正在打包 (干净环境 + 排除冗余模块) ...
echo       这可能需要 1~3 分钟 ...
echo.

%VENV_PYTHON% -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name DiskOut ^
    %ICON_OPT% ^
    %ADD_DATA_OPT% ^
    --hidden-import windnd ^
    --exclude-module numpy ^
    --exclude-module pandas ^
    --exclude-module matplotlib ^
    --exclude-module scipy ^
    --exclude-module PIL ^
    --exclude-module tkinter.test ^
    --exclude-module unittest ^
    --exclude-module pydoc ^
    --exclude-module doctest ^
    --exclude-module lib2to3 ^
    --exclude-module xmlrpc ^
    --exclude-module multiprocessing ^
    --strip ^
    --clean ^
    diskout.py

if errorlevel 1 (
    echo.
    echo [错误] 打包失败，请查看上方错误信息
    pause
    exit /b 1
)

REM ── 显示结果 ──
echo.
echo [5/5] 打包完成!
echo.
for %%A in (dist\DiskOut.exe) do (
    set "SIZE=%%~zA"
    set /a "SIZE_MB=!SIZE! / 1048576"
    echo  ╔══════════════════════════════════════╗
    echo  ║  输出: dist\DiskOut.exe              ║
    echo  ║  大小: !SIZE_MB! MB                          ║
    echo  ╚══════════════════════════════════════╝
)
echo.

REM ── 提示 ──
echo [提示] 如果资源管理器中图标未更新，请执行:
echo        ie4uinit.exe -show
echo        或将 EXE 复制到新文件夹查看
echo.

REM ── 清理 ──
set /p CLEANUP="是否清理打包临时文件 (build/, *.spec, .build_venv/)? [Y/n]: "
if /i "%CLEANUP%"=="n" goto :DONE
if exist "build" rd /s /q build
if exist "DiskOut.spec" del /q DiskOut.spec
if exist "__pycache__" rd /s /q __pycache__
if exist ".build_venv" rd /s /q .build_venv
echo [OK] 临时文件已清理

if exist "diskout.ico" (
    set /p CLEAN_ICON="是否删除 diskout.ico? [y/N]: "
    if /i "!CLEAN_ICON!"=="y" (
        del /q diskout.ico
        echo [OK] diskout.ico 已删除
    ) else (
        echo [保留] diskout.ico
    )
)

:DONE
echo.
pause