@echo off
chcp 65001 >nul
echo.
echo  ╔══════════════════════════════════════╗
echo  ║     DiskOut 安全弹盘工具 - 打包      ║
echo  ╚══════════════════════════════════════╝
echo.

REM ── 检测 Python ──
set "PYTHON_CMD="

REM 优先尝试 py 启动器（Windows 官方安装通常自带）
py --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py"
    goto :PYTHON_FOUND
)

REM 尝试 python
python --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    goto :PYTHON_FOUND
)

REM 尝试 python3
python3 --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=python3"
    goto :PYTHON_FOUND
)

REM 都找不到，让用户手动指定
echo [错误] 未检测到 Python (已尝试 py / python / python3)
echo.
set /p PYTHON_CMD="请输入你的 Python 命令或完整路径 (例如 C:\Python312\python.exe): "
if "%PYTHON_CMD%"=="" (
    echo [错误] 未输入任何内容，退出
    pause
    exit /b 1
)
%PYTHON_CMD% --version >nul 2>&1
if errorlevel 1 (
    echo [错误] "%PYTHON_CMD%" 无法执行，请确认路径正确
    pause
    exit /b 1
)

:PYTHON_FOUND
echo [OK] 使用 Python 命令: %PYTHON_CMD%
for /f "tokens=*" %%i in ('%PYTHON_CMD% --version 2^>^&1') do echo     %%i

REM ── 检测 pip ──
%PYTHON_CMD% -m pip --version >nul 2>&1
if errorlevel 1 (
    echo [错误] pip 不可用，请先安装 pip
    echo        %PYTHON_CMD% -m ensurepip --upgrade
    pause
    exit /b 1
)
echo [OK] pip 可用

REM ── 检测源码 ──
if not exist "diskout.py" (
    echo.
    echo [错误] 当前目录下未找到 diskout.py
    echo        请将 diskout.py 放到与本脚本相同的目录
    pause
    exit /b 1
)
echo [OK] diskout.py 已找到

REM ── 安装 / 升级 PyInstaller ──
echo.
echo [1/4] 安装 PyInstaller ...
%PYTHON_CMD% -m pip install --upgrade pyinstaller -q
if errorlevel 1 (
    echo [错误] PyInstaller 安装失败，请检查网络或 pip 配置
    pause
    exit /b 1
)
echo [OK] PyInstaller 就绪

REM ── 生成 / 检测图标 ──
echo.
echo [2/4] 准备图标 ...
set "ICON_OPT="

REM 优先级1：用户手动放置的 diskout.ico（最高优先级）
if exist "diskout.ico" (
    echo [OK] 检测到已有 diskout.ico，直接使用
    set "ICON_OPT=--icon=diskout.ico"
    goto :ICON_DONE
)

REM 优先级2：通过 gen_icon.py 自动生成
if not exist "gen_icon.py" (
    echo [跳过] gen_icon.py 不存在且无 diskout.ico，将使用默认图标
    echo        如需自定义图标，请将 diskout.ico 放到当前目录
    goto :ICON_DONE
)

REM gen_icon.py 存在，确保 Pillow 已安装
echo [INFO] 检测 Pillow 依赖 ...
%PYTHON_CMD% -c "import PIL" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Pillow 未安装，正在自动安装 ...
    %PYTHON_CMD% -m pip install Pillow -q
    if errorlevel 1 (
        echo [跳过] Pillow 安装失败，将使用默认图标
        echo        手动安装: %PYTHON_CMD% -m pip install Pillow
        goto :ICON_DONE
    )
    echo [OK] Pillow 安装成功
) else (
    echo [OK] Pillow 已就绪
)

REM 运行图标生成脚本
echo [INFO] 运行 gen_icon.py 生成图标 ...
%PYTHON_CMD% gen_icon.py
if errorlevel 1 (
    echo [跳过] gen_icon.py 运行出错，将使用默认图标
    goto :ICON_DONE
)
if exist "diskout.ico" (
    echo [OK] 图标生成成功
    set "ICON_OPT=--icon=diskout.ico"
) else (
    echo [跳过] gen_icon.py 运行完毕但未生成 diskout.ico，将使用默认图标
)

:ICON_DONE

REM ── 打包 ──
echo.
echo [3/4] 正在打包为单文件 EXE ...
if defined ICON_OPT (
    echo        图标: diskout.ico
) else (
    echo        图标: 系统默认
)
echo        这可能需要 1~3 分钟，请耐心等待 ...
echo.

%PYTHON_CMD% -m PyInstaller --onefile --windowed --uac-admin --name DiskOut %ICON_OPT% --clean diskout.py
if errorlevel 1 (
    echo.
    echo [错误] 打包失败，请查看上方错误信息
    pause
    exit /b 1
)

REM ── 完成 ──
echo.
echo [4/4] 打包完成!
echo.
echo  ╔══════════════════════════════════════╗
echo  ║  输出文件:  dist\DiskOut.exe         ║
echo  ╚══════════════════════════════════════╝
echo.
echo  提示: 首次运行会请求管理员权限，这是正常行为。
echo.

REM ── 清理临时文件（可选）──
set /p CLEANUP="是否清理打包临时文件 (build/, *.spec)? [Y/n]: "
if /i "%CLEANUP%"=="n" goto :DONE
if exist "build" rd /s /q build
if exist "DiskOut.spec" del /q DiskOut.spec
if exist "__pycache__" rd /s /q __pycache__
echo [OK] 临时文件已清理

REM 图标文件单独询问（用户可能想保留）
if exist "diskout.ico" (
    set /p CLEAN_ICON="是否删除 diskout.ico? [y/N]: "
    if /i "!CLEAN_ICON!"=="y" (
        del /q diskout.ico
        echo [OK] diskout.ico 已删除
    ) else (
        echo [保留] diskout.ico（下次打包可直接复用）
    )
)

:DONE
echo.
pause