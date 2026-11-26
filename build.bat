@echo off
chcp 65001 >nul
REM ============================================================
REM  打包脚本（报告自动上传助手）
REM  适配环境：Python 3.14.6 + PyInstaller 6.22.3
REM  已包含的 workaround：
REM   1) PyInstaller 隔离子进程在 Python 3.14 下崩溃 -> 已在
REM      E:\python\Lib\site-packages\PyInstaller\isolated\_parent.py
REM      中禁用隔离（备份见 _parent.py.bak）
REM   2) OpenBLAS 多线程内存分配失败 -> 限制单线程
REM   3) spec 内已排除 PyQt5 等无关依赖：PyQt5 自带的旧版 VC 运行库
REM      (msvcp140.dll 14.26) 会覆盖系统新版，导致 onnxruntime/opencv 初始化
REM      时在 msvcp140.dll 内触发 0xC0000005，exe 双击即崩溃。
REM  如重装/升级 PyInstaller，需重新应用隔离禁用补丁。
REM ============================================================
set OPENBLAS_NUM_THREADS=1
set OMP_NUM_THREADS=1
cd /d "%~dp0"
E:\python\Scripts\pyinstaller.exe --clean "自动上传检测报告.spec"
pause
