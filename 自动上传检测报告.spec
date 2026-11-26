# -*- mode: python ; coding: utf-8 -*-
"""报告自动上传助手 —— PyInstaller 打包配置（单文件 / 无控制台）。

【为什么不能直接用 collect_all】
旧写法用 collect_all('ddddocr'/'onnxruntime'/'openpyxl')，会把这些包的
全部子模块与依赖都塞进来（IPython / pandas / pyarrow / transformers /
PyQt5 ...），exe 达 271MB。更要命的是 PyQt5\\Qt5\\bin 自带的 VC 运行库
(msvcp140.dll / vcruntime140.dll，版本 14.26) 会被提升到包根目录，顶掉系统
里更新的 14.51 版本；onnxruntime / opencv 这些用新版 VC 编译的原生模块在
这种混用环境下会在 msvcp140.dll 内部触发 0xC0000005 访问冲突，exe 双击即
崩溃，且 windowed 模式下没有任何提示。

【现在的做法】只按需收集：
  - ddddocr 的 *.onnx 模型文件（代码按包目录相对路径加载，必须随包）；
  - onnxruntime 的原生 DLL 与附带数据；
其余交给 PyInstaller 静态分析，并显式排除用不到的大模块。
"""

from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

datas = []
binaries = []
hiddenimports = []

# ddddocr：验证码模型 common_old.onnx / common.onnx / common_det.onnx
try:
    datas += collect_data_files('ddddocr')
except Exception:
    pass

# onnxruntime：推理引擎原生库
try:
    binaries += collect_dynamic_libs('onnxruntime')
    datas += collect_data_files('onnxruntime')
except Exception:
    pass

# ddddocr 内部部分模块为动态导入，显式声明避免漏打
hiddenimports += [
    'ddddocr.compat.v1',
    'ddddocr.core.ocr_engine',
    'ddddocr.core.detection_engine',
    'ddddocr.core.slide_engine',
    'ddddocr.models.model_loader',
    'ddddocr.models.charset_manager',
    'ddddocr.preprocessing.image_processor',
    'ddddocr.preprocessing.color_filter',
    'onnxruntime.capi.onnxruntime_inference_collection',
    'onnxruntime.capi.onnxruntime_pybind11_state',
    'openpyxl',
]

# 体积大且用不到的依赖（尤其 PyQt5：它会带来旧版 VC 运行库）
excludes = [
    'PyQt5', 'PyQt6', 'PySide2', 'PySide6',
    'IPython', 'ipykernel', 'jupyter_client', 'jupyter_core', 'notebook', 'nbformat',
    'pandas', 'pyarrow', 'datasets', 'transformers', 'huggingface_hub', 'hf_xet',
    'botocore', 'matplotlib', 'zmq', 'debugpy', 'tiktoken', 'tqdm',
    'fastapi', 'uvicorn', 'mcp', 'pydantic', 'rich', 'click', 'markdown_it',
]

a = Analysis(
    ['auto_upload_cdp_gui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='报告自动上传助手',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
