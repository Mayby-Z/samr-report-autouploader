# -*- coding: utf-8 -*-
"""
自动化上传报告 demo：通过 Chromium 远程调试协议 (CDP) 自动控制
「检验检测报告编号上传系统」Electron 应用。

增强版 GUI：
    - 点击「启动」才开始自动化；
    - **启动时先检测应用是否已经打开**：已打开则复用该实例（不再重复启动应用）；
      若检测到已处于登录后（首页）状态，则连登录流程一起跳过，直接执行登录后的逻辑；
    - 支持「暂停」和「继续」；
    - 支持「重新执行」；
    - **图形验证码自动识别并填入**（ddddocr，识别失败自动重试/降级人工）；
    - 「重载代码」功能已注释停用（保留代码备查）：
      # 原本支持热重载本文件的最新代码，并且只执行登录之后的后续逻辑
      #（不会重新打开应用、也不会重新走登录流程）；
    - 支持打包成 exe。

注意：
    登录页有图形验证码，默认由脚本自动识别并填写；
    若未安装 ddddocr 或连续识别失败，会降级为人工在窗口中输入。
"""

import base64
import ctypes
import datetime
import io
import json
import os
import queue
import re
import shutil
import socket
import sqlite3
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import zipfile
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from ctypes import wintypes
from tkinter import filedialog, messagebox, scrolledtext
from typing import Any, TypedDict
from urllib.parse import quote
from xml.sax.saxutils import escape as _xml_escape

import requests
import websocket

# Windows 原生对话框（保存对话框）自动化需要 user32
_WIN32 = sys.platform.startswith("win")

# 复用无界面版 auto_login_cdp.py 里的 CDP 客户端与“点击登录”脚本
import auto_upload_cdp as _cdp_mod
CDPClient = _cdp_mod.CDPClient
JS_CLICK_LOGIN = _cdp_mod.JS_CLICK_LOGIN

# ===================== 验证码 OCR =====================
# 需要: pip install ddddocr   （未安装时自动降级为人工输入验证码）
_OCR_AVAILABLE = False
try:
    import ddddocr
    from PIL import Image
    _OCR_AVAILABLE = True
except Exception as _ocr_err:  # pragma: no cover
    ddddocr = None
    Image = None
    print("[!] 未安装 ddddocr，验证码将需要人工输入: " + repr(_ocr_err))

# ===================== 配置区 =====================
# 以下为默认值；首次运行若 EXE_PATH 为空，进入页面会强制要求先「设置」软件路径。
# 运行时的可修改配置统一放在 CONFIG 字典里（避免对全大写常量重复赋值）。

class _AppConfig(TypedDict):
    EXE_PATH: str
    USERNAME: str
    PASSWORD: str
    DEBUG_PORT: int
    KEEP_APP_OPEN: bool
    AUTO_DOWNLOAD_TEMPLATE: bool
    AUTO_OPEN_TEMPLATE: bool
    AUTO_FILL_TEMPLATE: bool
    CLOSE_XLSX_AFTER_OPEN: bool
    AUTO_IMPORT_TEMPLATE: bool

DEFAULT_CONFIG: _AppConfig = {
    "EXE_PATH": "",
    "USERNAME": "18922735867",
    "PASSWORD": "Zz791125##",
    "DEBUG_PORT": 9222,
    "KEEP_APP_OPEN": True,   # True: 登录成功后保留应用，便于后续上传操作；False: 流程结束关闭应用
    "AUTO_DOWNLOAD_TEMPLATE": True,  # True: 登录进入首页后自动执行「下载表格模板」流程
    "AUTO_OPEN_TEMPLATE": True,      # True: 模板下载完成后用系统默认程序打开该 .xlsx
    "AUTO_FILL_TEMPLATE": True,      # True: 打开该 .xlsx 之前自动按规则填写数据
    "CLOSE_XLSX_AFTER_OPEN": True,   # True: 打开 xlsx 后等待几秒自动关闭该文件（释放文件锁）
    "AUTO_IMPORT_TEMPLATE": True,    # True: 关闭 xlsx 后自动点「选择数据表格导入」并选中该文件导入
}

STARTUP_TIMEOUT = 30        # 等待应用启动并打开调试端口的秒数
AUTO_CLICK_LOGIN = True     # True: 填完后自动点登录（验证码需已填）；False: 只填表单不点登录
WAIT_CAPTCHA = True         # True: 点击登录前等待你手动输入验证码
CAPTCHA_TIMEOUT = 120       # 等待验证码输入的最长秒数

# ---- 登录时的「邮箱安全验证」人工等待 ----
# 登录提交后如果服务端要求邮箱验证，脚本会暂停自动点击/输入，
# 等你手工在同一个窗口完成验证后再自动继续（不需要再点「继续」）。
EMAIL_VERIFY_DETECT_TIMEOUT = 8    # 提交登录后探测验证页出现的最长秒数（未出现则走原正常路径）
EMAIL_VERIFY_TIMEOUT = 600         # 等待人工完成邮箱验证的最长秒数（默认 10 分钟，可按需修改）
EMAIL_VERIFY_POLL_INTERVAL = 0.5   # 等待期间轮询页面状态的间隔秒数
EMAIL_VERIFY_PROGRESS_LOG = 15     # 每隔多少秒在日志里回报一次「仍在等待」

# ---- 首页「下载表格模板」流程 ----
SAVE_DIALOG_TIMEOUT = 60    # 等待系统「保存」对话框弹出的最长秒数
SAVE_CLICK_TIMEOUT = 20     # 点击「保存」后等待对话框关闭的最长秒数
TEMPLATE_SAVE_TIMEOUT = 30  # 保存后等待文件真正写入磁盘的最长秒数

# ---- 关闭 xlsx + 导入数据表格 ----
CLOSE_XLSX_DELAY = 5          # 打开 xlsx 后等待多少秒再自动关闭该文件
CLOSE_XLSX_WAIT = 30          # 等待 xlsx 窗口关闭 / 文件解锁的最长秒数
OPEN_DIALOG_TIMEOUT = 60      # 等待系统「打开」对话框弹出的最长秒数
OPEN_DIALOG_CLICK_TIMEOUT = 20  # 点击「打开」后等待对话框关闭的最长秒数
IMPORT_RESULT_TIMEOUT = 15    # 「软确认」等待上限：仅做日志提示，超时不阻断流程
IMPORT_RESULT_PROGRESS_LOG = 5  # 等待期间每隔多少秒回报一次进度，避免看起来像卡死
IMPORT_BTN_SETTLE_DELAY = 1   # 关闭 xlsx 后至少等待多少秒再点「选择数据表格导入」
IMPORT_BTN_MAX_ROUNDS = 3     # 点击「选择数据表格导入」最多尝试几轮（以对话框弹出为判据）
IMPORT_DIALOG_APPEAR_TIMEOUT = 20  # 点击后等待文件选择框弹出的最长秒数
IMPORT_REIMPORT_CONFIRM_DELAY = 2  # 「确认重新导入」弹窗停留多少秒后再点「确认导入」

# ---- 「首次启动配置」弹窗自动配置 ----
# 应用首次启动会弹出「首次启动配置」，要求指定「报告备份路径」的根目录
# （应用会在其下自动创建 excel、report 子目录）。这里把这一步做成全自动：
#   先创建根目录 → 点「选择文件夹」选中它 → 点「完成配置」。
# 目录可在 gui_config.json 里用同名键 FIRST_LAUNCH_BACKUP_DIR 覆盖。
DEFAULT_FIRST_LAUNCH_BACKUP_DIR = os.path.join(
    os.environ.get("ProgramData") or r"C:\ProgramData", "AreportFile")
FIRST_LAUNCH_PICK_TIMEOUT = 30     # 等待弹窗里「选择文件夹」按钮出现的最长秒数
FIRST_LAUNCH_DIALOG_TIMEOUT = 30   # 等待原生目录选择对话框弹出的最长秒数
FIRST_LAUNCH_DONE_TIMEOUT = 30     # 点击「完成配置」后等待弹窗关闭的最长秒数

# ---- 图形验证码自动识别 ----
AUTO_CAPTCHA = True         # True: 自动识别并填入图形验证码
CAPTCHA_MAX_ROUNDS = 8      # 自动识别最多尝试几轮（每轮换一张新验证码）
CAPTCHA_MAX_CANDIDATES = 12 # 每张验证码最多校验几个候选答案
CAPTCHA_PRECHECK = True     # True: 先用服务端接口校验候选，确认无误后再填入并点登录
# 补白宽度（CTC 解码空间不足会吞掉末尾字符，补白是本系统最有效的预处理）
CAPTCHA_PADS = (24, 40, 48, 56, 64, 80)
# ====================================================

# ===================== 配置加载 / 保存 =====================
def _config_dir():
    """配置文件目录：打包后放在 exe 旁边，开发时放在脚本目录。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(_config_dir(), "gui_config.json")

# 密码混淆前缀：带此前缀说明已 base64 编码，方便配置文件里不直接看明文
_PWD_B64_PREFIX = "b64:"


def _encode_pwd(plain: str) -> str:
    """把明文密码编成 b64:xxx 形式（可逆，但 config 文件里看不见明文）。"""
    if not plain or plain.startswith(_PWD_B64_PREFIX):
        return plain
    return _PWD_B64_PREFIX + base64.b64encode(plain.encode("utf-8")).decode("ascii")


def _decode_pwd(maybe_enc: str) -> str:
    """把配置里读出来的密码解成明文；没前缀就当旧明文直接返回。"""
    if not maybe_enc:
        return maybe_enc
    if maybe_enc.startswith(_PWD_B64_PREFIX):
        try:
            return base64.b64decode(maybe_enc[len(_PWD_B64_PREFIX):]).decode("utf-8")
        except Exception:
            # 解码失败（被用户手工改过？），就原样返回，让下游自己报错
            return maybe_enc
    return maybe_enc


def load_config() -> _AppConfig:
    cfg: _AppConfig = {
        "EXE_PATH": DEFAULT_CONFIG["EXE_PATH"],
        "USERNAME": DEFAULT_CONFIG["USERNAME"],
        "PASSWORD": DEFAULT_CONFIG["PASSWORD"],
        "DEBUG_PORT": DEFAULT_CONFIG["DEBUG_PORT"],
        "KEEP_APP_OPEN": DEFAULT_CONFIG["KEEP_APP_OPEN"],
        "AUTO_DOWNLOAD_TEMPLATE": DEFAULT_CONFIG["AUTO_DOWNLOAD_TEMPLATE"],
        "AUTO_OPEN_TEMPLATE": DEFAULT_CONFIG["AUTO_OPEN_TEMPLATE"],
        "AUTO_FILL_TEMPLATE": DEFAULT_CONFIG["AUTO_FILL_TEMPLATE"],
        "CLOSE_XLSX_AFTER_OPEN": DEFAULT_CONFIG["CLOSE_XLSX_AFTER_OPEN"],
        "AUTO_IMPORT_TEMPLATE": DEFAULT_CONFIG["AUTO_IMPORT_TEMPLATE"],
    }
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            for k in DEFAULT_CONFIG:
                if k in data and data[k] not in (None, ""):
                    cfg[k] = data[k]
    except (FileNotFoundError, json.JSONDecodeError):
        pass
    try:
        cfg["DEBUG_PORT"] = int(cfg["DEBUG_PORT"])
    except (TypeError, ValueError):
        cfg["DEBUG_PORT"] = DEFAULT_CONFIG["DEBUG_PORT"]
    # 密码统一解成明文（兼容旧 config 里的明文、以及新 config 的 b64: 前缀）
    cfg["PASSWORD"] = _decode_pwd(cfg.get("PASSWORD", ""))
    return cfg

CONFIG: _AppConfig = load_config()

def save_config(cfg: _AppConfig) -> None:
    # 持久化前把密码编成 b64:xxx，gui_config.json 里不直接看明文
    try:
        to_save = dict(cfg)
        if to_save.get("PASSWORD"):
            to_save["PASSWORD"] = _encode_pwd(to_save["PASSWORD"])
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(to_save, f, ensure_ascii=False, indent=2)
        print("[+] 配置已保存到: " + CONFIG_FILE)
    except OSError as e:
        print("[!] 保存配置失败: " + str(e))

def is_exe_path_set() -> bool:
    """软件路径是否已配置且文件存在。"""
    return bool(CONFIG["EXE_PATH"]) and os.path.isfile(CONFIG["EXE_PATH"])
# ====================================================

# ===================== 运行状态机 =====================
# 由后台线程通过 set_state 切换；主线程的 _update_ui 每 250ms 轮询刷新界面。
IDLE, RUNNING, PAUSED, STATE_WAIT_CAPTCHA, STATE_SOLVE_CAPTCHA, STATE_WAIT_EMAIL_VERIFY, FINISHED, ERROR = range(8)
current_state = IDLE
# 保护 on_start 与 run_login_flow 初始态检查之间的竞态：
# 避免极快双击「启动」导致两个 run_login_flow 同时跑起来
_START_LOCK = threading.Lock()
# 专用「流程是否在跑」标志 —— 与 current_state 解耦。
# current_state 会在 on_start 里提前置 RUNNING 来防竞态，
# 所以 run_login_flow 内部不能再用 current_state 自己拦自己；
# 改成用这个独立标志，由 _START_LOCK 保护原子读写。
_FLOW_ACTIVE = False
app_proc = None
cdp_client = None
_ocr = None

def set_state(s):
    """线程安全地切换全局状态。"""
    global current_state
    current_state = s


# 「等待邮箱安全验证」的等待信息（后台线程写、主线程 _update_ui 每 250ms 读，
# 只存简单标量，读写天然安全；不在后台线程里直接操作 tkinter 控件）
EMAIL_VERIFY_WAIT = {"t0": 0.0, "active": False}
# ====================================================


# ===================== 流程步骤跟踪 =====================
# 目的：日志能完整显示「当前进行到第几步 / 这一步叫什么 / 成功还是失败」，
#      出问题时一眼就能看出卡在哪一步。
FLOW_STATE: dict = {
    "title": "",
    "index": 0,
    "total": 0,
    "t0": 0.0,
    "step_t0": 0.0,
    "current": "",
}


def flow_begin(title, step_names):
    """开始一个流程，并先把全部步骤列出来（便于对照进度）。"""
    FLOW_STATE.clear()
    FLOW_STATE.update({
        "title": title,
        "index": 0,
        "total": len(step_names),
        "t0": time.time(),
        "step_t0": time.time(),
        "current": "",
    })
    print("")
    print("=" * 68)
    print("[流程] ▶ 开始：" + title + "（共 " + str(len(step_names)) + " 步）")
    for i, name in enumerate(step_names, 1):
        print("[流程]      %d) %s" % (i, name))
    print("=" * 68)


def flow_step_start(name):
    """标记某一步开始。"""
    FLOW_STATE["index"] += 1
    FLOW_STATE["current"] = name
    FLOW_STATE["step_t0"] = time.time()
    total = max(FLOW_STATE.get("total", 0), FLOW_STATE["index"])
    print("[步骤 %d/%d] ▶ %s" % (FLOW_STATE["index"], total, name))


def flow_step_end(ok=True, detail=""):
    """标记某一步结束。"""
    idx = FLOW_STATE.get("index", 0)
    total = max(FLOW_STATE.get("total", 0), idx)
    name = FLOW_STATE.get("current", "")
    cost = time.time() - FLOW_STATE.get("step_t0", time.time())
    line = "[步骤 %d/%d] %s %s（耗时 %.1fs）" % (
        idx, total, "✔ 完成" if ok else "✖ 失败", name, cost)
    if detail:
        line += " | " + detail
    print(line)


def flow_summary(ok=True, detail=""):
    """流程结束时的总结。"""
    cost = time.time() - FLOW_STATE.get("t0", time.time())
    line = "[流程] " + ("✅ 全部完成" if ok else "❌ 中断") + "：" + \
           FLOW_STATE.get("title", "") + "，已完成 %d 步，总耗时 %.1fs" % (
               FLOW_STATE.get("index", 0), cost)
    if detail:
        line += " | " + detail
    print(line)


# 步骤返回该哨兵值表示「本轮流程无需继续」：日志仍记为完成，
# 调用方随即正常结束流程（不报错、也不执行后面的步骤）。
# 典型场景：接口没有「要上传」的数据，再往下写表/导入/上传都没意义。
FLOW_STOP = object()


def run_step(name, fn):
    """执行一步并输出统一样式的日志。

    Returns:
        True       步骤成功，继续下一步；
        FLOW_STOP  步骤成功但本轮流程正常结束（不继续后续步骤）；
        False      步骤失败。
    """
    flow_step_start(name)
    try:
        res = fn()
    except Exception as e:
        flow_step_end(False, "异常: " + repr(e))
        raise
    if res is FLOW_STOP:
        flow_step_end(True, "本轮无需继续，正常结束")
        return FLOW_STOP
    if res is False:
        flow_step_end(False, "返回 False")
        return False
    flow_step_end(True)
    return True
# ====================================================

# 在登录页执行的 JS：填充表单（兼容 Vue3 响应式，需用原生 setter + input 事件）
JS_FILL = r"""
(function (cfg) {
    function setNativeValue(el, value) {
        const setter = Object.getOwnPropertyDescriptor(el, 'value')
            ? Object.getOwnPropertyDescriptor(el, 'value').set : null;
        const proto = Object.getPrototypeOf(el);
        const protoSetter = Object.getOwnPropertyDescriptor(proto, 'value').set;
        (setter && setter !== protoSetter ? setter : protoSetter).call(el, value);
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
    }

    const userInput = document.querySelector('input[placeholder="请输入用户名或手机号码"]');
    const pwdInput  = document.querySelector('input[placeholder="请输入密码"]');
    if (!userInput || !pwdInput) {
        return { ok: false, msg: '未找到登录表单（可能不在登录页）' };
    }
    setNativeValue(userInput, cfg.username);
    setNativeValue(pwdInput, cfg.password);

    // 勾选"记住账号"复选框（Arco Design: label.arco-checkbox）
    let checked = false;
    const labels = document.querySelectorAll('.arco-checkbox');
    for (const lb of labels) {
        if (lb.textContent.includes('记住账号')) {
            const box = lb.querySelector('input[type="checkbox"]');
            if (box && !box.checked) lb.click();
            checked = true;
            break;
        }
    }
    return { ok: true, checkboxChecked: checked };
})(%s)
"""

# 检测验证码是否已填写
JS_CAPTCHA_FILLED = r"""
(function () {
    const el = document.querySelector('input[placeholder="请输入验证码"]');
    return !!(el && el.value && el.value.trim().length > 0);
})()
"""


# ===================== 验证码识别（OCR） =====================
class _CaptchaOCR:
    """基于 ddddocr 的图形验证码识别器。

    针对本系统验证码（80x40、彩色扭曲字符 + 干扰线）的有效做法：
      1) 多种预处理变体（补白 / 灰度 / 二值化 / 放大）分别识别；
      2) 对每个字符位置按出现次数+置信度加权投票，取最优串；
      3) 再用"次优字符"做 1~2 位替换扩展出候选答案。
    实测单张图首选命中率约 93%，候选集合命中率约 100%。
    """

    def __init__(self):
        self.models = []
        if not _OCR_AVAILABLE:
            return
        for name, kw in (("std", {}), ("beta", {"beta": True})):
            try:
                self.models.append((name, ddddocr.DdddOcr(show_ad=False, **kw)))
            except Exception as e:
                print("[!] OCR 模型(" + name + ")初始化失败: " + repr(e))

    @property
    def ready(self) -> bool:
        return bool(self.models)

    @staticmethod
    def _png(img) -> bytes:
        buf = io.BytesIO()
        img.convert("RGB").save(buf, format="PNG")
        return buf.getvalue()

    @staticmethod
    def _pad(img, pw: int):
        p = Image.new("RGB", (img.width + pw, img.height), "white")
        p.paste(img, (0, 0))
        return p

    def _variants(self, img):
        out = []
        g = img.convert("L")
        for pw in CAPTCHA_PADS:
            out.append(self._pad(img, pw))
            out.append(self._pad(g.convert("RGB"), pw))
        for pw in (24, 48, 80):
            out.append(self._pad(g.point(lambda p: 0 if p < 150 else 255).convert("RGB"), pw))
        for sc in (2, 3):
            big = img.resize((img.width * sc, img.height * sc), Image.LANCZOS)
            for pw in (0, 24, 48):
                out.append(self._pad(big, pw))
        return out

    def recognize(self, img, max_cands: int = CAPTCHA_MAX_CANDIDATES):
        """返回按可靠性排序的候选验证码字符串列表（服务端校验不区分大小写）。"""
        outputs = []
        for _name, ocr in self.models:
            for vim in self._variants(img):
                try:
                    r = ocr.classification(self._png(vim), probability=True)
                except Exception:
                    continue
                if isinstance(r, dict):
                    text = (r.get("text") or "").strip()
                    conf = float(r.get("confidence") or 0.0)
                else:
                    text = (r or "").strip()
                    conf = 0.5
                if text:
                    outputs.append((text, conf))
        if not outputs:
            return []

        # 众数长度：验证码一般为 4~5 位
        lens = Counter(len(t) for t, _ in outputs)
        n_len = lens.most_common(1)[0][0]
        usable = [(t, c) for t, c in outputs if len(t) == n_len] or outputs
        n_len = max(len(t) for t, _ in usable)

        # 逐位字符权重（大小写归一，服务端不区分大小写）
        pos = []
        for i in range(n_len):
            w = Counter()
            for t, c in usable:
                if i < len(t):
                    w[t[i].lower()] += 1.0 + c
            pos.append([ch for ch, _ in w.most_common()])

        best = "".join(p[0] if p else "" for p in pos)
        cands = [best] if best else []

        def expand(base, depth):
            res = []
            idxs = [i for i in range(len(base)) if len(pos[i]) > 1]
            for i in idxs:
                for alt in pos[i][1:]:
                    s = base[:i] + alt + base[i + 1:]
                    if s not in cands and s not in res:
                        res.append(s)
            if depth > 1:
                for i in idxs:
                    for alt in pos[i][1:]:
                        b1 = base[:i] + alt + base[i + 1:]
                        for j in idxs:
                            if j == i:
                                continue
                            for alt2 in pos[j][1:]:
                                s = b1[:j] + alt2 + b1[j + 1:]
                                if s not in cands and s not in res:
                                    res.append(s)
            return res

        cands += expand(best, 1)
        cands += expand(best, 2)
        for t, _c in sorted(outputs, key=lambda x: -x[1]):
            if t not in cands:
                cands.append(t)
        return cands[:max_cands]


def get_ocr():
    """懒加载 OCR 器（ddddocr 首次推理较慢，真正需要时才初始化）。"""
    global _ocr
    if _ocr is None:
        _ocr = _CaptchaOCR()
        if _ocr.ready:
            print("[+] 验证码 OCR 引擎就绪")
    return _ocr


# ===================== 验证码相关 JS =====================
# 抓取验证码图片：canvas -> base64 dataURL
JS_GET_CAPTCHA_IMG = r"""
(function () {
    const img = document.querySelector('.arco-image.code img') || document.querySelector('img.arco-image-img');
    if (!img) return { ok: false, msg: '未找到验证码图片' };
    if (!img.complete || !img.naturalWidth) return { ok: false, msg: '验证码图片尚未加载完成' };
    try {
        const c = document.createElement('canvas');
        c.width = img.naturalWidth; c.height = img.naturalHeight;
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        return { ok: true, data: c.toDataURL('image/png') };
    } catch (e) {
        return { ok: false, msg: String(e) };
    }
})()
"""

# 捕获验证码 uuid（服务端用它把图片和答案关联起来）
JS_PATCH_UUID = r"""
(function () {
    if (window.__captchaKeys) return 'already';
    window.__captchaKeys = [];
    const orig = crypto.randomUUID.bind(crypto);
    Object.defineProperty(crypto, 'randomUUID', {
        configurable: true, writable: true,
        value: function () { const u = orig(); window.__captchaKeys.push(u); return u; }
    });
    return 'patched';
})()
"""

# 读取验证码当前状态：图片是否就绪 / blob 地址 / 对应 uuid
JS_CAPTCHA_STATE = r"""
(function () {
    const img = document.querySelector('.arco-image.code img') || document.querySelector('img.arco-image-img');
    const keys = window.__captchaKeys || [];
    return {
        hasImg: !!img,
        complete: !!(img && img.complete && img.naturalWidth > 0),
        src: img ? (img.src || '') : '',
        nkeys: keys.length,
        key: keys.length ? keys[keys.length - 1] : ''
    };
})()
"""

# 点击验证码图片刷新（前端组件自带 onClick 刷新逻辑）
JS_REFRESH_CAPTCHA = r"""
(function () {
    const el = document.querySelector('.arco-image.code') || document.querySelector('img.arco-image-img');
    if (!el) return 'not found';
    el.click();
    return 'clicked';
})()
"""

# 填写验证码输入框（Vue3 需用原生 setter + input 事件）
JS_FILL_CAPTCHA = r"""
(function (v) {
    function setNativeValue(el, value) {
        const setter = Object.getOwnPropertyDescriptor(el, 'value')
            ? Object.getOwnPropertyDescriptor(el, 'value').set : null;
        const proto = Object.getPrototypeOf(el);
        const protoSetter = Object.getOwnPropertyDescriptor(proto, 'value').set;
        (setter && setter !== protoSetter ? setter : protoSetter).call(el, value);
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
    }
    const el = document.querySelector('input[placeholder="请输入验证码"]');
    if (!el) return { ok: false, msg: '未找到验证码输入框' };
    setNativeValue(el, v);
    return { ok: true };
})(%s)
"""

# 读取登录错误提示
JS_LOGIN_MESSAGE = r"""
(function () {
    const el = document.querySelector('.error-message');
    return el ? (el.innerText || '').trim() : '';
})()
"""

# 用应用自身的登录接口校验候选验证码（同一张图可多次校验；服务端不区分大小写）
JS_VERIFY_CAPTCHA = r"""
(async function (payload) {
    try {
        if (!window.electronAPI || !window.electronAPI.loginRequest) return { unsupported: true };
        const r = await window.electronAPI.loginRequest({
            url: '/api/oauth/clientLogin',
            data: payload,
            headers: { 'Content-Type': 'application/json' }
        });
        const d = (r && r.data) || null;
        return { code: d ? d.code : null, msg: d ? d.msg : '' };
    } catch (e) {
        return { error: String(e && e.message || e) };
    }
})(%s)
"""


# 检测「首次启动配置」弹窗（首次登录进入首页时弹出，要求设置备份路径）。
# 优先按 Arco Design 的 .arco-modal 容器检测，并以文案兜底；
# 若实际 DOM 用了别的容器类名，请在此调整选择器。
JS_FIRST_LAUNCH_MODAL = r"""
(function () {
    const containers = ['.arco-modal', '.arco-modal-wrapper', '.arco-drawer'];
    for (const sel of containers) {
        const node = document.querySelector(sel);
        if (node) {
            const txt = node.innerText || '';
            if (txt.includes('首次启动配置') || txt.includes('备份路径')) {
                return { found: true, text: txt.slice(0, 100) };
            }
        }
    }
    const bodyTxt = document.body.innerText || '';
    if (bodyTxt.includes('首次启动配置') && bodyTxt.includes('备份路径')) {
        return { found: true, text: bodyTxt.slice(0, 100) };
    }
    return { found: false };
})()
"""

# 探测当前页面处于「登录页」还是「登录后（首页）」状态
JS_CHECK_LOGIN_STATE = r"""
(function () {
    const hasLoginForm = !!document.querySelector('input[placeholder="请输入用户名或手机号码"]')
        || !!document.querySelector('input[placeholder="请输入验证码"]');
    function textOf(el) { return ((el.innerText || el.textContent || '') + '').trim(); }
    const nodes = Array.from(document.querySelectorAll('button, a, div, span'));
    const hasTemplateBtn = nodes.some(n => textOf(n) === '下载表格模板');
    return {
        hasLoginForm: hasLoginForm,
        hasTemplateBtn: hasTemplateBtn,
        url: (location && location.href) || '',
        title: document.title || '',
        ready: document.readyState
    };
})()
"""

# 探测当前页面是否处于「邮箱安全验证」页（只读探测，不做任何点击/输入）
# 判据覆盖截图里的特征：标题「邮箱安全验证」、输入框「请输入邮箱验证码」、
# 按钮「获取验证码 / 验证并登录」，同时带回登录表单与首页标志用于判断是否已验证完成。
JS_CHECK_EMAIL_VERIFY = r"""
(function () {
    function textOf(el) { return ((el.innerText || el.textContent || '') + '').trim(); }
    const bodyTxt = (document.body && (document.body.innerText || '')) || '';
    const nodes = Array.from(document.querySelectorAll('button, a, div, span, label'));
    return {
        // 邮箱安全验证页特征
        hasVerifyInput: !!document.querySelector('input[placeholder="请输入邮箱验证码"]'),
        hasVerifyTitle: bodyTxt.includes('邮箱安全验证') || (document.title || '').includes('邮箱安全验证'),
        hasVerifyBtn: nodes.some(n => textOf(n) === '验证并登录') || bodyTxt.includes('验证并登录'),
        hasGetCodeBtn: nodes.some(n => textOf(n) === '获取验证码') || bodyTxt.includes('获取验证码'),
        // 完成后的标志
        hasTemplateBtn: nodes.some(n => textOf(n) === '下载表格模板'),
        hasLoginForm: !!document.querySelector('input[placeholder="请输入用户名或手机号码"]')
            || !!document.querySelector('input[placeholder="请输入验证码"]'),
        url: (location && location.href) || '',
        title: document.title || '',
        ready: document.readyState
    };
})()
"""


def _email_verify_evidence(st):
    """组合多种判据判断是否为「邮箱安全验证」页，返回 (证据条数, 是否判定为验证页)。

    单一线索容易误判（例如正文顺带提到「验证」），因此要求至少命中 2 条
    相互独立的特征（输入框 / 标题 / 验证并登录按钮 / 获取验证码按钮）。
    """
    if not isinstance(st, dict):
        return 0, False
    title = st.get("title") or ""
    url = (st.get("url") or "").lower()
    ev = 0
    if st.get("hasVerifyInput"):
        ev += 1                      # 证据1：输入框 placeholder「请输入邮箱验证码」
    if st.get("hasVerifyTitle") or "邮箱安全验证" in title:
        ev += 1                      # 证据2：正文或页面标题含「邮箱安全验证」
    if st.get("hasVerifyBtn"):
        ev += 1                      # 证据3：按钮「验证并登录」
    if st.get("hasGetCodeBtn"):
        ev += 1                      # 证据4：按钮「获取验证码」
    if ev < 2 and ("verify" in url or "email" in url) and st.get("hasVerifyInput"):
        ev = 2                       # URL 指向验证页且有验证输入框，视为满足双证据
    return ev, ev >= 2


def _probe_page_state(target):
    """探测单个页面 target 的状态。

    返回 ("logged_in", target) | ("login_page", None) | ("unknown", None)
    """
    ws = target.get("webSocketDebuggerUrl")
    if not ws:
        return "unknown", None
    client = None
    st = None
    try:
        client = CDPClient(ws)
        st = client.eval_js(JS_CHECK_LOGIN_STATE)
    except Exception as e:
        print("[!] 探测页面状态失败(" + str(target.get("url", "")) + "): " + repr(e))
    finally:
        if client is not None:
            try:
                client.close()
            except Exception:
                pass
    if not isinstance(st, dict):
        return "unknown", None

    print("[*] 页面状态: " + json.dumps(st, ensure_ascii=False))
    if st.get("hasTemplateBtn"):
        # 首页上有「下载表格模板」按钮 => 已登录
        return "logged_in", target
    if st.get("hasLoginForm"):
        # 这一页是登录页
        return "login_page", None
    # 既没有登录表单也没有模板按钮：页面可能还没渲染完，先按「未知」处理
    if st.get("ready") == "complete" and "login" not in (st.get("url") or "").lower():
        return "logged_in", target
    return "unknown", None


def probe_login_state(max_wait=10):
    """探测应用当前是否已处于登录后（首页）状态。

    返回命中的 target（登录后页面）或 None（仍在登录页 / 无法判定）。
    """
    global detected_target
    detected_target = None

    deadline = time.time() + max_wait
    while True:
        try:
            r = requests.get(f"http://127.0.0.1:{CONFIG['DEBUG_PORT']}/json", timeout=5)
            targets = r.json()
        except Exception as e:
            print("[!] 获取页面列表失败: " + repr(e))
            return None

        pages = [t for t in targets if t.get("type") == "page"]
        if not pages:
            print("[!] 未找到页面 target，无法判断登录状态")
            return None

        # 先看非登录页，再看登录页
        ordered = sorted(pages, key=lambda t: 1 if "login" in (t.get("url") or "").lower() else 0)

        has_login_page = False
        for t in ordered:
            status, hit = _probe_page_state(t)
            if status == "logged_in":
                detected_target = hit
                print("[+] 检测到应用已在登录后（首页）状态。")
                return hit
            if status == "login_page":
                has_login_page = True

        if has_login_page:
            print("[*] 应用尚未登录（停留在登录页）。")
            return None

        # 页面还没渲染完，稍等片刻再探测
        if time.time() >= deadline:
            print("[!] 未能确认应用是否已登录，将按「未登录」处理。")
            return None
        time.sleep(1.0)


# ===================== 应用运行状态检测 =====================
# 应用（Electron）以 --remote-debugging-port 启动后，只要该调试端口能访问，
# 就说明应用已经打开（可能是用户手动打开的）。因此点击「启动」时先探测一次：
#   1) 应用未打开  -> 正常启动应用 -> 走登录流程；
#   2) 应用已打开  -> 复用该实例（不再重复启动），必要时继续登录流程；
#   3) 应用已打开且已在首页 -> 跳过「打开应用」与「登录」，直接跑登录后的逻辑。

_APP_RUNNING_CACHE = {"ts": 0.0, "value": False}
APP_RUNNING_CACHE_TTL = 3.0

# 自动找空闲端口时的探测范围：从配置端口起最多往后探多少个
PORT_SCAN_MAX_TRIES = 200

# 探测到的「已登录页面」target（登录后复用同一个页面连接，避免连错页）
detected_target = None


def _is_port_free(port: int) -> bool:
    """用 socket.bind 检测本机端口是否空闲（未被任何进程占用）。

    注意：bind 检测与实际启动之间存在 TOCTOU 窗口，但对本场景足够；
          Electron 启动时若端口被占会自动选别的端口，这里只是尽量减少冲突。
    """
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", int(port)))
        return True
    except OSError:
        return False
    finally:
        try:
            s.close()
        except Exception:
            pass


def _find_free_port(start_port: int, max_tries: int = PORT_SCAN_MAX_TRIES) -> int:
    """从 start_port 开始往后探测，返回第一个空闲端口。

    Args:
        start_port: 起始端口（不含自身，从 start_port+1 开始探，
                    因为调用方通常已经确认 start_port 被占）。
        max_tries: 最多探测多少个端口。

    Returns:
        第一个空闲端口。

    Raises:
        RuntimeError: 范围内未找到空闲端口。
    """
    for port in range(int(start_port) + 1, int(start_port) + 1 + max_tries):
        if _is_port_free(port):
            return port
    raise RuntimeError(
        "从端口 %d 起往后 %d 个端口均被占用，请手动在「设置」里换一个调试端口。"
        % (start_port, max_tries))


def _is_our_app_on_port(port: int) -> bool:
    """确认指定 CDP 端口上跑的是不是「我们的上传系统」应用。

    判断依据：连接该端口的所有 page target，只要有一个页面包含
    我们的登录表单（placeholder=请输入用户名或手机号码）或
    首页「下载表格模板」按钮，就认为是本应用。

    用于区分「端口被其它 Chromium 程序（如 Chrome）占用」与
    「端口被我们自己的应用占用」。
    """
    try:
        r = requests.get(f"http://127.0.0.1:{port}/json", timeout=3)
        if not r.ok:
            return False
        targets = r.json()
    except Exception:
        return False

    pages = [t for t in targets if t.get("type") == "page"]
    if not pages:
        return False

    for t in pages:
        ws = t.get("webSocketDebuggerUrl")
        if not ws:
            continue
        client = None
        try:
            client = CDPClient(ws)
            st = client.eval_js(JS_CHECK_LOGIN_STATE)
        except Exception:
            continue
        finally:
            if client is not None:
                try:
                    client.close()
                except Exception:
                    pass
        if isinstance(st, dict) and (st.get("hasTemplateBtn") or st.get("hasLoginForm")):
            return True
    return False


def _port_listen_pids(port: int) -> set[int]:
    """返回正在监听指定本机 TCP 端口的进程 PID 集合（netstat -ano）。

    中文版 Windows 的 netstat 状态列可能是「正在侦听」而不是 LISTENING，
    因此不单纯依赖状态文本：远端地址为 *:0 的监听行同样算命中。
    netstat 不可用 / 解析失败时返回空集合。
    """
    if not _WIN32:
        return set()
    try:
        out = subprocess.check_output(
            ["netstat", "-ano", "-p", "tcp"],
            stderr=subprocess.DEVNULL, timeout=5,
        ).decode("gbk", errors="ignore")
    except Exception:
        return set()

    pids: set[int] = set()
    for line in out.splitlines():
        parts = line.split()
        if len(parts) < 5 or parts[0].upper() != "TCP":
            continue
        local, remote, state, pid_s = parts[1], parts[2], parts[3], parts[4]
        if "LISTEN" not in state.upper() and not remote.endswith(":0"):
            continue
        m = re.search(r":(\d+)$", local)
        if not m or int(m.group(1)) != int(port):
            continue
        try:
            pid = int(pid_s)
            if pid > 0:
                pids.add(pid)
        except ValueError:
            continue
    return pids


def _tasklist_image_name_by_pid(pid: int) -> str:
    """按 PID 取映像名（_win_process_exe 查不到进程路径时的兜底）。

    /FO CSV 避免表格列宽把长中文 exe 名截断；tasklist 在代码页 65001 下
    输出 UTF-8、在 936 下输出 GBK，两种编码都试一遍。
    """
    if not _WIN32 or not pid:
        return ""
    try:
        raw = subprocess.check_output(
            ["tasklist", "/FI", "PID eq %d" % int(pid), "/NH", "/FO", "CSV"],
            stderr=subprocess.DEVNULL, timeout=5,
        )
    except Exception:
        return ""
    out = ""
    for enc in ("utf-8", "gbk"):
        try:
            out = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if not out:
        out = raw.decode("gbk", errors="ignore")
    for line in out.splitlines():
        s = line.strip()
        if not s or "INFO:" in s:
            continue
        parts = s.split('","')
        if len(parts) < 2:
            continue
        name = parts[0].strip().strip('"').lower()
        got_pid = parts[1].strip().strip('"')
        if name.endswith(".exe") and got_pid == str(pid):
            return name
    return ""


def _is_port_held_by_our_exe(port: int) -> bool:
    """判断占用该调试端口的进程，是不是「当前已经打开的上传系统 exe」。

    端口被占用有两种可能：
      1) 当前已打开、已挂载 --remote-debugging-port 的本应用占着 → 沿用该端口；
      2) 其它程序占着 → 才需要另找空闲端口。

    这里不依赖页面内容识别（_is_our_app_on_port）：页面可能停在邮箱验证页、
    首次配置弹窗或还在加载中，页面特征探测会失败，但端口确实是本应用占的。
    """
    exe_path = (CONFIG.get("EXE_PATH") or "").strip()
    if not exe_path:
        return False
    exe_name = os.path.basename(exe_path).lower()
    target_path = os.path.normcase(os.path.abspath(exe_path))

    # _win_process_exe / _is_exe_process_running 定义在文件后面，
    # 运行时按名字查找即可（模块加载完毕后才会调用）。
    pids = _port_listen_pids(int(port))
    for pid in pids:
        proc_path = _win_process_exe(pid)
        if proc_path:
            if (os.path.normcase(os.path.abspath(proc_path)) == target_path
                    or os.path.basename(proc_path).lower() == exe_name):
                return True
        elif _tasklist_image_name_by_pid(pid) == exe_name:
            return True
    if pids:
        # 能查到监听 PID，但都不是配置的 exe → 确实是其它程序占用
        return False

    # netstat 解析失败 / 无权限时退回宽松判据：CDP 端点有响应 + 配置的 exe 进程在跑
    try:
        r = requests.get(f"http://127.0.0.1:{int(port)}/json/version", timeout=1.5)
        cdp_ok = bool(r.ok)
    except Exception:
        cdp_ok = False
    return cdp_ok and _is_exe_process_running()


def _ensure_free_debug_port() -> int:
    """启动 exe 前确认调试端口，只有「没有挂载本应用 exe」时才找空闲端口。

    判断顺序：
      1) 端口上是本应用的页面（_is_our_app_on_port）→ 已打开，直接复用原端口；
      2) 端口被占，但占用者就是当前已打开、已挂载调试端口的 exe
         （_is_port_held_by_our_exe）→ 同样复用原端口，不找空闲端口；
      3) 端口空闲 → 直接用；
      4) 端口被其它程序占用（没有挂载本应用 exe）→ 才往后找第一个空闲端口，
         写回 CONFIG 并持久化到 gui_config.json。

    返回最终选定的端口号。
    """
    port = int(CONFIG["DEBUG_PORT"])

    # 1) 端口上若已经是我们的应用，直接复用，不换端口
    if _is_our_app_on_port(port):
        print("[*] 端口 " + str(port) + " 上运行的是本应用，继续复用该端口。")
        return port

    if _is_port_free(port):
        # 3) 端口空闲，直接用于启动应用
        print("[*] 端口 " + str(port) + " 空闲，将用于启动应用。")
        return port

    # 2) 端口被占不等于冲突：可能就是当前已经打开的程序占的，
    #    此时沿用原端口，既不找空闲端口，也不改写配置（否则会连错端口）。
    if _is_port_held_by_our_exe(port):
        print("[*] 端口 " + str(port)
              + " 被当前已打开的上传系统占用（已挂载调试端口），继续复用该端口，不寻找空闲端口。")
        return port

    # 4) 确认没有挂载本应用 exe，端口是被其它程序占用，才去找空闲端口
    print("[!] 端口 " + str(port) + " 被其它程序占用（未挂载本应用），自动寻找空闲端口 ...")
    new_port = _find_free_port(port)
    print("[*] 已选定空闲端口 " + str(new_port) + "（原端口 " + str(port) + " 被占用）。")
    CONFIG["DEBUG_PORT"] = new_port
    # 端口变了，缓存的应用运行状态作废
    _APP_RUNNING_CACHE["ts"] = 0.0
    _APP_RUNNING_CACHE["value"] = False
    save_config(CONFIG)
    return new_port


def _is_exe_process_running() -> bool:
    """用 tasklist 检测 CONFIG['EXE_PATH'] 对应的进程是否已存在。

    场景：用户可能手动双击 exe 打开了应用（没带 --remote-debugging-port），
    此时调试端口不会响应，但进程确实在跑。本函数补上这一层检测，
    避免脚本误以为应用没打开而再启动一个新实例。
    """
    exe_path = CONFIG.get("EXE_PATH", "")
    if not exe_path:
        return False
    exe_name = os.path.basename(exe_path).strip()
    if not exe_name:
        return False
    try:
        # tasklist /FI "IMAGENAME eq xxx.exe"  /NH  （/NH 不显示表头）
        out = subprocess.check_output(
            ["tasklist", "/FI", "IMAGENAME eq " + exe_name, "/NH"],
            stderr=subprocess.DEVNULL, timeout=5,
        ).decode("gbk", errors="ignore")
    except Exception:
        return False
    # tasklist 找不到时会输出 "INFO: 没有运行的任务匹配指定标准。" 或英文 "INFO: No tasks are running..."
    # 只要输出里不含 "INFO:" 就说明找到了进程
    if not out.strip() or "INFO:" in out:
        return False
    return True


def is_app_running(use_cache: bool = True) -> bool:
    """检测上传系统应用是否已经打开（调试端口是否可访问）。"""
    if use_cache:
        ts = _APP_RUNNING_CACHE.get("ts", 0.0)
        if ts and (time.time() - ts) < APP_RUNNING_CACHE_TTL:
            return bool(_APP_RUNNING_CACHE.get("value"))

    ok = False
    try:
        r = requests.get(f"http://127.0.0.1:{CONFIG['DEBUG_PORT']}/json/version", timeout=1.5)
        ok = bool(r.ok)
    except requests.RequestException:
        ok = False
    except Exception as e:
        print("[!] 检测应用运行状态失败: " + repr(e))
        ok = False

    _APP_RUNNING_CACHE["ts"] = time.time()
    _APP_RUNNING_CACHE["value"] = ok
    return ok


_APP_PROBE_STATE = {"started": False}


def _start_app_probe_worker():
    """后台线程定时刷新「应用是否运行」缓存，主线程只读缓存。

    IDLE 时 _update_ui 每 250ms 要显示状态，若直接在主线程调
    is_app_running()，端口不通时 requests 最长会阻塞 1.5 秒，界面像卡死。
    """
    if _APP_PROBE_STATE["started"]:
        return
    _APP_PROBE_STATE["started"] = True

    def loop():
        while True:
            try:
                _ = is_app_running(use_cache=False)
            except Exception:
                pass
            time.sleep(APP_RUNNING_CACHE_TTL)

    threading.Thread(target=loop, name="app-running-probe", daemon=True).start()


def launch_app():
    """以远程调试模式启动 exe"""
    print("[*] 启动应用: " + CONFIG["EXE_PATH"])
    app_proc = subprocess.Popen([
        CONFIG["EXE_PATH"],
        f"--remote-debugging-port={CONFIG['DEBUG_PORT']}",
        "--remote-allow-origins=*"
    ])
    deadline = time.time() + STARTUP_TIMEOUT
    while time.time() < deadline:
        try:
            r = requests.get(f"http://127.0.0.1:{CONFIG['DEBUG_PORT']}/json/version", timeout=2)
            if r.ok:
                print("[+] 调试端口已就绪: " + str(r.json().get("Browser")))
                _APP_RUNNING_CACHE["ts"] = time.time()
                _APP_RUNNING_CACHE["value"] = True
                return app_proc
        except requests.RequestException:
            pass
        time.sleep(0.5)
    print("[!] 等待调试端口超时，应用可能启动失败或不支持远程调试")
    return app_proc


def find_login_page(prefer_logged_in=False):
    """找到登录页的 page target

    prefer_logged_in=True 时优先返回不是登录页的 target（登录后重连用）。

    抛 RuntimeError 时会带「根本原因」：
      - 连接失败 → 应用没起来 / 端口被其它程序占用
      - 超时 → 应用起来了但响应慢
      - 非 JSON → 端口上不是我们的 CDP endpoint
      - 无 page target → Electron 还没打开首页
    """
    url = "http://127.0.0.1:%d/json" % CONFIG["DEBUG_PORT"]
    try:
        r = requests.get(url, timeout=5)
    except requests.exceptions.ConnectionError as e:
        raise RuntimeError(
            "连接 CDP 端口失败（%s）：应用可能没启动，"
            "或端口 %d 被其它程序占用。原始错误: %s"
            % (url, CONFIG["DEBUG_PORT"], repr(e)))
    except requests.exceptions.Timeout as e:
        raise RuntimeError(
            "连接 CDP 端口超时（%s）：应用启动了但 5 秒内没响应。原始错误: %s"
            % (url, repr(e)))
    except requests.exceptions.RequestException as e:
        raise RuntimeError(
            "请求 CDP 端口异常（%s）：%s" % (url, repr(e)))

    try:
        targets = r.json()
    except ValueError as e:
        raise RuntimeError(
            "CDP 端口 %d 返回了非 JSON 数据（HTTP %d），"
            "可能端口上跑的不是 Electron CDP。原始错误: %s"
            % (CONFIG["DEBUG_PORT"], r.status_code, repr(e)))

    pages = [t for t in targets if t.get("type") == "page"]
    if not pages:
        raise RuntimeError(
            "CDP 已连通但没有 type=page 的 target（Electron 可能还没打开首页）。"
            "当前 targets 类型: " + str([t.get("type") for t in targets]))
    if prefer_logged_in:
        for t in pages:
            if "login" not in (t.get("url") or "").lower():
                return t
    return pages[0]


def wait_for_login_form(client, timeout=30):
    """等待登录表单渲染完成"""
    print("[*] 等待登录表单加载 ...")
    deadline = time.time() + timeout
    while time.time() < deadline:
        ok = client.eval_js(
            "!!document.querySelector('input[placeholder=\"请输入用户名或手机号码\"]')"
        )
        if ok:
            print("[+] 登录表单已出现")
            return True
        time.sleep(0.5)
    return False


def _wait_until_resumed():
    """等待从暂停状态恢复。

    本函数只在后台线程里被 run_login_flow 调用；若 UI 主线程意外调到，
    会立刻返回避免卡死 Tkinter 事件循环。
    """
    if threading.current_thread() is threading.main_thread():
        return
    while current_state == PAUSED:
        time.sleep(0.2)


def run_login_flow():
    """
    完整登录流程 + 首页自动流程，运行在后台线程。
    它会按步骤执行，并支持启动、暂停、继续。

    步骤清单会完整打印到日志里（含「第几步 / 这一步叫什么 / 成功还是失败」），
    登录成功后会自动继续执行 POST_LOGIN_STEPS（下载表格模板等）。

    启动前会先探测应用是否已经打开：
        - 未打开：启动应用 -> 登录 -> 登录后逻辑；
        - 已打开但未登录：复用该实例 -> 登录 -> 登录后逻辑；
        - 已打开且已登录：跳过打开与登录，直接执行登录后逻辑。
    """
    global app_proc, cdp_client, current_state, _FLOW_ACTIVE

    # 用 _FLOW_ACTIVE 做竞态双保险（与 on_start 里 set_state(RUNNING) 解耦）
    with _START_LOCK:
        if _FLOW_ACTIVE:
            return
        _FLOW_ACTIVE = True

    # ---------- 预检测：应用是否已经打开 / 是否已经登录 ----------
    # 两层检测：
    #   1) 调试端口 + CDP 页面内容（_is_our_app_on_port） —— 最准确，能确认是我们的应用
    #   2) 进程名匹配 tasklist（_is_exe_process_running） —— 兜底，检测用户手动打开（未带调试端口）的情况
    # 需要启动新实例时，_ensure_free_debug_port 还会再判一层：
    #   3) 端口占用者是不是当前已打开、已挂载调试端口的 exe（_is_port_held_by_our_exe）
    #      —— 是则沿用原端口，只有没挂载本应用时才去找空闲端口
    configured_port = int(CONFIG["DEBUG_PORT"])
    our_app_on_port = _is_our_app_on_port(configured_port)
    exe_proc_running = _is_exe_process_running()
    logged_in_target = None

    if our_app_on_port:
        print("[*] 检测到本应用已在端口 " + str(configured_port) + " 上运行，将复用该实例。")
        logged_in_target = probe_login_state()
    elif exe_proc_running:
        # 进程在跑但页面特征没识别到本应用：可能是手动双击打开的（没带调试端口），
        # 也可能是页面停在邮箱验证 / 加载中，端口其实已被这个已打开的程序占用
        print("[!] 检测到应用进程已在运行，但未能在调试端口上识别到本应用页面。")
        print("[!] 将启动一个新实例（带调试端口），请留意是否需要先关闭旧实例。")
        # 若端口其实已被这个已打开的程序占用，函数内部会沿用原端口，不找空闲端口
        _ensure_free_debug_port()
        print("[*] 将使用调试端口 " + str(CONFIG["DEBUG_PORT"]) + " 启动新的应用实例。")
    else:
        # 进程和端口都没找到 → 应用确实没打开
        _ensure_free_debug_port()
        print("[*] 将使用调试端口 " + str(CONFIG["DEBUG_PORT"]) + " 启动应用。")

    if logged_in_target is not None:
        # ===== 情况三：应用已打开且已登录 -> 跳过打开应用与登录流程 =====
        login_steps = [
            ("连接已打开的应用（CDP，复用现有页面）", _step_connect_running_app),
            ("处理「首次启动配置」弹窗", _step_handle_first_launch_modal),
        ]
        post_steps = _post_login_steps() if CONFIG.get("AUTO_DOWNLOAD_TEMPLATE", True) else []
        steps = list(login_steps) + post_steps
        title = "登录后流程（应用已打开并已登录，已跳过登录）"
    elif our_app_on_port:
        # ===== 情况二：应用已打开但未登录 -> 复用实例，继续登录 =====
        login_steps = [
            ("连接已打开的应用（CDP，复用现有页面）", _step_connect_running_app),
            ("等待登录表单加载", _step_wait_login_form),
            ("填写用户名 / 密码并勾选「记住账号」", _step_fill_form),
            ("自动识别图形验证码并提交登录", _step_solve_and_login),
            ("等待「邮箱安全验证」（人工完成后自动继续）", _step_wait_email_verify),
            ("处理「首次启动配置」弹窗", _step_handle_first_launch_modal),
        ]
        steps = list(login_steps)
        if CONFIG.get("AUTO_DOWNLOAD_TEMPLATE", True):
            steps += _post_login_steps()
        title = "登录 + 模板下载流程（应用已打开，复用实例）"
    else:
        # ===== 情况一：应用未打开 或 进程在跑但未带调试端口 -> 启动新实例 + 登录 =====
        login_steps = [
            ("启动上传系统应用（开启远程调试端口）", _step_launch_app),
            ("连接应用页面（CDP）", _step_find_login_page),
            ("等待登录表单加载", _step_wait_login_form),
            ("填写用户名 / 密码并勾选「记住账号」", _step_fill_form),
            ("自动识别图形验证码并提交登录", _step_solve_and_login),
            ("等待「邮箱安全验证」（人工完成后自动继续）", _step_wait_email_verify),
            ("处理「首次启动配置」弹窗", _step_handle_first_launch_modal),
        ]
        steps = list(login_steps)
        if CONFIG.get("AUTO_DOWNLOAD_TEMPLATE", True):
            steps += _post_login_steps()
        if exe_proc_running:
            title = "登录 + 模板下载流程（检测到旧实例未带调试端口，启动新实例）"
        else:
            title = "登录 + 模板下载流程"

    flow_begin(title, [name for name, _ in steps])

    step_index = 0
    all_ok = False
    # 记录「登录阶段」的步骤数，进入该位置之后就算进入首页自动流程
    # （以前直接拿 len(login_steps)，但 login_steps 在上面三个分支里长度不同，
    #  且 steps 会拼接 login_steps + post_steps —— 这里直接记录拼接前长度最稳）
    _pre_login_count = len(login_steps)

    try:
        while step_index < len(steps):
            _wait_until_resumed()

            name, fn = steps[step_index]
            step_index += 1
            # 刚走完登录阶段的最后一步 —— 进入首页自动流程，状态回到「运行中」
            if step_index == _pre_login_count + 1:
                set_state(RUNNING)
            res = run_step(name, fn)
            if res is FLOW_STOP:
                # 本轮没有可执行的数据：不记为失败，也不继续后面的步骤
                all_ok = True
                print("[*] 本轮无需继续执行后续步骤，流程正常结束。")
                break
            if res is False:
                all_ok = False
                break
        else:
            all_ok = True
        if all_ok:
            set_state(FINISHED)
    except Exception as e:
        set_state(ERROR)
        print("[!] 自动化过程出错: " + repr(e))
        raise

    finally:
        if CONFIG.get("KEEP_APP_OPEN"):
            # 保留应用与 CDP 连接，便于后续继续执行登录后的逻辑
            print("[*] 应用保持运行，CDP 连接已保留，可继续执行登录后的逻辑。")
        else:
            if cdp_client is not None:
                try:
                    cdp_client.close()
                except Exception:
                    pass
            cdp_client = None
            if app_proc is not None:
                try:
                    app_proc.terminate()
                except Exception:
                    pass
                app_proc = None
        flow_summary(all_ok)
        # 流程结束即删掉 JSONL 中转文件，避免几十万行数据长期占用磁盘
        _discard_rows_file()
        # 关键：释放「流程正在跑」标志，允许 on_retry / 再次启动
        _FLOW_ACTIVE = False


def _step_launch_app():
    global app_proc
    if app_proc is not None:
        # KEEP_APP_OPEN 场景下点「重新执行」时复用已在运行的实例
        print("[*] 应用进程已在运行，跳过启动。")
        set_state(RUNNING)
        return True
    if is_app_running(use_cache=False):
        # 启动前复核：调试端口有响应 → 应用已打开且带调试端口，直接复用
        print("[*] 检测到应用已经打开（调试端口就绪），跳过启动，直接复用现有实例。")
        set_state(RUNNING)
        return True
    if _is_exe_process_running():
        # 进程在跑但调试端口没响应 → 用户手动打开的旧实例，未带 --remote-debugging-port
        print("[!] 应用进程已在运行，但未开启远程调试端口。")
        print("[!] 将强制启动带调试端口的新实例，请留意是否需要手动关闭旧实例。")
    app_proc = launch_app()
    set_state(RUNNING)
    return True


def _step_connect_running_app():
    """连接已在运行的应用（不再重复启动应用）。"""
    global cdp_client
    if cdp_client is not None:
        try:
            cdp_client.close()
        except Exception:
            pass
        cdp_client = None
    target = detected_target or find_login_page(prefer_logged_in=True)
    print("[+] 已连接页面: " + str(target.get("title", "")) + " | " + str(target.get("url", "")))
    cdp_client = CDPClient(target["webSocketDebuggerUrl"])
    set_state(RUNNING)
    return True


def _step_find_login_page():
    global cdp_client
    if cdp_client is not None:
        try:
            cdp_client.close()
        except Exception:
            pass
    target = find_login_page()
    print("[+] 已连接页面: " + str(target.get("title", "")) + " | " + str(target.get("url", "")))
    cdp_client = CDPClient(target["webSocketDebuggerUrl"])
    return True


def _step_wait_login_form():
    assert cdp_client is not None, "cdp_client 未初始化"
    if not wait_for_login_form(cdp_client):
        set_state(ERROR)
        raise RuntimeError("超时未等到登录表单，退出")
    return True


def _step_fill_form():
    assert cdp_client is not None, "cdp_client 未初始化"
    cfg = json.dumps({"username": CONFIG["USERNAME"], "password": CONFIG["PASSWORD"]}, ensure_ascii=False)
    res = cdp_client.eval_js(JS_FILL % cfg)
    print("[+] 表单填充结果: " + str(res))
    if not res or not res.get("ok"):
        set_state(ERROR)
        raise RuntimeError("登录表单填充失败或未找到表单: " + str(res))
    return True


def _wait_for_captcha_ready(client, timeout=12, prev_src="", prev_nkeys=0):
    """等待验证码刷新完成，返回 {key, src, complete}。"""
    deadline = time.time() + timeout
    st = None
    while time.time() < deadline:
        st = client.eval_js(JS_CAPTCHA_STATE)
        if (st and st.get("complete") and st.get("key")
                and st.get("src") != prev_src and st.get("nkeys", 0) > prev_nkeys):
            return st
        time.sleep(0.3)
    return st


def _fetch_captcha_image(client):
    """从页面抓取验证码图片，返回 PIL Image；失败返回 None。"""
    res = client.eval_js(JS_GET_CAPTCHA_IMG)
    if not res or not res.get("ok"):
        print("[!] 验证码图片获取失败: " + str(res))
        return None
    try:
        raw = base64.b64decode(res["data"].split(",", 1)[1])
        return Image.open(io.BytesIO(raw)).convert("RGB")
    except Exception as e:
        print("[!] 验证码图片解析失败: " + repr(e))
        return None


def _verify_captcha_candidate(client, key, code):
    """调用应用自身登录接口校验候选验证码。

    返回 True 表示该候选已通过验证码校验；None 表示接口不可用（无法预校验）。
    """
    payload = {
        "username": CONFIG["USERNAME"],
        "password": CONFIG["PASSWORD"],
        "key": key,
        "code": code,
        "deviceInfo": "Unknown Device",
        "clientInfo": json.dumps({"version": "1.0.0"}),
        "rsaEncrypted": False,
    }
    try:
        res = client.eval_js(JS_VERIFY_CAPTCHA % json.dumps(payload, ensure_ascii=False),
                             await_promise=True)
    except Exception as e:
        print("[!] 验证码预校验调用失败: " + repr(e))
        return None
    if not isinstance(res, dict):
        return None
    if res.get("unsupported"):
        return None
    if res.get("error"):
        print("[!] 验证码预校验异常: " + str(res.get("error")))
        return None
    code_no = res.get("code")
    if code_no is None:
        return None
    # 服务端：code=3 表示"验证码错误"，其余 code 均说明验证码已通过
    return code_no != 3


def _solve_captcha_once():
    """执行一轮：刷新验证码 -> 抓图 -> 识别 -> 预校验 -> 填入输入框。

    成功返回填入的验证码字符串，失败返回 None。
    """
    assert cdp_client is not None, "cdp_client 未初始化"

    # 1) 刷新验证码，确保拿到与图片配对的 uuid
    before = cdp_client.eval_js(JS_CAPTCHA_STATE) or {}
    cdp_client.eval_js(JS_REFRESH_CAPTCHA)
    st = _wait_for_captcha_ready(
        cdp_client,
        prev_src=before.get("src", ""),
        prev_nkeys=before.get("nkeys", 0),
    )
    if not st or not st.get("complete") or not st.get("key"):
        print("[!] 验证码未就绪（无法获取图片或 uuid）")
        return None

    # 2) 抓图 + 识别
    img = _fetch_captcha_image(cdp_client)
    if img is None:
        return None
    cands = get_ocr().recognize(img)
    if not cands:
        print("[!] 验证码识别无结果")
        return None
    print("[*] 识别候选: " + str(cands[:5]))

    # 3) 预校验（同一张图可多次校验，服务端不区分大小写）
    confirmed = None
    if CAPTCHA_PRECHECK:
        prechecked = False
        for cand in cands:
            ok = _verify_captcha_candidate(cdp_client, st.get("key"), cand)
            if ok is None:
                break            # 接口不可用，退化为直接填写
            prechecked = True
            if ok:
                confirmed = cand
                print("[+] 验证码校验通过: " + cand)
                break
        if prechecked and confirmed is None:
            print("[!] 所有候选均未通过服务端校验，直接换一张重试")
            return None
    if confirmed is None:
        confirmed = cands[0]
        print("[*] 采用最优候选: " + confirmed)

    # 4) 填入验证码输入框
    filled = cdp_client.eval_js(JS_FILL_CAPTCHA % json.dumps(confirmed))
    if not filled or not filled.get("ok"):
        print("[!] 验证码填入失败: " + str(filled))
        return None
    print("[+] 已填入验证码: " + confirmed)
    return confirmed


def _step_solve_captcha():
    """自动识别图形验证码并填入（供单独调用时使用）。"""
    if not AUTO_CAPTCHA:
        return _step_wait_captcha()

    ocr = get_ocr()
    if not ocr.ready:
        print("[!] 验证码 OCR 不可用（请 pip install ddddocr），改为人工输入")
        return _step_wait_captcha()

    assert cdp_client is not None, "cdp_client 未初始化"
    set_state(STATE_SOLVE_CAPTCHA)

    # 注入 uuid 捕获：服务端用 uuid 把验证码图片和答案绑定
    try:
        cdp_client.eval_js(JS_PATCH_UUID)
    except Exception as e:
        print("[!] 注入 uuid 捕获失败: " + repr(e))

    for rnd in range(1, CAPTCHA_MAX_ROUNDS + 1):
        _wait_until_resumed()
        print("[*] 验证码识别第 %d/%d 轮" % (rnd, CAPTCHA_MAX_ROUNDS))
        if _solve_captcha_once() is not None:
            return True
    print("[!] 自动识别 %d 轮均未成功，降级为人工输入验证码" % CAPTCHA_MAX_ROUNDS)
    return _step_wait_captcha()


def _step_solve_and_login():
    """识别验证码 -> 填入 -> 点登录；验证码被服务端拒绝则换图重试。"""
    if not AUTO_CAPTCHA:
        _step_wait_captcha()
        return _step_click_login()

    ocr = get_ocr()
    if not ocr.ready:
        print("[!] 验证码 OCR 不可用（请 pip install ddddocr），改为人工输入")
        _step_wait_captcha()
        return _step_click_login()

    assert cdp_client is not None, "cdp_client 未初始化"
    set_state(STATE_SOLVE_CAPTCHA)
    try:
        cdp_client.eval_js(JS_PATCH_UUID)
    except Exception as e:
        print("[!] 注入 uuid 捕获失败: " + repr(e))

    for rnd in range(1, CAPTCHA_MAX_ROUNDS + 1):
        _wait_until_resumed()
        print("[*] 验证码识别第 %d/%d 轮" % (rnd, CAPTCHA_MAX_ROUNDS))

        if _solve_captcha_once() is None:
            continue

        if not AUTO_CLICK_LOGIN:
            print("[*] AUTO_CLICK_LOGIN=False，请在窗口中自行点击登录")
            return True

        time.sleep(0.3)
        res = cdp_client.eval_js(JS_CLICK_LOGIN)
        print("[+] 点击登录: " + str(res))
        if not res or not res.get("ok"):
            print("[!] 点击登录失败: " + str(res))
            return False

        time.sleep(1.5)
        msg = cdp_client.eval_js(JS_LOGIN_MESSAGE) or ""
        if msg and ("验证码" in msg) and ("邮箱" not in msg):
            print("[!] 验证码未通过（" + msg + "），换一张重试")
            continue
        print("[+] 登录请求已提交" + ("，服务端提示: " + msg if msg else ""))
        set_state(FINISHED)
        return True

    print("[!] 自动识别 %d 轮均未成功，降级为人工输入验证码" % CAPTCHA_MAX_ROUNDS)
    _step_wait_captcha()
    return _step_click_login()


def _step_wait_captcha():
    if not WAIT_CAPTCHA:
        return True

    set_state(STATE_WAIT_CAPTCHA)
    assert cdp_client is not None, "cdp_client 未初始化"
    while True:
        # 现在如果已暂停，就继续等待恢复，恢复后继续等待验证码填写。
        if current_state == PAUSED:
            while current_state == PAUSED:
                time.sleep(0.2)
            continue

        deadline = time.time() + CAPTCHA_TIMEOUT
        while time.time() < deadline:
            if current_state == PAUSED:
                continue
            if cdp_client.eval_js(JS_CAPTCHA_FILLED):
                print("[+] 检测到验证码已填写")
                return True
            time.sleep(0.5)

        print("[!] 等待验证码超时")
        set_state(ERROR)
        raise RuntimeError("等待验证码超时")


def _step_click_login():
    if AUTO_CLICK_LOGIN:
        assert cdp_client is not None, "cdp_client 未初始化"
        time.sleep(0.3)
        res = cdp_client.eval_js(JS_CLICK_LOGIN)
        print("[+] 点击登录: " + str(res))
        if not res or not res.get("ok"):
            set_state(ERROR)
            raise RuntimeError("点击登录失败: " + str(res))
    else:
        print("[*] AUTO_CLICK_LOGIN=False，请在窗口中自行点击登录")

    if current_state == PAUSED:
        set_state(PAUSED)
        return True

    set_state(FINISHED)
    return True


def _step_wait_email_verify():
    """提交登录后若出现「邮箱安全验证」页，转入等待状态而不是报错退出。

    处理方式：
        1. 提交登录后在短窗口（EMAIL_VERIFY_DETECT_TIMEOUT）内探测验证页；
           全程没出现 -> 说明本次登录无需邮箱验证，直接返回，保持原有正常路径；
        2. 检测到验证页 -> 置为 STATE_WAIT_EMAIL_VERIFY，自动化暂停一切点击/输入，
           只做只读 eval_js 轮询，让用户在同一个窗口里手工完成邮箱验证；
        3. 轮询到验证完成（验证页特征消失 + 出现首页/登录页标志）-> 自动继续
           执行后续上传步骤，不需要用户再点「继续」；
        4. 超过 EMAIL_VERIFY_TIMEOUT（默认 10 分钟）仍未完成 -> 置为出错，
           日志给出可操作提示，可点「重新执行」重试，不会卡死。

    本函数运行在 run_login_flow 的后台线程里，等待期间只 sleep 后台线程，
    不会阻塞 tkinter 主线程；界面提示由主线程 _update_ui 轮询刷新。
    """
    global cdp_client
    assert cdp_client is not None, "cdp_client 未初始化"

    # 上一步（点登录）可能把状态置为 FINISHED，这里先回到「运行中」避免状态闪烁
    set_state(RUNNING)

    # ---------- 1) 探测验证页是否出现 ----------
    st = None
    verify_ev = 0
    is_verify = False
    detect_deadline = time.time() + EMAIL_VERIFY_DETECT_TIMEOUT
    while time.time() < detect_deadline:
        try:
            st = cdp_client.eval_js(JS_CHECK_EMAIL_VERIFY)
        except Exception as e:
            print("[!] 探测邮箱验证页失败: " + repr(e))
            return True    # 页面异常交由后续步骤按原逻辑处理
        verify_ev, is_verify = _email_verify_evidence(st)
        if is_verify:
            break
        # 已进入首页（登录成功且无需邮箱验证），提前结束探测；
        # 仍停在登录表单时先不判定，验证页可能稍后才跳转出来，继续探测到窗口结束。
        if isinstance(st, dict) and st.get("hasTemplateBtn"):
            break
        time.sleep(EMAIL_VERIFY_POLL_INTERVAL)

    if not is_verify:
        print("[*] 未出现「邮箱安全验证」页面，继续执行后续步骤。")
        return True

    # ---------- 2) 检测到验证页 -> 转入等待（不报错、不退出） ----------
    t0 = time.time()
    EMAIL_VERIFY_WAIT["t0"] = t0
    EMAIL_VERIFY_WAIT["active"] = True
    set_state(STATE_WAIT_EMAIL_VERIFY)
    st = st if isinstance(st, dict) else {}
    print("[!] 检测到「邮箱安全验证」页面（证据 %d 条）: %s"
          % (verify_ev, str(st.get("title") or st.get("url") or "")))
    print("[!] 请在应用窗口中手工完成邮箱验证（获取验证码 -> 输入验证码 -> 验证并登录）。")
    print("[!] 自动化已暂停点击/输入，验证完成后会自动继续，无需再点「继续」。"
          "（最长等待 %d 秒）" % EMAIL_VERIFY_TIMEOUT)

    # ---------- 3) 轮询等待验证完成（全程只读探测） ----------
    last_report = t0
    try:
        while True:
            # 用户点了「暂停」：保持只读等待，绝不与用户抢页面
            if current_state == PAUSED:
                time.sleep(0.2)
                continue
            # 暂停后点了「继续」：恢复为「等待邮箱验证」状态，界面提示保持可见
            if current_state == RUNNING:
                set_state(STATE_WAIT_EMAIL_VERIFY)

            now = time.time()
            # 4) 超时保护：给出可操作提示并退出等待，不卡死线程
            if now - t0 >= EMAIL_VERIFY_TIMEOUT:
                print("[!] 等待邮箱验证超时（已等待 %d 秒，上限 %d 秒）。"
                      % (int(now - t0), EMAIL_VERIFY_TIMEOUT))
                print("[!] 请在窗口中完成邮箱验证后点「重新执行」重试；"
                      "如经常超时，可调大脚本配置区的 EMAIL_VERIFY_TIMEOUT。")
                set_state(ERROR)
                raise RuntimeError("等待邮箱安全验证超时")

            try:
                st = cdp_client.eval_js(JS_CHECK_EMAIL_VERIFY)
            except Exception as e:
                print("[!] 轮询邮箱验证页状态失败: " + repr(e))
                time.sleep(EMAIL_VERIFY_POLL_INTERVAL)
                continue

            verify_ev, is_verify = _email_verify_evidence(st)
            if not is_verify:
                # 验证页特征已消失 -> 按去向判断是否验证完成
                st = st if isinstance(st, dict) else {}
                if st.get("hasTemplateBtn"):
                    print("[+] 邮箱验证已完成，已进入首页（检测到「下载表格模板」）。")
                    break
                if st.get("hasLoginForm"):
                    print("[!] 邮箱验证页已消失但回到登录表单（可能验证未通过或被取消），"
                          "继续执行后续步骤。")
                    break
                if st.get("ready") == "complete":
                    print("[+] 邮箱验证页已离开（URL: %s），视为验证完成。"
                          % str(st.get("url", "")))
                    break

            # 定期回报等待进度，日志能看到「还在等 / 等了多久」
            if now - last_report >= EMAIL_VERIFY_PROGRESS_LOG:
                last_report = now
                print("[*] 仍在等待邮箱验证 ... 已等待 %d 秒（上限 %d 秒）。"
                      % (int(now - t0), EMAIL_VERIFY_TIMEOUT))
            time.sleep(EMAIL_VERIFY_POLL_INTERVAL)
    finally:
        EMAIL_VERIFY_WAIT["active"] = False

    print("[+] 邮箱验证等待结束，共等待 %d 秒，恢复自动执行后续步骤。"
          % int(time.time() - t0))
    set_state(RUNNING)
    return True


# 在「首次启动配置」弹窗里定位两个按钮：
#   btnPick —— 「选择文件夹」（点它会拉起 Windows 原生目录选择框）
#   btnDone —— 「完成配置」（选好目录之后才可用）
# 只在可见弹窗容器内找，且优先精确匹配文案，避免命中提示文字里的片段。
JS_FIRST_LAUNCH_BUTTONS = r"""
(function () {
    const MODALS = ['.arco-modal', '.arco-modal-wrapper', '[role="dialog"]'];
    const PICK = '选择文件夹';
    const DONE = '完成配置';
    function norm(s) { return ((s || '') + '').replace(/\s+/g, ''); }
    function visible(el) {
        const r = el.getBoundingClientRect();
        if (r.width <= 0 || r.height <= 0) return false;
        const st = window.getComputedStyle(el);
        return st.display !== 'none' && st.visibility !== 'hidden';
    }
    function info(el) {
        const r = el.getBoundingClientRect();
        return {
            x: Math.round(r.x + r.width / 2),
            y: Math.round(r.y + r.height / 2),
            tag: el.tagName,
            cls: (el.className || '') + '',
            text: norm(el.innerText || el.textContent || ''),
            disabled: !!(el.disabled || el.getAttribute('aria-disabled') === 'true')
        };
    }
    function pick(scope, name) {
        const els = Array.from(scope.querySelectorAll(
            'button, a, [role="button"], .arco-btn')).filter(visible);
        for (const el of els) {              // 先精确匹配
            if (norm(el.innerText || el.textContent || '') === name) return info(el);
        }
        for (const el of els) {              // 再退到包含匹配
            const t = norm(el.innerText || el.textContent || '');
            if (t.indexOf(name) >= 0 && t.length <= name.length + 6) return info(el);
        }
        return null;
    }
    for (const sel of MODALS) {
        for (const scope of Array.from(document.querySelectorAll(sel)).filter(visible)) {
            const txt = scope.innerText || '';
            if (txt.indexOf('首次启动配置') < 0 && txt.indexOf('备份路径') < 0) continue;
            const btnPick = pick(scope, PICK);
            const btnDone = pick(scope, DONE);
            if (btnPick || btnDone) {
                return { ok: true, btnPick: btnPick, btnDone: btnDone };
            }
        }
    }
    return { ok: false };
})()
"""


def _first_launch_backup_dir() -> str:
    """「首次启动配置」要选择的备份根目录（默认 %ProgramData%\\AreportFile）。

    可在 gui_config.json 里用同名键 FIRST_LAUNCH_BACKUP_DIR 覆盖，无需改代码。
    """
    return _optional_str_config("FIRST_LAUNCH_BACKUP_DIR",
                                DEFAULT_FIRST_LAUNCH_BACKUP_DIR)


def _first_launch_find_buttons():
    """在「首次启动配置」弹窗里找「选择文件夹」/「完成配置」两个按钮。"""
    try:
        return cdp_client.eval_js(JS_FIRST_LAUNCH_BUTTONS)
    except Exception as e:
        print("[!] 定位「首次启动配置」弹窗按钮异常: " + repr(e))
        return None


def _first_launch_click_button(kind, timeout=15):
    """等待并点击弹窗里的按钮。

    kind: 'btnPick'（选择文件夹） / 'btnDone'（完成配置）
    返回点击到的按钮信息（dict）；超时或按钮不可用返回 None。
    """
    deadline = time.time() + timeout
    last_tip = ""
    while time.time() < deadline:
        res = _first_launch_find_buttons()
        if res and res.get("ok"):
            btn = res.get(kind)
            if btn and not btn.get("disabled"):
                print("[*] 点击弹窗里的「%s」（%s，坐标 %d,%d） ..."
                      % (btn.get("text"), str(btn.get("tag")),
                         int(btn.get("x", 0)), int(btn.get("y", 0))))
                _cdp_click_at(cdp_client, int(btn.get("x", 0)), int(btn.get("y", 0)))
                return btn
            if btn and btn.get("disabled"):
                tip = "弹窗按钮「%s」当前不可用，继续等待 ..." % btn.get("text")
                if tip != last_tip:            # 只在状态变化时打印，避免刷屏
                    print("[*] " + tip)
                    last_tip = tip
        time.sleep(0.5)
    return None


def _first_launch_manual_fallback():
    """自动配置失败时的兜底：置为 PAUSED，等人工在应用窗口配好后点「继续」。"""
    print("[!] 改为人工处理：请切换到应用窗口，点「选择文件夹」指定备份路径，再点「完成配置」；")
    print("[!] 完成后回到本助手点击「继续」，自动化将继续执行。")
    set_state(PAUSED)

    # 1) 等待用户在 GUI 中点「继续」
    while current_state == PAUSED:
        time.sleep(0.2)

    # 2) 用户点继续后，轮询等待弹窗真正关闭
    deadline = time.time() + 180
    while time.time() < deadline:
        if current_state == PAUSED:
            while current_state == PAUSED:
                time.sleep(0.2)
        try:
            st = cdp_client.eval_js(JS_FIRST_LAUNCH_MODAL)
        except Exception as e:
            print("[!] 检测弹窗状态失败: " + repr(e))
            break
        if not st or not st.get("found"):
            print("[+] 「首次启动配置」弹窗已关闭，配置完成。")
            set_state(FINISHED)
            return True
        time.sleep(1)

    print("[!] 等待「首次启动配置」弹窗关闭超时（仍视为已完成）。")
    set_state(FINISHED)
    return True


def _step_handle_first_launch_modal():
    """登录后若弹出「首次启动配置」弹窗，自动完成备份路径配置。

    自动流程：
        1) 先创建备份根目录（默认 %ProgramData%\\AreportFile；原生目录框不接受不存在的路径）；
        2) 点弹窗里的「选择文件夹」→ 拉起 Windows 原生目录选择框 → 写入该目录并确认；
        3) 回到页面点「完成配置」，弹窗关闭即视为配置完成。

    任何一步失败（按钮改名、弹的不是原生目录框、权限不足等）都自动回退到
    「暂停 + 人工处理」，不中断整个流程。
    """
    assert cdp_client is not None, "cdp_client 未初始化"

    # 上一步已置为 FINISHED，此处回到 RUNNING，避免出现「已完成」与「已暂停」之间的状态闪烁
    set_state(RUNNING)

    # 给首页路由与弹窗渲染留一点时间
    time.sleep(2)
    res = cdp_client.eval_js(JS_FIRST_LAUNCH_MODAL)
    if not res or not res.get("found"):
        print("[*] 未检测到「首次启动配置」弹窗，直接进入首页。")
        return True

    print("[!] 检测到「首次启动配置」弹窗，开始自动配置备份路径 ...")

    # 1) 先把备份根目录建好（原生目录选择框里输入不存在的路径会被拒绝）
    target = _first_launch_backup_dir()
    try:
        os.makedirs(target, exist_ok=True)
        print("[+] 已确保备份根目录存在: " + target)
    except OSError as e:
        print("[!] 创建备份根目录失败: " + repr(e))
        return _first_launch_manual_fallback()

    # 2) 点「选择文件夹」——它会拉起 Windows 原生目录选择框
    if not _first_launch_click_button("btnPick", FIRST_LAUNCH_PICK_TIMEOUT):
        print("[!] 未找到可点击的「选择文件夹」按钮。")
        return _first_launch_manual_fallback()

    # 3) 等原生目录选择框弹出并选中目标目录
    #    注意：原生模态框会阻塞渲染进程，这一段只能用 Win32 侧操作，不能再 eval_js
    dlg, title, btn_text = 0, "", ""
    deadline = time.time() + FIRST_LAUNCH_DIALOG_TIMEOUT
    while time.time() < deadline:
        dlg, title, btn_text = _win_find_folder_dialog()
        if dlg:
            break
        time.sleep(0.3)
    if not dlg:
        print("[!] 等待原生「选择文件夹」对话框超时（该弹窗可能不是原生目录框）。")
        return _first_launch_manual_fallback()
    print("[+] 已定位原生目录选择对话框：标题=" + title + "，按钮=" + btn_text)

    if not _win_folder_dialog_pick(dlg, target):
        print("[!] 在原生对话框中选中备份根目录失败。")
        _win_dismiss_error_boxes()
        return _first_launch_manual_fallback()

    # 4) 回到页面点「完成配置」
    time.sleep(0.5)
    if not _first_launch_click_button("btnDone", FIRST_LAUNCH_DONE_TIMEOUT):
        print("[!] 未找到可点击的「完成配置」按钮。")
        return _first_launch_manual_fallback()

    # 5) 等弹窗真正关闭
    deadline = time.time() + FIRST_LAUNCH_DONE_TIMEOUT
    while time.time() < deadline:
        time.sleep(0.5)
        try:
            st = cdp_client.eval_js(JS_FIRST_LAUNCH_MODAL)
        except Exception as e:
            print("[!] 检测弹窗状态失败: " + repr(e))
            break
        if not st or not st.get("found"):
            print("[+] 「首次启动配置」弹窗已关闭，备份路径自动配置完成。")
            set_state(FINISHED)
            return True

    print("[!] 点击「完成配置」后弹窗仍未关闭，回退人工处理。")
    return _first_launch_manual_fallback()


# ===================== 原生「保存」对话框自动化 =====================
# 说明：Electron 弹出的保存对话框是 Windows 原生窗口（class "#32770"），
#       CDP 协议操作不到它，这里直接用 user32 找到窗口、读取文件名并点击「保存」。
_user32: Any = None
_WNDPROC: Any = None
if _WIN32:
    try:
        _user32 = ctypes.WinDLL("user32", use_last_error=True)
        _WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
        _user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
        _user32.GetWindowTextW.restype = ctypes.c_int
        _user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
        _user32.GetClassNameW.restype = ctypes.c_int
        _user32.EnumWindows.argtypes = [_WNDPROC, ctypes.c_long]
        _user32.EnumChildWindows.argtypes = [wintypes.HWND, _WNDPROC, ctypes.c_long]
        _user32.GetWindow.argtypes = [wintypes.HWND, ctypes.c_uint]
        _user32.GetWindow.restype = wintypes.HWND
        _user32.IsWindowVisible.argtypes = [wintypes.HWND]
        _user32.IsWindowVisible.restype = wintypes.BOOL
        _user32.IsWindow.argtypes = [wintypes.HWND]
        _user32.IsWindow.restype = wintypes.BOOL
        _user32.GetParent.argtypes = [wintypes.HWND]
        _user32.GetParent.restype = wintypes.HWND
        _user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.c_void_p]
        _user32.GetWindowRect.restype = wintypes.BOOL
        _user32.GetDlgCtrlID.argtypes = [wintypes.HWND]
        _user32.GetDlgCtrlID.restype = ctypes.c_int
        _user32.PostMessageW.argtypes = [wintypes.HWND, ctypes.c_uint, ctypes.c_ulong,
                                         ctypes.c_longlong]
        _user32.PostMessageW.restype = wintypes.BOOL
        _user32.SendMessageW.argtypes = [wintypes.HWND, ctypes.c_uint, ctypes.c_ulong,
                                         ctypes.c_longlong]
        _user32.SendMessageW.restype = ctypes.c_longlong
        _user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND,
                                                     ctypes.POINTER(ctypes.c_ulong)]
        _user32.GetWindowThreadProcessId.restype = ctypes.c_ulong
    except Exception as _win_err:  # pragma: no cover
        _user32 = None
        print("[!] Win32 API 初始化失败: " + repr(_win_err))
else:  # pragma: no cover
    _user32 = None

_kernel32: Any = None
if _WIN32:
    try:
        _kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        _kernel32.OpenProcess.argtypes = [ctypes.c_uint, ctypes.c_int, ctypes.c_ulong]
        _kernel32.OpenProcess.restype = wintypes.HANDLE
        _kernel32.QueryFullProcessImageNameW.argtypes = [
            wintypes.HANDLE, ctypes.c_uint, wintypes.LPWSTR,
            ctypes.POINTER(ctypes.c_ulong)]
        _kernel32.QueryFullProcessImageNameW.restype = ctypes.c_int
        _kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
        _kernel32.CloseHandle.restype = ctypes.c_int
    except Exception as _kern_err:  # pragma: no cover
        _kernel32 = None
        print("[!] kernel32 初始化失败: " + repr(_kern_err))

_GW_CHILD = 5
_GW_HWNDNEXT = 2
_BM_CLICK = 0x00F5


def _win_class(hwnd):
    """读取窗口类名。"""
    if not _user32:
        return ""
    buf = ctypes.create_unicode_buffer(256)
    try:
        _user32.GetClassNameW(wintypes.HWND(hwnd), buf, 256)
    except Exception:
        return ""
    return buf.value


def _win_text(hwnd):
    """读取窗口/控件文本。"""
    if not _user32:
        return ""
    buf = ctypes.create_unicode_buffer(1024)
    try:
        _user32.GetWindowTextW(wintypes.HWND(hwnd), buf, 1024)
    except Exception:
        return ""
    return buf.value


def _win_top_windows():
    """枚举所有可见的顶层窗口，返回 [(hwnd, class, title)]。"""
    out = []
    if not _user32:
        return out

    def _cb(hwnd, _lparam):
        try:
            if not _user32.IsWindowVisible(wintypes.HWND(hwnd)):
                return True
            out.append((int(hwnd), _win_class(hwnd), _win_text(hwnd)))
        except Exception:
            pass
        return True

    try:
        _user32.EnumWindows(_WNDPROC(_cb), 0)
    except Exception as e:
        print("[!] 枚举顶层窗口失败: " + repr(e))
    return out


def _win_direct_children(hwnd):
    """枚举直接子窗口（用 GW_CHILD / GW_HWNDNEXT 遍历，避免跨进程枚举的歧义）。"""
    out = []
    if not _user32:
        return out
    try:
        child = _user32.GetWindow(wintypes.HWND(hwnd), _GW_CHILD)
        while child:
            h = int(child)
            out.append((h, _win_class(h), _win_text(h)))
            child = _user32.GetWindow(wintypes.HWND(h), _GW_HWNDNEXT)
    except Exception:
        pass
    return out


def _win_all_descendants(hwnd):
    """枚举全部后代窗口（含孙窗口）。"""
    out = []
    if not _user32:
        return out

    def _cb(h, _lparam):
        try:
            out.append((int(h), _win_class(int(h)), _win_text(int(h))))
        except Exception:
            pass
        return True

    try:
        _user32.EnumChildWindows(wintypes.HWND(hwnd), _WNDPROC(_cb), 0)
    except Exception:
        pass
    return out


def _win_find_button(hwnd, prefixes):
    """在窗口内查找文本以 prefixes 中任一项开头的按钮，返回 (hwnd, text)。"""
    for h, cls, txt in _win_all_descendants(hwnd):
        if cls != "Button":
            continue
        norm = (txt or "").replace("&", "").strip()
        for p in prefixes:
            if norm.startswith(p):
                return h, txt
    return None, ""


def _win_window_pid(hwnd):
    """取窗口所属进程 PID。"""
    if not _user32:
        return 0
    pid = ctypes.c_ulong(0)
    try:
        _user32.GetWindowThreadProcessId(wintypes.HWND(hwnd), ctypes.byref(pid))
    except Exception:
        return 0
    return int(pid.value)


def _win_process_exe(pid):
    """取进程的可执行文件路径。"""
    if not pid or not _kernel32:
        return ""
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    try:
        h = _kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not h:
            return ""
        try:
            size = ctypes.c_ulong(260)
            buf = ctypes.create_unicode_buffer(260)
            if _kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
                return buf.value
            return ""
        finally:
            _kernel32.CloseHandle(h)
    except Exception:
        return ""


# ---------------------------------------------------------------------
# 进程级强制关闭辅助（窗口关闭失败时的兜底，避免文件锁长期残留）
# ---------------------------------------------------------------------
_PROCESS_TERMINATE = 0x0001
_PROCESS_QUERY_INFORMATION = 0x0400


def _pid_alive(pid):
    """判断 PID 是否仍在运行（GetExitCodeProcess → STILL_ACTIVE 判断）。"""
    if not pid or not _kernel32:
        return False
    STILL_ACTIVE = 259
    try:
        h = _kernel32.OpenProcess(_PROCESS_QUERY_INFORMATION, False, pid)
        if not h:
            return False
        try:
            exit_code = ctypes.c_ulong(0)
            if _kernel32.GetExitCodeProcess(h, ctypes.byref(exit_code)):
                return exit_code.value == STILL_ACTIVE
            return False
        finally:
            _kernel32.CloseHandle(h)
    except Exception:
        return False


def _terminate_pid(pid, wait_timeout=5, force_kill_timeout=2):
    """严格四步关闭：terminate → wait → kill → wait。返回是否已退出。

    顺序：
      1) 先给进程一个体面终止（TerminateProcess 传 0，等价于 WM_CLOSE 的
         强制版，但比 taskkill /F 温和）；
      2) 轮询 pid_alive，最多 wait_timeout 秒；
      3) 还没死 → TerminateProcess(exit_code=1) 强制杀死；
      4) 再等 force_kill_timeout 秒确认真的退出。
    """
    if not pid or not _kernel32:
        return False
    try:
        h = _kernel32.OpenProcess(_PROCESS_TERMINATE | _PROCESS_QUERY_INFORMATION, False, pid)
        if not h:
            return False
        try:
            # step 1: 体面终止
            _kernel32.TerminateProcess(h, 0)
            deadline = time.time() + wait_timeout
            while time.time() < deadline:
                if not _pid_alive(pid):
                    return True
                time.sleep(0.3)

            # step 3: 强制杀死
            _kernel32.TerminateProcess(h, 1)
            deadline = time.time() + force_kill_timeout
            while time.time() < deadline:
                if not _pid_alive(pid):
                    return True
                time.sleep(0.3)
            return not _pid_alive(pid)
        finally:
            _kernel32.CloseHandle(h)
    except Exception as e:
        print("[!] terminate_pid(%d) 异常: %s" % (pid, repr(e)))
        return False


def _win_process_cmdline(pid):
    """取进程的完整命令行（WMI 方式，Win32_Process.CommandLine）。"""
    try:
        import wmi  # 可选依赖；没装就退化
    except ImportError:
        return ""
    try:
        c = wmi.WMI()
        for p in c.Win32_Process(ProcessId=int(pid)):
            return p.CommandLine or ""
    except Exception:
        return ""


def _enum_pid_top_windows(pid):
    """枚举指定 PID 的所有顶层窗口，返回 [(hwnd, title, exe_lower), ...]。"""
    if not pid or not _user32:
        return []
    results = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    def _cb(hwnd, _lparam):
        if not _user32.IsWindowVisible(hwnd):
            return True
        wpid = _win_window_pid(hwnd)
        if wpid != pid:
            return True
        try:
            length = _user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                _user32.GetWindowTextW(hwnd, buf, length + 1)
                title = buf.value
            else:
                title = ""
        except Exception:
            title = ""
        results.append((hwnd, title, _win_process_exe_lower(hwnd)))
        return True

    _user32.EnumWindows(_cb, 0)
    return results


def _kill_office_procs_for_file(target_path, wait_timeout=15):
    """兜底：仅关闭持有目标文件的窗口/进程，绝不动同进程的其它文档。

    顺序（温和 → 强制）：
      1) 收集所有 office PID，其顶层窗口标题里包含目标文件名 → 只对这些
         匹配窗口发 WM_CLOSE（PID 里其它文档窗口不碰）
      2) 等待文件锁释放（期间持续 dismiss 可能出现的保存弹窗）
      3) 仍被锁 → 再看一轮：
           - 匹配窗口还在？→ 继续发 WM_CLOSE + 等
           - 匹配窗口全没了，但 PID 还有其它 office 窗口？→ **不 terminate**
             （那是用户的其他文档，让他自己关；文件锁可能是 Excel 内部还在清理）
           - 匹配窗口全没了 + PID 也没 office 窗口了（进程正在退出或僵尸）
             且进程还活着？→ terminate（最后兜底）

    返回 True 表示文件锁已释放；False 表示仍被占用。
    """
    if not _WIN32 or not _user32:
        return True
    if not target_path:
        return True

    base = os.path.splitext(os.path.basename(target_path))[0].lower()
    # 收集：pid → [该 pid 下所有窗口标题含 base 的 hwnd]
    pid_to_matched_hwnds = {}
    for hwnd, _cls, title in _win_top_windows():
        if not title:
            continue
        if base not in title.lower():
            continue
        exe = _win_process_exe_lower(hwnd)
        if not any(k in exe for k in _OFFICE_EXE_KEYS):
            continue
        pid = _win_window_pid(hwnd)
        if pid:
            pid_to_matched_hwnds.setdefault(pid, []).append(hwnd)

    if not pid_to_matched_hwnds:
        # 没找到任何匹配窗口 → 可能被后台/隐藏窗口持有，直接走等待 + terminate 兜底
        pass
    else:
        for pid, hwnds in pid_to_matched_hwnds.items():
            exe_name = os.path.basename(_win_process_exe_lower(hwnds[0]) or "").strip() or "office进程"
            print("[*] 兜底：向 %s (PID=%d) 的 %d 个目标文件窗口发 WM_CLOSE ..."
                  % (exe_name, pid, len(hwnds)))
            for hwnd in hwnds:
                _win_close_window(hwnd)

    # 轮询等待：dismiss 弹窗 + 查文件锁
    deadline = time.time() + wait_timeout
    while time.time() < deadline:
        _win_dismiss_prompts()
        if not _win_file_locked(target_path):
            return True
        time.sleep(0.5)

    # 还是被锁 → 检查每个 PID 当前状态再决定要不要 terminate
    if pid_to_matched_hwnds:
        for pid in list(pid_to_matched_hwnds.keys()):
            if not _pid_alive(pid):
                continue
            matched_now = _enum_pid_top_windows(pid)
            # 还存在 office 窗口（不管是不是我们的） → 可能只是 Excel 内部清理慢，
            # 再等一会儿
            if any(any(k in exe for k in _OFFICE_EXE_KEYS) for _h, _t, exe in matched_now):
                print("[*] PID=%d 仍有 office 窗口，暂不强制 terminate 以保护其它文档，"
                      "继续等待文件锁释放 ..." % pid)
                # 再等一小段
                sub_deadline = time.time() + 5
                while time.time() < sub_deadline:
                    _win_dismiss_prompts()
                    if not _win_file_locked(target_path):
                        return True
                    time.sleep(0.5)
                continue

            # pid 已无 office 顶层窗口 → 僵尸/正在退出 → 可以 terminate
            exe_name = "office进程"
            print("[!] PID=%d 已无 office 顶层窗口但进程仍活着 → terminate 兜底 ..."
                  % pid)
            _terminate_pid(pid, wait_timeout=3, force_kill_timeout=2)

        # terminate 后再确认锁
        deadline2 = time.time() + 5
        while time.time() < deadline2:
            if not _win_file_locked(target_path):
                return True
            time.sleep(0.3)

    return not _win_file_locked(target_path)


def _win_find_save_dialog():
    """查找系统「保存」对话框，返回 (hwnd, title, button_text)。

    优先返回属于「上传系统」进程的对话框，避免误点其它软件的保存框。
    """
    want_exe = os.path.normcase(CONFIG.get("EXE_PATH") or "")
    fallback = (None, "", "")
    for hwnd, cls, title in _win_top_windows():
        if cls != "#32770":
            continue
        btn, txt = _win_find_button(hwnd, ("保存", "Save"))
        if not btn:
            continue
        if want_exe:
            proc_exe = _win_process_exe(_win_window_pid(hwnd))
            if proc_exe and os.path.normcase(proc_exe) == want_exe:
                return hwnd, title, txt
        if fallback[0] is None:
            fallback = (hwnd, title, txt)
    return fallback


def _win_find_confirm_dialog():
    """查找「文件已存在，是否替换」之类确认框（带「是/替换」按钮）。

    优先返回标题里带「确认 / 另存为 / Confirm」的窗口，避免误点无关对话框。
    """
    fallback = (None, "", "")
    for hwnd, cls, title in _win_top_windows():
        if cls != "#32770":
            continue
        btn, txt = _win_find_button(hwnd, ("是", "替换", "Yes"))
        if not btn:
            continue
        if any(k in title for k in ("确认", "另存为", "Confirm", "替换")):
            return hwnd, title, txt
        if fallback[0] is None:
            fallback = (hwnd, title, txt)
    return fallback


def _win_edit_controls(hwnd):
    """窗口内所有 Edit 控件：[(hwnd, text)]。"""
    return [(h, (txt or "").strip())
            for h, cls, txt in _win_all_descendants(hwnd) if cls == "Edit"]


def _win_combo_controls(hwnd):
    """窗口内所有 ComboBox / ComboBoxEx32 控件。"""
    return [(h, cls, (txt or "").strip())
            for h, cls, txt in _win_all_descendants(hwnd)
            if cls in ("ComboBox", "ComboBoxEx32")]


def _win_looks_like_filename(s):
    """文本是否像一个文件名（而不是「Excel文件 (*.xlsx)」这类类型下拉）。"""
    if not s:
        return False
    s = s.strip()
    if s.lower().endswith((".xlsx", ".xls", ".xlsm", ".csv")):
        return True
    if "\\" in s or ":" in s:
        return False
    return bool(re.search(r"\.[A-Za-z0-9]{1,5}$", s))


def _win_get_filename_edit(dlg_hwnd):
    """读取「文件名(N)」输入框里的内容，返回 (文件名, 控件句柄)。"""
    if not dlg_hwnd:
        return "", None

    edits = _win_edit_controls(dlg_hwnd)
    # 1) 内容像文件名的 Edit（现代对话框：ComboBox → Edit）
    for h, txt in edits:
        if _win_looks_like_filename(txt):
            return txt, h
    # 2) ComboBox 自身文本像文件名
    for h, cls, txt in _win_combo_controls(dlg_hwnd):
        if _win_looks_like_filename(txt):
            return txt, h
    # 3) 兜底：任意非空 Edit
    for h, txt in edits:
        if txt:
            return txt, h
    return "", None


def _win_get_current_folder(dlg_hwnd):
    """从地址栏（面包屑）读取当前所在目录，例如 C:\\ProgramData\\AreportFile。"""
    if not dlg_hwnd:
        return ""
    for h, cls, txt in _win_all_descendants(dlg_hwnd):
        if cls != "ToolbarWindow32":
            continue
        t = (txt or "").strip()
        if not t or ":" not in t:
            continue
        head, _, tail = t.partition(":")
        if len(head) > 12:          # 排除「地址区段」这类非地址栏控件
            continue
        path = tail.strip()
        if re.match(r"^[A-Za-z]:[\\/]", path) or path.startswith("\\\\"):
            return path
    return ""


def _win_click_button(hwnd):
    """点击按钮（BM_CLICK 跨进程可用）。"""
    if not _user32 or not hwnd:
        return False
    try:
        _user32.PostMessageW(wintypes.HWND(hwnd), _BM_CLICK, 0, 0)
        return True
    except Exception as e:
        print("[!] 点击按钮失败: " + repr(e))
        return False


# ===================== 跨进程窗口文本读写 =====================
# 说明：GetWindowTextW 对其它进程的控件取不到内容，必须用 SendMessage 发
#       WM_GETTEXT / WM_SETTEXT（系统会自动跨进程封送这两个消息）。
_WM_CLOSE = 0x0010
_WM_GETTEXT = 0x000D
_WM_SETTEXT = 0x000C


def _win_get_text_msg(hwnd, max_len=1024):
    """用 WM_GETTEXT 读取控件文本（可跨进程）。"""
    if not _user32 or not hwnd:
        return ""
    buf = ctypes.create_unicode_buffer(max_len)
    try:
        _user32.SendMessageW(wintypes.HWND(hwnd), _WM_GETTEXT, max_len,
                              ctypes.c_longlong(ctypes.cast(buf, ctypes.c_void_p).value or 0))
        return buf.value
    except Exception as e:
        print("[!] 读取控件文本失败: " + repr(e))
        return ""


def _win_set_text(hwnd, text):
    """用 WM_SETTEXT 设置控件文本（可跨进程）。"""
    if not _user32 or not hwnd:
        return False
    buf = ctypes.create_unicode_buffer(text)
    try:
        _user32.SendMessageW(wintypes.HWND(hwnd), _WM_SETTEXT, 0,
                              ctypes.c_longlong(ctypes.cast(buf, ctypes.c_void_p).value or 0))
        return True
    except Exception as e:
        print("[!] 设置控件文本失败: " + repr(e))
        return False


def _win_notify_edit_changed(hwnd):
    """通知父窗口：Edit 内容已变化（EN_CHANGE）。

    ⚠ 注意：给文件对话框的「文件名」框补发 EN_CHANGE 是**有害**的——
    实测外壳会把路径规范化成正斜杠 "C:/ProgramData/..."，点「打开」时报
    「文件名无效」，对话框关不掉。因此本模块只在必要时使用该函数，
    选择文件的流程里不调用它。
    """
    if not _user32 or not hwnd:
        return False
    try:
        parent = int(_user32.GetParent(wintypes.HWND(hwnd)) or 0)
        if not parent:
            return False
        cid = int(_user32.GetDlgCtrlID(wintypes.HWND(hwnd)) or 0)
        wparam = (cid & 0xFFFF) | (0x0300 << 16)     # EN_CHANGE
        _user32.PostMessageW(wintypes.HWND(parent), 0x0111, wparam, hwnd)
        return True
    except Exception as e:
        print("[!] 通知文本变化失败: " + repr(e))
        return False


def _win_dismiss_error_boxes():
    """关闭残留的错误提示框（例如「文件名无效」这类模态错误框）。

    错误框也是 #32770，但没有「打开」按钮、只有「确定 / OK」，
    会挡住文件选择框，导致后续点击全部无效。
    """
    closed = 0
    for hwnd, cls, title in _win_top_windows():
        if cls != "#32770":
            continue
        if _win_find_button(hwnd, ("打开", "Open"))[0]:
            continue          # 这是文件选择框本身
        btn, _txt = _win_find_button(hwnd, ("确定", "OK"))
        if not btn:
            continue
        print("[*] 关闭残留的错误提示框: " + title)
        _win_click_button(btn)
        closed += 1
        time.sleep(0.3)
    return closed


def _win_window_exists(hwnd):
    """窗口是否仍然存在。"""
    if not _user32 or not hwnd:
        return False
    try:
        return bool(_user32.IsWindow(wintypes.HWND(hwnd)))
    except Exception:
        return False


class _WinRect(ctypes.Structure):
    _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG),
                ("right", wintypes.LONG), ("bottom", wintypes.LONG)]


def _win_get_rect(hwnd):
    """取窗口在屏幕上的位置 (left, top, right, bottom)。"""
    if not _user32 or not hwnd:
        return (0, 0, 0, 0)
    r = _WinRect()
    try:
        _user32.GetWindowRect(wintypes.HWND(hwnd), ctypes.byref(r))
        return (int(r.left), int(r.top), int(r.right), int(r.bottom))
    except Exception:
        return (0, 0, 0, 0)


def _win_close_window(hwnd):
    """向窗口发送 WM_CLOSE（跨进程可用）。"""
    if not _user32 or not hwnd:
        return False
    try:
        _user32.PostMessageW(wintypes.HWND(hwnd), _WM_CLOSE, 0, 0)
        return True
    except Exception as e:
        print("[!] 关闭窗口失败: " + repr(e))
        return False


# ===================== 关闭已打开的 xlsx 文件 =====================
# 表格软件（Excel / WPS / LibreOffice）打开 .xlsx 后会独占文件，
# 应用再读同一个文件会失败，所以流程里先把该窗口关掉、释放文件锁。
_OFFICE_EXE_KEYS = ("excel.exe", "et.exe", "wps.exe", "soffice.exe", "libreoffice.exe",
                    "kingsoft", "wpsspreadsheet", "etmain", "soffice.bin")


def _win_process_exe_lower(hwnd):
    """窗口所属进程的可执行文件路径（小写）。"""
    try:
        return (_win_process_exe(_win_window_pid(hwnd)) or "").lower()
    except Exception:
        return ""


def _win_find_doc_windows(base_name, office_only=False):
    """查找标题里包含 base_name 的文档窗口。

    office_only=True 时只返回表格软件（Excel / WPS 等）打开的窗口。
    """
    out = []
    for hwnd, cls, title in _win_top_windows():
        if not title or cls == "#32770":
            continue
        if base_name and base_name not in title:
            continue
        exe = _win_process_exe_lower(hwnd)
        if office_only:
            if not any(k in exe for k in _OFFICE_EXE_KEYS):
                continue
        else:
            # 排除本脚本自己的 GUI 窗口与打包后的 exe
            if "auto_upload" in exe or "python" in exe:
                continue
        out.append((hwnd, cls, title, exe))
    return out


_PROMPT_DIALOG_TITLE_KEYS = (
    # Excel 系列（多语言/版本）
    "Microsoft Excel", "Excel", "Microsoft Office Excel",
    # WPS 系列
    "WPS", "金山", "WPS Office", "表格", "ET",
    # LibreOffice / OpenOffice
    "LibreOffice", "OpenOffice", "soffice",
    # 通用提示词（中英）
    "确认", "提示", "另存为", "保存", "警告", "错误",
    "Confirm", "Warning", "Save", "Error",
    # Excel 保存提示框典型标题（Excel 会把"文件名 - Excel"显示在标题栏）
    "未响应", "Not Responding",
)
_PROMPT_DIALOG_BTN_KEYS = (
    # 中文（优先顺序：不保存 / 否 → 取消 → 保存 / 是）
    "不保存", "否", "取消", "保存", "是", "确定", "重试",
    # 英文
    "Don't Save", "No", "Cancel", "Save", "Yes", "OK", "Retry",
)


def _win_find_prompt_dialog():
    """查找可能弹出的确认/提示框（保存修改、覆盖文件等）。

    优先返回属于 office 表格进程（Excel / WPS / LibreOffice）的弹窗，
    避免误把上传系统自己的弹窗当成 office 提示框。
    """
    fallback = (None, "", "")
    for hwnd, cls, title in _win_top_windows():
        if cls != "#32770":
            continue
        btn, txt = _win_find_button(hwnd, _PROMPT_DIALOG_BTN_KEYS)
        if not btn:
            continue

        # 1) 优先：窗口所属进程是 office 表格 → 直接命中
        exe = _win_process_exe_lower(hwnd)
        is_office_proc = any(k in exe for k in _OFFICE_EXE_KEYS)

        # 2) 标题命中关键词（office 或通用提示）
        hit_title = any(k.lower() in (title or "").lower() for k in _PROMPT_DIALOG_TITLE_KEYS)

        if is_office_proc or hit_title:
            return hwnd, title, txt

        # 3) 兜底：没有任何 office 特征时，保留第一个带按钮的 #32770
        if fallback[0] is None:
            fallback = (hwnd, title, txt)
    return fallback


def _win_dismiss_prompts(prefer=("不保存", "否", "Don't Save", "No", "取消", "Cancel")):
    """处理弹出的提示框：优先点「不保存 / 否」，避免覆盖已写好的数据。"""
    hwnd, title, txt = _win_find_prompt_dialog()
    if not hwnd:
        return False
    for p in prefer:
        btn, _ = _win_find_button(hwnd, (p,))
        if btn:
            print("[*] 弹窗「" + title + "」→ 点击「" + p + "」")
            _win_click_button(btn)
            time.sleep(0.4)
            return True
    # 没有可选项时点「确定 / 是」
    for p in ("确定", "OK", "是", "Yes"):
        btn, _ = _win_find_button(hwnd, (p,))
        if btn:
            print("[*] 弹窗「" + title + "」→ 点击「" + p + "」")
            _win_click_button(btn)
            time.sleep(0.4)
            return True
    return False


def _win_file_locked(path):
    """文件是否仍被其它程序占用（Excel 打开时以写方式打开会失败）。"""
    if not path or not os.path.isfile(path):
        return False
    try:
        with open(path, "r+b"):
            pass
        return False
    except PermissionError:
        return True
    except OSError:
        return False


def _step_close_xlsx_file():
    """打开 xlsx 后等待几秒，再关闭该文件（释放文件锁，便于后续导入）。

    关闭策略（按温和到暴力）：
      1) PostMessage WM_CLOSE 关闭窗口 + 处理「是否保存」弹窗；
      2) 等待窗口消失 + 文件锁释放（CLOSE_XLSX_WAIT 秒）；
      3) 仍失败 → 兜底：定点 terminate 持有目标文件名的 office 进程；
      4) 兜底后再次确认文件锁，释放则返回 True，否则返回 False。
    """
    path = DOWNLOAD_CONTEXT.get("path")
    if not path:
        print("[!] 未定位到 xlsx 文件，跳过关闭。")
        return False

    row_count = int(DOWNLOAD_CONTEXT.get("row_count") or 0)
    if row_count >= LARGE_DATA_ROW_THRESHOLD:
        print("[*] 全量数据 %d 行：上一步未打开 xlsx，跳过「关闭 xlsx」。" % row_count)
        return True

    if not CONFIG.get("CLOSE_XLSX_AFTER_OPEN", True):
        print("[*] 已关闭「打开后自动关闭 xlsx」，跳过。")
        return True

    base = os.path.splitext(os.path.basename(path))[0]
    print("[*] 等待 %d 秒后关闭已打开的 xlsx: %s" % (CLOSE_XLSX_DELAY, base))
    time.sleep(CLOSE_XLSX_DELAY)

    # 1) 找到打开该文件的窗口（先找表格软件，再退化为任意标题匹配的窗口）
    deadline = time.time() + 20
    wins = []
    while time.time() < deadline:
        wins = _win_find_doc_windows(base, office_only=True)
        if not wins:
            wins = _win_find_doc_windows(base, office_only=False)
        if wins:
            break
        time.sleep(0.5)

    if not wins:
        # 窗口没找到 → 先判断文件是不是还被锁（可能被后台进程持有）
        if _win_file_locked(path):
            print("[!] 未找到打开该 xlsx 的窗口，但文件仍被占用 → 进程级兜底关闭 ...")
            if _kill_office_procs_for_file(path, wait_timeout=15):
                print("[+] 进程级兜底关闭成功，文件锁已释放。")
                return True
            print("[!] 进程级兜底关闭仍未释放文件锁: " + path)
            return False
        print("[*] 未找到打开该 xlsx 的窗口，且文件未被占用，跳过关闭。")
        return True

    for hwnd, cls, title, exe in wins:
        print("[*] 关闭窗口: " + title + "（" + (exe or "未知程序") + "）")
        _win_close_window(hwnd)

    # 2) 等待窗口真正关闭，并顺带处理「是否保存」之类的弹窗
    deadline = time.time() + CLOSE_XLSX_WAIT
    while time.time() < deadline:
        alive = [w for w in wins if _win_window_exists(w[0])]
        if not alive:
            break
        _win_dismiss_prompts()
        time.sleep(0.5)

    # 3) 等待文件锁释放（Excel 完全退出后才会释放）
    deadline = time.time() + CLOSE_XLSX_WAIT
    while time.time() < deadline:
        if not _win_file_locked(path):
            print("[+] 已关闭 xlsx 文件并释放文件锁: " + path)
            return True
        _win_dismiss_prompts()
        time.sleep(0.5)

    # 4) WM_CLOSE 走完仍没释放锁 → 进程级兜底定点 terminate
    print("[!] 窗口关闭后文件锁仍未释放 → 进入进程级强制关闭兜底 ...")
    if _kill_office_procs_for_file(path, wait_timeout=15):
        print("[+] 进程级兜底关闭成功，文件锁已释放: " + path)
        return True

    print("[!] 进程级兜底关闭仍未释放文件锁，请手动关闭后再继续: " + path)
    return False


# ===================== 系统「打开」对话框自动化 =====================
def _win_find_open_dialog():
    """查找系统「打开」对话框，返回 (hwnd, title, button_text)。

    优先返回属于「上传系统」进程的对话框，避免误点其它软件的打开框。
    """
    want_exe = os.path.normcase(CONFIG.get("EXE_PATH") or "")
    fallback = (None, "", "")
    for hwnd, cls, title in _win_top_windows():
        if cls != "#32770":
            continue
        btn, txt = _win_find_button(hwnd, ("打开", "Open"))
        if not btn:
            continue
        if want_exe:
            proc_exe = _win_process_exe(_win_window_pid(hwnd))
            if proc_exe and os.path.normcase(proc_exe) == want_exe:
                return hwnd, title, txt
        if fallback[0] is None:
            fallback = (hwnd, title, txt)
    return fallback


def _win_find_filename_edit(dlg_hwnd):
    """定位「文件名(N)」输入框，返回 (hwnd, 当前文本)。

    Vista 风格对话框里该输入框是 ComboBoxEx32 → ComboBox → Edit 结构；
    对话框顶部还有一个「搜索」输入框也是 ComboBox+Edit，因此这里再按
    「位置最靠下（文件名框在底部）+ 内容像文件名」来区分。
    """
    if not dlg_hwnd:
        return None, ""

    candidates = []
    for h, cls, txt in _win_all_descendants(dlg_hwnd):
        if cls != "Edit":
            continue
        parent = 0
        try:
            parent = int(_user32.GetParent(wintypes.HWND(h)) or 0)
        except Exception:
            parent = 0
        pcls = _win_class(parent) if parent else ""
        rect = _win_get_rect(h)
        # 跨进程读 Edit 文本要用 WM_GETTEXT（GetWindowTextW 读不到别的进程的控件）
        real = _win_get_text_msg(h)
        if not real:
            real = txt
        candidates.append({
            "hwnd": h, "txt": real, "pcls": pcls, "top": rect[1],
            "filename_like": _win_looks_like_filename(real),
        })

    if not candidates:
        return None, ""

    # 1) 父级是 ComboBox / ComboBoxEx32 的 Edit（即文件名 / 搜索输入框）
    combos = [c for c in candidates if c["pcls"] in ("ComboBox", "ComboBoxEx32")]
    if combos:
        combos.sort(key=lambda c: (not c["filename_like"], -c["top"]))
        return combos[0]["hwnd"], combos[0]["txt"]

    # 2) 内容已经是文件名的 Edit（位置最靠下的优先）
    named = [c for c in candidates if c["filename_like"]]
    if named:
        named.sort(key=lambda c: -c["top"])
        return named[0]["hwnd"], named[0]["txt"]

    # 3) 兜底：任意非空 Edit（位置最靠下的优先）
    nonempty = [c for c in candidates if c["txt"]]
    if nonempty:
        nonempty.sort(key=lambda c: -c["top"])
        return nonempty[0]["hwnd"], nonempty[0]["txt"]

    # 4) 最后兜底：第一个 Edit
    return candidates[0]["hwnd"], candidates[0]["txt"]


def _step_select_file_in_open_dialog():
    """在系统「打开」对话框中选中指定的 xlsx，并点击「打开」。"""
    if not _WIN32 or _user32 is None:
        print("[!] 当前环境不支持原生对话框自动化，无法自动选择文件。")
        return False

    path = DOWNLOAD_CONTEXT.get("path")
    if not path:
        print("[!] 未定位到 xlsx 文件，无法选择。")
        return False
    if not os.path.isfile(path):
        print("[!] 文件不存在: " + path)
        return False

    print("[*] 等待系统「打开」对话框（资源管理器风格文件选择框）弹出 ...")
    dlg = 0
    title = ""
    btn_text = ""
    # 上一步已经确认过对话框弹出，这里优先复用（避免漏检），否则重新等待
    cached = IMPORT_CONTEXT.get("dialog", 0)
    if cached and _win_window_exists(cached):
        dlg, title, btn_text = _win_find_open_dialog()
        if not dlg:
            dlg = cached
    if not dlg:
        deadline = time.time() + OPEN_DIALOG_TIMEOUT
        while time.time() < deadline:
            dlg, title, btn_text = _win_find_open_dialog()
            if dlg:
                break
            time.sleep(0.5)

    if not dlg:
        print("[!] 等待「打开」对话框超时（%d 秒）。" % OPEN_DIALOG_TIMEOUT)
        return False

    print("[+] 已定位「打开」对话框：标题=" + title + "，按钮=" + btn_text)

    # 定位「文件名」输入框
    edit, cur = _win_find_filename_edit(dlg)
    if not edit:
        print("[!] 未找到「文件名」输入框。")
        return False
    print("[*] 文件名输入框 hwnd=" + str(edit) + "，当前内容: " + (cur or "(空)"))

    btn, _btxt = _win_find_button(dlg, ("打开", "Open"))
    if not btn:
        print("[!] 未找到「打开」按钮")
        return False

    base = os.path.basename(path)
    want_dir = os.path.dirname(path)

    # ---- 选文件策略（实测结论）----
    # 1) 在「文件名」框里写「完整路径」时，外壳会把它规范化成正斜杠
    #    "C:/ProgramData/..."，点「打开」会报「文件名无效」，对话框关不掉；
    # 2) 写「纯文件名」则完全正常（前提是对话框当前目录已是该文件所在目录）；
    # 3) 所以：目录已对 → 直接写文件名；目录不对 → 先写全路径让外壳跳目录，
    #    下一轮再写文件名。
    cur_dir = _win_get_current_folder(dlg)
    if cur_dir:
        print("[*] 对话框当前目录: " + cur_dir)
    same_dir = bool(cur_dir) and (
        os.path.normcase(os.path.abspath(cur_dir))
        == os.path.normcase(os.path.abspath(want_dir)))
    plan = [base] if same_dir else [path, base]
    print("[*] 目标目录: " + want_dir + "，选文件方案: " + str(plan))

    # 写入并点「打开」（最多重试 3 次）
    for attempt in range(1, 4):
        # 先清掉可能残留的错误提示框（例如上一次失败留下的「文件名无效」）
        _win_dismiss_error_boxes()

        use_text = plan[min(attempt - 1, len(plan) - 1)]
        _win_set_text(edit, use_text)
        time.sleep(0.3)

        got = _win_get_text_msg(edit)
        if got and not got.strip().lower().endswith(base.lower()):
            print("[!] 文件名输入框内容校验失败（当前: " + got + "），重试 ...")
            continue
        print("[*] 第 %d 次点击「打开」按钮（文件名框内容: %s） ..."
              % (attempt, got or use_text))
        _win_click_button(btn)

        deadline = time.time() + max(3, OPEN_DIALOG_CLICK_TIMEOUT // 3)
        while time.time() < deadline:
            if not _win_find_open_dialog()[0]:
                print("[+] 「打开」对话框已关闭，已提交文件: " + path)
                return True
            time.sleep(0.3)

        print("[!] 点击后对话框仍未关闭 ...")
        # 对话框可能还在初始化或被错误框挡住，重新取一次句柄
        dlg2, _t, _b = _win_find_open_dialog()
        if dlg2:
            dlg = dlg2
            edit, _cur = _win_find_filename_edit(dlg)
            btn, _btxt = _win_find_button(dlg, ("打开", "Open"))
            # 目录可能已被外壳切换过，重新判断
            cur_dir = _win_get_current_folder(dlg)
            if cur_dir and os.path.normcase(os.path.abspath(cur_dir)) == \
                    os.path.normcase(os.path.abspath(want_dir)):
                plan = [base] + [p for p in plan if p != base]

    _win_dismiss_error_boxes()
    print("[!] 多次点击「打开」后对话框仍未关闭（可能文件选择失败）。")
    return False


# ===================== 页面「选择数据表格导入」按钮 =====================
# 导入过程中的上下文（记录已弹出的文件选择框等）
IMPORT_CONTEXT: dict = {}

JS_CLICK_IMPORT_TABLE = r"""
(function () {
    const KEY = '选择数据表格导入';
    function norm(s) { return ((s || '') + '').replace(/\s+/g, ''); }
    const KEYN = norm(KEY);
    function textOf(el) { return norm(el.innerText || el.textContent || ''); }

    const all = Array.from(document.querySelectorAll('button, a, div, span, label, i, p, em'));
    const hits = all.filter(e => textOf(e).includes(KEYN));
    if (!hits.length) {
        return { ok: false, msg: '未找到「' + KEY + '」按钮' };
    }

    // 关键：优先选真正可点击的元素（button / a / .arco-btn），
    // 否则会点到外层容器 DIV（如 .upload-card），点击不会有任何反应。
    const clickables = hits.filter(e =>
        /^(BUTTON|A)$/i.test(e.tagName) ||
        e.getAttribute('role') === 'button' ||
        /arco-btn|upload-btn/i.test((e.className || '') + ''));
    const pool = clickables.length ? clickables : hits;
    pool.sort((a, b) => textOf(a).length - textOf(b).length);
    const el = pool[0];

    if (el.disabled || el.getAttribute('aria-disabled') === 'true') {
        return { ok: false, msg: '「' + KEY + '」按钮当前不可用' };
    }
    try { el.scrollIntoView({ block: 'center' }); } catch (e) {}
    el.click();
    return {
        ok: true,
        tag: el.tagName,
        cls: (el.className || '') + '',
        text: textOf(el),
        candidates: hits.length,
        clickables: clickables.length
    };
})()
"""

# 备用方案：直接点击页面里的 <input type="file">（有些实现按钮只是触发它）
JS_IMPORT_BUTTON_RECT = r"""
(function () {
    const KEY = '选择数据表格导入';
    function norm(s) { return ((s || '') + '').replace(/\s+/g, ''); }
    const KEYN = norm(KEY);
    const all = Array.from(document.querySelectorAll('button, a, div, span, label, i, p, em'));
    const hits = all.filter(e => norm(e.innerText || e.textContent || '').indexOf(KEYN) >= 0);
    const clickables = hits.filter(e =>
        /^(BUTTON|A)$/i.test(e.tagName) ||
        e.getAttribute('role') === 'button' ||
        /arco-btn/i.test((e.className || '') + ''));
    const pool = clickables.length ? clickables : hits;
    pool.sort((a, b) => norm(a.innerText || a.textContent || '').length
                       - norm(b.innerText || b.textContent || '').length);
    const el = pool[0];
    if (!el) return null;
    try { el.scrollIntoView({ block: 'center' }); } catch (e) {}
    const r = el.getBoundingClientRect();
    return {
        x: Math.round(r.x + r.width / 2),
        y: Math.round(r.y + r.height / 2),
        w: Math.round(r.width),
        h: Math.round(r.height),
        tag: el.tagName,
        cls: (el.className || '') + ''
    };
})()
"""

# 备用方案：直接点击页面里的 <input type="file">（有些实现按钮只是触发它）
JS_CLICK_FILE_INPUT = r"""
(function () {
    const inputs = Array.from(document.querySelectorAll('input[type="file"]'));
    if (!inputs.length) return { ok: false, msg: '页面里没有 input[type=file]' };
    for (const inp of inputs) {
        if (inp.disabled) continue;
        try { inp.click(); } catch (e) {}
        return { ok: true, count: inputs.length };
    }
    return { ok: false, msg: '所有 input[type=file] 都不可用' };
})()
"""

# 页面已导入过表格时，原按钮变成「重新选择」，需要先点它
JS_FIND_RESELECT_BTN = r"""
(function () {
    const KEY = '重新选择';
    function norm(s) { return ((s || '') + '').replace(/\s+/g, ''); }
    const KEYN = norm(KEY);
    const all = Array.from(document.querySelectorAll('button, a, div, span, label, i, p, em'));
    let hits = all.filter(e => norm(e.innerText || e.textContent || '') === KEYN);
    if (!hits.length) {
        hits = all.filter(e => norm(e.innerText || e.textContent || '').indexOf(KEYN) >= 0);
    }
    if (!hits.length) return null;
    const clickables = hits.filter(e =>
        /^(BUTTON|A)$/i.test(e.tagName) ||
        e.getAttribute('role') === 'button' ||
        /arco-btn/i.test((e.className || '') + ''));
    const pool = clickables.length ? clickables : hits;
    pool.sort((a, b) => norm(a.innerText || a.textContent || '').length
                       - norm(b.innerText || b.textContent || '').length);
    const el = pool[0];
    if (!el) return null;
    try { el.scrollIntoView({ block: 'center' }); } catch (e) {}
    const r = el.getBoundingClientRect();
    return {
        x: Math.round(r.x + r.width / 2),
        y: Math.round(r.y + r.height / 2),
        tag: el.tagName,
        cls: (el.className || '') + ''
    };
})()
"""

# 「重新选择」按钮的 JS 点击（兜底用）
JS_CLICK_RESELECT = r"""
(function () {
    const KEY = '重新选择';
    function norm(s) { return ((s || '') + '').replace(/\s+/g, ''); }
    const KEYN = norm(KEY);
    const all = Array.from(document.querySelectorAll('button, a, div, span, label, i, p, em'));
    let hits = all.filter(e => norm(e.innerText || e.textContent || '') === KEYN);
    if (!hits.length) {
        hits = all.filter(e => norm(e.innerText || e.textContent || '').indexOf(KEYN) >= 0);
    }
    if (!hits.length) return { ok: false, msg: '未找到「' + KEY + '」按钮' };
    const clickables = hits.filter(e =>
        /^(BUTTON|A)$/i.test(e.tagName) ||
        e.getAttribute('role') === 'button' ||
        /arco-btn/i.test((e.className || '') + ''));
    const pool = clickables.length ? clickables : hits;
    pool.sort((a, b) => norm(a.innerText || a.textContent || '').length
                       - norm(b.innerText || b.textContent || '').length);
    const el = pool[0];
    el.click();
    return { ok: true, tag: el.tagName };
})()
"""

# 「确认重新导入」弹窗（点「重新选择」后出现），需要点里面的「确认导入」
JS_CONFIRM_REIMPORT = r"""(function () {
    const body = document.body.innerText || '';
    if (body.indexOf('确认重新导入') < 0) return { found: false };
    const scopes = Array.from(document.querySelectorAll(
        '.arco-modal, .arco-modal-wrapper, [role="dialog"]'));
    const list = scopes.length ? scopes : [document];
    for (const scope of list) {
        const btns = Array.from(scope.querySelectorAll('button'))
            .filter(b => ((b.innerText || '').trim()) === '确认导入');
        if (btns.length) {
            const r = btns[0].getBoundingClientRect();
            return {
                found: true,
                x: Math.round(r.x + r.width / 2),
                y: Math.round(r.y + r.height / 2),
                text: '确认导入'
            };
        }
    }
    return { found: true };
})()
"""

# 检查页面上是否已出现导入的文件名 / 导入完成提示（软确认，不阻断流程）
JS_CHECK_IMPORT_RESULT = r"""
(function () {
    const name = %s;
    const body = (document.body.innerText || '') + '';
    const flat = body.replace(/\s+/g, '');
    const byName = !!name && body.indexOf(name) >= 0;
    const byDone = flat.indexOf('导入成功') >= 0 || flat.indexOf('导入完成') >= 0
                || flat.indexOf('共导入') >= 0 || flat.indexOf('已选择文件') >= 0;
    return {
        found: byName || byDone,
        byName: byName,
        byDone: byDone,
        snippet: body.slice(0, 200)
    };
})()
"""

# 导入后应用会弹出「编号失效确认」（模板里的编号在系统中已存在），需要点「确认失效」
# 注意：
# 1) Arco 弹窗关闭后 DOM 里的文字仍然存在，所以必须用「可见性」判断，
#    否则会把已关闭的弹窗误判成还开着；
# 2) 弹窗刚出现时表格还在异步加载，会显示「暂无数据」，此时点「确认失效」
#    前端会直接忽略、弹窗不关闭，必须等数据加载出来再点。
JS_CHECK_INVALID_MODAL = r"""
(function () {
    const scopes = Array.from(document.querySelectorAll(
        '.arco-modal, .arco-modal-wrapper, [role="dialog"]'));
    for (const scope of scopes) {
        const r = scope.getBoundingClientRect();
        if (r.width <= 0 || r.height <= 0) {
            continue;                       // 已关闭（隐藏）的弹窗
        }
        const txt = scope.innerText || '';
        if (txt.indexOf('编号失效确认') < 0) {
            continue;
        }
        // 数据是否已加载：有数据行 or 只有「暂无数据」占位
        const rows = scope.querySelectorAll('.arco-table-tbody .arco-table-tr').length;
        const empty = (rows === 0) && (txt.indexOf('暂无数据') >= 0);

        function pick(name) {
            const bs = Array.from(scope.querySelectorAll('button'))
                .filter(b => ((b.innerText || '').trim()) === name);
            for (const b of bs) {
                const br = b.getBoundingClientRect();
                if (br.width > 0 && br.height > 0) {
                    return {
                        x: Math.round(br.x + br.width / 2),
                        y: Math.round(br.y + br.height / 2)
                    };
                }
            }
            return null;
        }

        return {
            found: true,
            empty: empty,
            rows: rows,
            btn: pick('确认失效'),
            cancel: pick('取消')
        };
    }
    return { found: false };
})()
"""


def _step_handle_invalid_confirm_modal():
    """导入后若弹出「编号失效确认」弹窗，点「确认失效」确认（把这些报告编号置为失效）。

    流程：等弹窗出现 → 等表格数据加载完（避免「暂无数据」时点击无效）
          → 点「确认失效」→ 校验弹窗是否真的关闭，未关闭则重试。
    """
    assert cdp_client is not None, "cdp_client 未初始化"

    # 1) 等弹窗出现
    info = None
    deadline = time.time() + 15
    while time.time() < deadline:
        try:
            info = cdp_client.eval_js(JS_CHECK_INVALID_MODAL)
        except Exception as e:
            print("[!] 检查「编号失效确认」弹窗失败: " + repr(e))
            return True
        if info and info.get("found"):
            break
        time.sleep(0.5)

    if not info or not info.get("found"):
        print("[*] 未出现「编号失效确认」弹窗。")
        return True
    if not info.get("btn"):
        print("[!] 出现「编号失效确认」弹窗，但没找到可见的「确认失效」按钮，"
              "弹窗会挡住后续操作，终止流程。")
        return False

    # 2) 等表格数据加载完（「暂无数据」时点击会被前端忽略）
    deadline = time.time() + 12
    while time.time() < deadline:
        if not info.get("empty"):
            break
        time.sleep(0.5)
        try:
            info = cdp_client.eval_js(JS_CHECK_INVALID_MODAL)
        except Exception as e:
            print("[!] 轮询「编号失效确认」弹窗失败: " + repr(e))
            break
        if not info or not info.get("found"):
            break

    if not info or not info.get("found"):
        print("[*] 「编号失效确认」弹窗已消失（无需处理）。")
        return True

    if info.get("empty"):
        # 等不到数据时无法点「确认失效」（点了也不会生效），只能取消
        print("[!] 「编号失效确认」弹窗表格始终为「暂无数据」，无法确认失效，改点「取消」关闭。")
        cancel = info.get("cancel")
        if cancel:
            _cdp_click_at(cdp_client, int(cancel["x"]), int(cancel["y"]))
            time.sleep(1)
        return True

    print("[*] 「编号失效确认」弹窗数据已加载，点击「确认失效」 ...")

    # 3) 点「确认失效」，并校验弹窗是否真的关闭（最多 3 次）
    for attempt in range(1, 4):
        if not info or not info.get("btn"):
            info = cdp_client.eval_js(JS_CHECK_INVALID_MODAL)
        if not info or not info.get("found"):
            print("[+] 「编号失效确认」弹窗已关闭。")
            return True
        btn = info.get("btn")
        if not btn:
            break
        print("[*] 第 %d 次点击「确认失效」（坐标 %d,%d） ..."
              % (attempt, int(btn["x"]), int(btn["y"])))
        _cdp_click_at(cdp_client, int(btn["x"]), int(btn["y"]))

        # 等待弹窗关闭
        deadline = time.time() + 5
        while time.time() < deadline:
            time.sleep(0.4)
            try:
                info = cdp_client.eval_js(JS_CHECK_INVALID_MODAL)
            except Exception:
                info = None
            if not info or not info.get("found"):
                print("[+] 已点击「确认失效」，弹窗已关闭。")
                return True
        print("[!] 点击后弹窗未关闭，准备重试 ...")

    # 多次点击仍不关闭：若是「暂无数据」状态（前端会忽略点击），只能关掉弹窗
    if info and info.get("found") and info.get("empty") and info.get("cancel"):
        print("[!] 弹窗表格为「暂无数据」，前端会忽略「确认失效」点击，改点「取消」关闭。")
        _cdp_click_at(cdp_client, int(info["cancel"]["x"]), int(info["cancel"]["y"]))
        time.sleep(1)
        return True

    print("[!] 「确认失效」点击后弹窗仍未关闭，终止流程，请手动确认后重试。")
    return False


def _cdp_click_at(client, x, y):
    """用 CDP Input 事件在指定坐标做一次「真实」点击（trusted click）。

    JS 合成的 .click() 事件是 untrusted 的，Electron 里点文件选择按钮
    可能不会弹出系统对话框；用 Input.dispatchMouseEvent 才是浏览器认可的
    用户手势。
    """
    try:
        client.call("Input.dispatchMouseEvent",
                    {"type": "mouseMoved", "x": x, "y": y, "button": "none",
                     "clickCount": 0})
        time.sleep(0.05)
        client.call("Input.dispatchMouseEvent",
                    {"type": "mousePressed", "x": x, "y": y, "button": "left",
                     "clickCount": 1})
        # 按下与抬起之间留一点间隔，部分 UI 框架（Arco 等）才认为是完整点击
        time.sleep(0.08)
        client.call("Input.dispatchMouseEvent",
                    {"type": "mouseReleased", "x": x, "y": y, "button": "left",
                     "clickCount": 1})
        return True
    except Exception as e:
        print("[!] CDP 模拟点击失败: " + repr(e))
        return False


def _wait_for_open_dialog(timeout=20):
    """等待系统「打开」对话框（资源管理器风格的文件选择框）弹出。

    返回 (hwnd, title, button_text)；超时返回 (0, "", "")。
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        dlg, title, btn = _win_find_open_dialog()
        if dlg:
            return dlg, title, btn
        time.sleep(0.3)
    return 0, "", ""


def _eval_js_quick(client, js, timeout=3, quiet=False):
    """带短超时的 JS 求值。

    页面弹出原生模态框（如系统文件选择框）时渲染进程会被阻塞，
    Runtime.evaluate 会一直挂着；这里用短超时快速返回 None，
    避免整个流程卡住。quiet=True 时不打印单次超时日志（由调用方统一节流）。
    """
    old = None
    try:
        old = client.ws.gettimeout()
    except Exception:
        old = None
    try:
        client.ws.settimeout(timeout)
        return client.eval_js(js)
    except Exception as e:
        if not quiet:
            print("[*] 页面 JS 检测超时/异常（可能被原生弹窗阻塞）: " + type(e).__name__)
        return None
    finally:
        try:
            client.ws.settimeout(old if old else 30)
        except Exception:
            pass


def _wait_import_ready_dialog(timeout=20):
    """点击导入按钮后，等待进入「可以选文件」的状态。

    两种情况都能处理：
      A) 直接弹出系统文件选择框（资源管理器风格的「打开」框）→ 成功；
      B) 先弹出应用内「确认重新导入」弹窗 → 停留几秒后点「确认导入」，再继续等 A。

    顺序很重要：先用 Win32 查原生框（不依赖页面），页面未被阻塞时
    才去查应用内弹窗；否则 eval 会被原生模态框挂住。
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        dlg, title, btn = _win_find_open_dialog()
        if dlg:
            print("[+] 文件选择对话框已弹出：标题=" + title + "，按钮=" + btn)
            IMPORT_CONTEXT["dialog"] = dlg
            IMPORT_CONTEXT["title"] = title
            return True

        # 没有原生框时才去页面里找「确认重新导入」弹窗
        info = _eval_js_quick(cdp_client, JS_CONFIRM_REIMPORT, timeout=3)
        if info and info.get("found"):
            if not info.get("x"):
                print("[!] 出现「确认重新导入」弹窗，但没找到「确认导入」按钮")
                return False
            print("[*] 出现「确认重新导入」弹窗，停留 %d 秒后点击「确认导入」 ..."
                  % IMPORT_REIMPORT_CONFIRM_DELAY)
            time.sleep(IMPORT_REIMPORT_CONFIRM_DELAY)
            info2 = _eval_js_quick(cdp_client, JS_CONFIRM_REIMPORT, timeout=3)
            if info2 and info2.get("x"):
                _cdp_click_at(cdp_client, int(info2["x"]), int(info2["y"]))
                time.sleep(1)
                print("[+] 已点击「确认导入」，继续等待文件选择框 ...")
                continue
            print("[!] 重新获取「确认导入」按钮坐标失败")
            return False

        time.sleep(0.5)

    print("[!] 等待文件选择框弹出超时（%d 秒）" % timeout)
    return False


def _wait_and_confirm_reimport():
    """等待「确认重新导入」弹窗，停留几秒后点「确认导入」。

    场景：页面已导入过表格时按钮变成「重新选择」，点它会弹出
    「导入表格将覆盖原有数据，是否继续导入？」，需要点「确认导入」
    之后才会弹出系统文件选择框。
    """
    deadline = time.time() + 15
    info = None
    while time.time() < deadline:
        info = _eval_js_quick(cdp_client, JS_CONFIRM_REIMPORT, timeout=3)
        if info and info.get("found"):
            break
        time.sleep(0.3)

    if not info or not info.get("found"):
        print("[!] 未出现「确认重新导入」弹窗")
        return False
    if not info.get("x"):
        print("[!] 「确认重新导入」弹窗里没找到「确认导入」按钮")
        return False

    print("[*] 出现「确认重新导入」弹窗，停留 %d 秒后点击「确认导入」 ..."
          % IMPORT_REIMPORT_CONFIRM_DELAY)
    time.sleep(IMPORT_REIMPORT_CONFIRM_DELAY)

    # 重新取一次坐标（弹窗可能有动画位移）
    info = _eval_js_quick(cdp_client, JS_CONFIRM_REIMPORT, timeout=3)
    if not info or not info.get("x"):
        print("[!] 重新获取「确认导入」按钮坐标失败")
        return False
    _cdp_click_at(cdp_client, int(info["x"]), int(info["y"]))
    time.sleep(1)
    print("[+] 已点击「确认导入」。")
    return True


def _step_click_import_table():
    """点击页面中的「选择数据表格导入」按钮。

    - 正常情况：页面上有「选择数据表格导入」按钮，点击后弹出文件选择框；
    - 已导入过的情况：按钮变成「重新选择」，需要先点它，
      再在「确认重新导入」弹窗里点「确认导入」，之后才会弹出文件选择框。

    成功判据：文件选择对话框（资源管理器风格的「打开」框）真的弹出来了。
    若点击后对话框没出现（按钮没生效），会重新点击，最多 3 轮。
    """
    assert cdp_client is not None, "cdp_client 未初始化"

    # 关闭 xlsx 后先稳定一下，避免页面还在处理文件锁
    wait_sec = max(1, IMPORT_BTN_SETTLE_DELAY)
    print("[*] 等待 %d 秒后再点击「选择数据表格导入」 ..." % wait_sec)
    time.sleep(wait_sec)

    last_res = None
    for attempt in range(1, IMPORT_BTN_MAX_ROUNDS + 1):
        # 1) 定位按钮，取它的屏幕坐标
        #    优先「选择数据表格导入」；页面上若已选过表格，该按钮会变成「重新选择」
        rect = None
        kind = "import"
        deadline = time.time() + 15
        while time.time() < deadline:
            try:
                rect = cdp_client.eval_js(JS_IMPORT_BUTTON_RECT)
            except Exception as e:
                print("[!] 定位「选择数据表格导入」按钮异常: " + repr(e))
                rect = None
            if rect and rect.get("x"):
                break
            # 没有该按钮时，看看是不是「重新选择」
            try:
                rect = cdp_client.eval_js(JS_FIND_RESELECT_BTN)
            except Exception as e:
                print("[!] 定位「重新选择」按钮异常: " + repr(e))
                rect = None
            if rect and rect.get("x"):
                kind = "reselect"
                break
            time.sleep(0.5)

        if not rect:
            print("[!] 第 %d 轮未找到「选择数据表格导入」/「重新选择」按钮" % attempt)
            # 最后一轮再试试直接点 input[type=file]
            if attempt >= IMPORT_BTN_MAX_ROUNDS:
                try:
                    fres = cdp_client.eval_js(JS_CLICK_FILE_INPUT, user_gesture=True)
                except Exception as e:
                    fres = {"ok": False, "msg": repr(e)}
                print("[*] 尝试直接点击 input[type=file]: " + str(fres))
                dlg, title, btn = _wait_for_open_dialog(IMPORT_DIALOG_APPEAR_TIMEOUT)
                if dlg:
                    print("[+] 文件选择对话框已弹出：标题=" + title + "，按钮=" + btn)
                    IMPORT_CONTEXT["dialog"] = dlg
                    IMPORT_CONTEXT["title"] = title
                    return True
            continue

        # 2) 用 CDP Input 事件做一次真实点击（最可靠，用户手势）
        btn_name = "重新选择" if kind == "reselect" else "选择数据表格导入"
        print("[*] 第 %d 轮：模拟真实点击「%s」（%s，%s，坐标 %d,%d） ..."
              % (attempt, btn_name, str(rect.get("tag")), str(rect.get("cls")),
                 int(rect.get("x", 0)), int(rect.get("y", 0))))
        _cdp_click_at(cdp_client, int(rect.get("x", 0)), int(rect.get("y", 0)))

        # 2.5) 等待进入「可以选文件」状态：
        #      直接弹出系统文件选择框，或先弹「确认重新导入」（停留 2 秒后点确认导入）
        if _wait_import_ready_dialog(IMPORT_DIALOG_APPEAR_TIMEOUT):
            return True

        # 3) 没弹出来，再用 JS .click()（带用户手势）试一次
        print("[!] 第 %d 轮真实点击后未检测到文件选择对话框，改用 JS click 重试 ..."
              % attempt)
        try:
            if kind == "reselect":
                res = cdp_client.eval_js(JS_CLICK_RESELECT, user_gesture=True)
            else:
                res = cdp_client.eval_js(JS_CLICK_IMPORT_TABLE, user_gesture=True)
        except Exception as e:
            print("[!] JS click 异常: " + repr(e))
            res = None
        last_res = res

        if _wait_import_ready_dialog(IMPORT_DIALOG_APPEAR_TIMEOUT):
            return True

        time.sleep(1)

    print("[!] 点击「选择数据表格导入」失败（未弹出文件选择对话框）: " + str(last_res))
    return False


def _step_check_import_result():
    """软确认：页面上是否出现「导入完成」标志。

    这一步不是关键步骤：数据表格是否提交成功由第 10 步「打开对话框关闭」判定，
    这里只是给日志一个提示。应用文案变化、页面正在解析大 xlsx、或被原生弹窗
    阻塞时都不得阻断流程，否则会卡在第 13 步并中断后续选文件夹 / 上传。

    单次探测用 3 秒短超时，每 5 秒回报一次进度，最多等 IMPORT_RESULT_TIMEOUT
    秒后自动放行。
    """
    assert cdp_client is not None, "cdp_client 未初始化"
    path = DOWNLOAD_CONTEXT.get("path") or ""
    base = os.path.basename(path)
    if not base:
        return True

    js = JS_CHECK_IMPORT_RESULT % json.dumps(base, ensure_ascii=False)
    t0 = time.time()
    deadline = t0 + IMPORT_RESULT_TIMEOUT
    last_report = t0
    last_res = None
    blocked = 0

    while time.time() < deadline:
        res = _eval_js_quick(cdp_client, js, timeout=3, quiet=True)
        if res is None:
            blocked += 1
        elif isinstance(res, dict):
            last_res = res
            if res.get("found"):
                how = "文件名" if res.get("byName") else "导入完成提示"
                print("[+] 页面上已确认数据表格导入（依据：%s，用时 %d 秒）"
                      % (how, int(time.time() - t0)))
                return True

        now = time.time()
        if now - last_report >= IMPORT_RESULT_PROGRESS_LOG:
            last_report = now
            print("[*] 仍在等待页面确认导入 ... 已等待 %d 秒（上限 %d 秒，超时自动继续）%s"
                  % (int(now - t0), IMPORT_RESULT_TIMEOUT,
                     "，页面暂时无响应" if blocked else ""))
        time.sleep(1)

    snippet = ""
    if isinstance(last_res, dict):
        snippet = str(last_res.get("snippet") or "").replace("\n", " ").strip()[:120]
    print("[*] 未在页面上匹配到导入完成标志（软确认，不阻断流程，继续执行后续步骤）。")
    if snippet:
        print("[*] 页面当前文本片段: " + snippet)
    if blocked:
        print("[*] 期间有 %d 次页面无响应（可能被原生弹窗阻塞或应用正在解析表格）。" % blocked)
    return True


# ===================== 应用日志（用于精确定位保存路径） =====================
# 应用在写盘成功后会记录： 最新模板文件保存成功 {"savePath":"C:\\...\\数据导入模板.xlsx"}
_RE_SAVE_PATH = re.compile(r'"savePath"\s*:\s*"((?:[^"\\]|\\.)*)"')


def _app_log_dirs():
    """应用日志目录（安装目录 logs + AppData logs）。"""
    dirs = []
    exe = CONFIG.get("EXE_PATH") or ""
    if exe:
        d = os.path.join(os.path.dirname(exe), "logs")
        if os.path.isdir(d):
            dirs.append(d)
    appdata = os.environ.get("APPDATA")
    if appdata:
        d = os.path.join(appdata, "detection-report-fixation-tool", "logs")
        if os.path.isdir(d):
            dirs.append(d)
    return dirs


def _app_log_snapshot():
    """记录当前各日志文件的大小，之后只读取新增内容。"""
    snap = {}
    for d in _app_log_dirs():
        try:
            names = os.listdir(d)
        except OSError:
            continue
        for name in names:
            if name.startswith("app-") and name.endswith(".log"):
                p = os.path.join(d, name)
                try:
                    snap[p] = os.path.getsize(p)
                except OSError:
                    continue
    return snap


def _read_log_delta(snap):
    """读取日志中新增的内容。"""
    chunks = []
    for d in _app_log_dirs():
        try:
            names = os.listdir(d)
        except OSError:
            continue
        for name in names:
            if not (name.startswith("app-") and name.endswith(".log")):
                continue
            p = os.path.join(d, name)
            prev = (snap or {}).get(p, 0)
            try:
                with open(p, "r", encoding="utf-8", errors="ignore") as f:
                    f.seek(prev)
                    chunks.append(f.read())
            except OSError:
                continue
    return "\n".join(chunks)


def _extract_save_paths(text):
    """从日志文本里解析 savePath（JSON 中的 \\\\ 需还原成单个反斜杠）。"""
    out = []
    for m in _RE_SAVE_PATH.finditer(text or ""):
        out.append(m.group(1).replace("\\\\", "\\"))
    return out


def _candidate_dirs():
    """模板文件可能保存到的目录。"""
    dirs = []
    program_data = os.environ.get("ProgramData") or r"C:\ProgramData"
    dirs.append(os.path.join(program_data, "AreportFile"))
    # 兜底：原生对话框会记住上次目录，上一轮选过 OrigDoc 后，
    # 本轮的保存对话框可能默认到 OrigDoc，xlsx 可能落在这里
    dirs.append(os.path.join(program_data, "AreportFile", "OrigDoc"))
    home = os.path.expanduser("~")
    for sub in ("Downloads", "Desktop", "Documents"):
        dirs.append(os.path.join(home, sub))
    exe = CONFIG.get("EXE_PATH") or ""
    if exe:
        dirs.append(os.path.dirname(exe))
    seen, out = set(), []
    for d in dirs:
        key = os.path.normcase(os.path.abspath(d))
        if key in seen:
            continue
        seen.add(key)
        if os.path.isdir(d):
            out.append(d)
    return out


def _scan_for_xlsx(filename, since_ts):
    """在候选目录中查找刚保存的 .xlsx：优先同名文件，否则取最新的那个。"""
    want = os.path.basename(filename) if filename else ""
    same_name, newest = None, None
    newest_mt = 0.0
    for d in _candidate_dirs():
        try:
            entries = os.listdir(d)
        except OSError:
            continue
        for name in entries:
            if not name.lower().endswith(".xlsx"):
                continue
            p = os.path.join(d, name)
            try:
                mt = os.path.getmtime(p)
            except OSError:
                continue
            if mt + 2 < since_ts:
                continue
            if want and name == want:
                same_name = p
                break
            if mt > newest_mt:
                newest_mt = mt
                newest = p
        if same_name:
            break
    return same_name or newest


# ===================== 首页「下载表格模板」流程 =====================
# 页面上「下载表格模板」是一个 arco 按钮（button[type=primary]），文本完全匹配即可定位。
JS_CLICK_DOWNLOAD_TEMPLATE = r"""
(function () {
    const KEY = '下载表格模板';
    function textOf(el) {
        return ((el.innerText || el.textContent || '') + '').trim();
    }
    const buttons = Array.from(document.querySelectorAll('button'));
    let btn = buttons.find(b => textOf(b) === KEY);
    if (!btn) btn = buttons.find(b => textOf(b).includes(KEY));
    if (!btn) {
        const others = Array.from(document.querySelectorAll('a, div, span'));
        btn = others.find(e => textOf(e) === KEY);
    }
    if (!btn) {
        return { ok: false, msg: '未找到「' + KEY + '」按钮' };
    }
    if (btn.disabled || btn.getAttribute('aria-disabled') === 'true') {
        return { ok: false, msg: '「' + KEY + '」按钮当前不可用（可能正在下载）' };
    }
    try { btn.scrollIntoView({ block: 'center' }); } catch (e) {}
    btn.click();
    return { ok: true, text: textOf(btn), tag: btn.tagName, cls: (btn.className || '') + '' };
})()
"""

# 三个步骤之间传递的上下文
DOWNLOAD_CONTEXT: dict = {}


# ---- 下载新模板前：关掉已经打开的模板 xlsx ----
# 新模板要写回 %ProgramData%\AreportFile 下的同名文件，而 Excel / WPS 打开
# xlsx 时会独占它：保存对话框点「保存」会提示文件正在使用 / 一直弹框，
# 新模板就下不下来。所以点「下载表格模板」之前先把这个文件关掉，等锁释放。
def _areportfile_dir():
    """应用数据目录：%ProgramData%\\AreportFile。"""
    return os.path.join(os.environ.get("ProgramData") or r"C:\ProgramData", "AreportFile")


def _template_xlsx_paths():
    """AreportFile 下现存的模板 xlsx（排除 Excel 的 ~$ 临时文件）。"""
    dir_path = _areportfile_dir()
    try:
        names = sorted(os.listdir(dir_path))
    except OSError:
        return []
    return [os.path.join(dir_path, n) for n in names
            if n.lower().endswith(".xlsx") and not n.startswith("~$")]


def _template_windows(paths):
    """查找打开了这些模板 xlsx 的窗口，返回 {hwnd: (hwnd, cls, title, exe)}。"""
    wins = {}
    for p in paths:
        base = os.path.splitext(os.path.basename(p))[0]
        for w in _win_find_doc_windows(base, office_only=True):
            wins[w[0]] = w
    if not wins:
        # 被其它表格软件打开（进程名不在已知列表里）时，放宽到任意窗口按标题匹配
        for p in paths:
            base = os.path.splitext(os.path.basename(p))[0]
            for w in _win_find_doc_windows(base, office_only=False):
                wins[w[0]] = w
    return wins


def _close_open_template_xlsx():
    """关闭用户已经打开的模板 xlsx（Excel / WPS 等），释放文件锁。

    关闭策略（按温和到暴力）：
      1) PostMessage WM_CLOSE 每个命中的窗口 + 处理「是否保存」弹窗；
      2) 等待窗口消失 + 全部文件锁释放；
      3) 仍有文件被锁 → 对每个残留文件跑进程级定点 terminate 兜底；
      4) 兜底后仍有锁则返回 False，否则返回 True。

    返回 True 表示调用后已无残留占用（或本来就没打开）；
    返回 False 表示文件仍被占用，调用方决定是否继续。
    """
    if not _WIN32 or _user32 is None:
        print("[*] 非 Windows 环境，跳过「关闭已打开的模板 xlsx」检查。")
        return True

    paths = _template_xlsx_paths()
    if not paths:
        print("[*] %s 下暂无 xlsx，无需关闭。" % _areportfile_dir())
        return True

    names = "、".join(os.path.basename(p) for p in paths)
    locked = [p for p in paths if _win_file_locked(p)]
    wins = _template_windows(paths)

    if not wins and not locked:
        print("[*] 模板 xlsx 未被打开（%s），直接下载新模板。" % names)
        return True

    # 文件被占用但没找到窗口：刚双击打开时标题可能还没出来，稍等一下再看
    if locked and not wins:
        deadline = time.time() + 5
        while time.time() < deadline and not wins:
            time.sleep(0.5)
            wins = _template_windows(paths)

    if not wins:
        # 没有窗口但文件被锁 → 直接进程级兜底
        print("[!] 模板 xlsx 被占用，但没找到对应的表格窗口（可能被后台进程占用）"
              + " → 进程级兜底关闭 ...")
        any_ok = True
        for p in locked:
            if not _kill_office_procs_for_file(p, wait_timeout=15):
                print("[!] 进程级兜底仍未释放: " + p)
                any_ok = False
            else:
                print("[+] 进程级兜底释放: " + p)
        return any_ok

    for hwnd, _cls, title, exe in wins.values():
        print("[*] 检测到模板 xlsx 已被打开，先关闭: " + title
              + "（" + (exe or "未知程序") + "）")
        _win_close_window(hwnd)

    # 等窗口真正关掉（期间处理「是否保存修改」弹窗 → 点「不保存」）
    deadline = time.time() + CLOSE_XLSX_WAIT
    while time.time() < deadline:
        _win_dismiss_prompts()
        if not any(_win_window_exists(h) for h in wins):
            break
        time.sleep(0.5)

    # 等文件锁释放（窗口关掉后表格进程退出才会解锁）
    deadline = time.time() + CLOSE_XLSX_WAIT
    while time.time() < deadline:
        still = [p for p in paths if _win_file_locked(p)]
        if not still:
            print("[+] 模板 xlsx 已关闭、文件锁已释放，可以下载新模板。")
            return True
        _win_dismiss_prompts()
        time.sleep(0.5)

    # WM_CLOSE 走完仍有锁 → 对每个残留文件跑进程级兜底
    still = [p for p in paths if _win_file_locked(p)]
    print("[!] 窗口关闭后仍有 %d 个模板被锁 → 进程级强制关闭兜底 ..." % len(still))
    any_ok = True
    for p in still:
        if not _kill_office_procs_for_file(p, wait_timeout=15):
            print("[!] 进程级兜底仍未释放: " + p)
            any_ok = False
        else:
            print("[+] 进程级兜底释放: " + p)

    if any_ok:
        print("[+] 所有模板 xlsx 已关闭、文件锁已释放，可以下载新模板。")
    else:
        still_final = [p for p in paths if _win_file_locked(p)]
        print("[!] 仍有模板 xlsx 被占用，新模板可能无法覆盖: " + "、".join(still_final))
    return any_ok


def _step_click_download_template():
    """第一步：进入首页后自动点击「下载表格模板」按钮。

    点击前先检查模板 xlsx（%ProgramData%\\AreportFile\\数据导入模板.xlsx 等）
    是否已被 Excel / WPS 打开：打开状态下文件被独占，新模板写不进去，
    因此先关掉该文件、等锁释放，再点下载。
    """
    assert cdp_client is not None, "cdp_client 未初始化"

    # ---- 先关掉已打开的模板 xlsx，避免新模板因文件占用下载/保存失败 ----
    if not _close_open_template_xlsx():
        print("[!] 模板 xlsx 仍被占用，继续尝试下载；若第 4 步保存失败，请手动关闭该文件后重试。")

    deadline = time.time() + 30
    res = None
    while time.time() < deadline:
        res = cdp_client.eval_js(JS_CLICK_DOWNLOAD_TEMPLATE)
        if res and res.get("ok"):
            break
        time.sleep(0.5)

    if not res or not res.get("ok"):
        print("[!] 点击「下载表格模板」失败: " + str(res))
        return False
    print("[+] 已点击「下载表格模板」按钮（" + str(res.get("text", "")) + "）")
    return True


def _step_handle_save_dialog():
    """第二步：在系统「保存」对话框中点击「保存」按钮。"""
    if not _WIN32 or _user32 is None:
        print("[!] 当前环境不支持原生对话框自动化，请在弹出的窗口中手动点击「保存」。")
        return False

    print("[*] 等待系统保存对话框弹出 ...")
    dlg = 0
    title = ""
    btn_text = ""
    deadline = time.time() + SAVE_DIALOG_TIMEOUT
    while time.time() < deadline:
        dlg, title, btn_text = _win_find_save_dialog()
        if dlg:
            break
        time.sleep(0.5)

    if not dlg:
        print("[!] 等待保存对话框超时（%d 秒），模板下载流程可能未触发。" % SAVE_DIALOG_TIMEOUT)
        return False

    fname, _edit = _win_get_filename_edit(dlg)
    folder = _win_get_current_folder(dlg)
    print("[+] 已定位保存对话框：标题=" + title + "，文件名=" + (fname or "(空)")
          + "，目录=" + (folder or "(未知)") + "，按钮=" + btn_text)

    # 先记录日志位置与时间，便于之后精确定位保存出来的文件
    snap = _app_log_snapshot()
    t_click = time.time()

    # ---- 强制保存到 %ProgramData%\AreportFile ----
    # 原生对话框会记住「上次使用的目录」：上一轮流程末尾用文件夹选择框
    # 选中过 OrigDoc，本轮保存对话框就可能默认到 OrigDoc，导致 xlsx 落进
    # AreportFile\OrigDoc（还会被当成报告原件一起导入）。这里显式纠正保存目录：
    # 第 1 轮写「完整路径」直接指定保存位置；若外壳拒绝（沿用打开对话框的实测
    # 结论：完整路径可能被规范化成正斜杠报「文件名无效」），第 2 轮改写纯文件名，
    # 此时目录一般已被第 1 轮跳正。
    want_dir = os.path.join(os.environ.get("ProgramData") or r"C:\ProgramData", "AreportFile")
    try:
        os.makedirs(want_dir, exist_ok=True)
    except OSError as e:
        print("[!] 创建目标目录失败: " + repr(e))
    base_name = os.path.basename(fname) if fname else ""
    plan = [os.path.join(want_dir, base_name), base_name] if base_name else [""]

    # ---- 方案A：保存前若目标文件正被 Excel/WPS 等打开（独占写锁），
    #      先关掉占用它的窗口（弹「是否保存」时点「不保存」），释放锁后再保存 ----
    if base_name:
        target_path = os.path.join(want_dir, base_name)
        if os.path.isfile(target_path) and _win_file_locked(target_path):
            base = os.path.splitext(base_name)[0]
            print("[*] 目标文件正被其它程序占用，尝试关闭占用窗口: " + target_path)
            wins = _win_find_doc_windows(base, office_only=True)
            if not wins:
                wins = _win_find_doc_windows(base, office_only=False)
            if not wins:
                print("[!] 未找到占用该文件的窗口（可能被后台进程占用），跳过关闭。")
            for hwnd, _cls, title, exe in wins:
                print("[*] 关闭占用窗口: " + title + "（" + (exe or "未知程序") + "）")
                _win_close_window(hwnd)
            lock_deadline = time.time() + CLOSE_XLSX_WAIT
            while wins and time.time() < lock_deadline:
                _win_dismiss_prompts()  # 「是否保存修改」弹窗 → 点「不保存」
                if not _win_file_locked(target_path):
                    print("[+] 占用已释放，可正常覆盖保存。")
                    break
                time.sleep(0.5)
            else:
                if wins:
                    print("[!] 等待文件锁释放超时，继续尝试保存。")

    closed = False
    for attempt in range(1, 4):
        _win_dismiss_error_boxes()

        # 每轮重新取控件句柄（上一次点击可能让外壳跳了目录 / 弹过错误框）
        dlg_now, _t, _b = _win_find_save_dialog()
        if not dlg_now:
            closed = True
            break
        dlg = dlg_now
        _cur_txt, edit = _win_get_filename_edit(dlg)
        btn, _ = _win_find_button(dlg, ("保存", "Save"))
        if not btn:
            print("[!] 未找到「保存」按钮")
            return False

        use_text = plan[min(attempt - 1, len(plan) - 1)]
        if use_text and edit:
            _win_set_text(edit, use_text)
            time.sleep(0.3)

        print("[*] 第 %d 次点击「保存」（目标目录: %s，写入: %s） ..."
              % (attempt, want_dir, use_text or "(沿用对话框内容)"))
        _win_click_button(btn)

        # 等待对话框关闭；若出现「文件已存在，是否替换」确认框则点「是」
        deadline = time.time() + max(5, SAVE_CLICK_TIMEOUT // 2)
        while time.time() < deadline:
            cdlg, ctitle, cbtn = _win_find_confirm_dialog()
            if cdlg:
                print("[*] 弹出确认框「" + ctitle + "」，点击「" + cbtn + "」")
                yes_btn, _ = _win_find_button(cdlg, ("是", "替换", "Yes"))
                _win_click_button(yes_btn)
                time.sleep(0.5)
                continue
            if not _win_find_save_dialog()[0]:
                closed = True
                break
            time.sleep(0.3)
        if closed:
            break
        print("[!] 点击后保存对话框仍未关闭，准备重试 ...")

    if not closed:
        _win_dismiss_error_boxes()
        print("[!] 多次点击「保存」后对话框仍未关闭。")
        return False

    print("[+] 保存对话框已关闭，等待应用写入文件 ...")
    expected = ""
    if base_name:
        # 已把保存目录纠正为 AreportFile，预期路径以它为准
        expected = os.path.join(want_dir, base_name)
    elif folder and fname:
        expected = os.path.join(folder, os.path.basename(fname))
    DOWNLOAD_CONTEXT["filename"] = fname
    DOWNLOAD_CONTEXT["folder"] = want_dir if base_name else folder
    DOWNLOAD_CONTEXT["expected_path"] = expected
    DOWNLOAD_CONTEXT["since"] = t_click
    DOWNLOAD_CONTEXT["log_snapshot"] = snap
    return True


def _describe_excel(path):
    """用 openpyxl 读一下表头，方便确认模板内容（未安装则跳过）。"""
    try:
        import openpyxl
    except Exception:
        print("[*] 未安装 openpyxl，跳过模板内容检查。")
        return
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        names = list(wb.sheetnames)
        print("[*] 模板工作表: " + ", ".join(str(n) for n in names))
        if names:
            ws = wb[names[0]]
            for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
                cells = [str(c) for c in row if c is not None]
                if cells:
                    print("[*] 首行表头: " + " | ".join(cells[:12]))
                break
        wb.close()
    except Exception as e:
        print("[*] 读取模板内容失败（不影响后续）: " + repr(e))


def _locate_template_file():
    """找到刚刚保存下来的 .xlsx 文件，返回路径（找不到返回 None）。"""
    fname = DOWNLOAD_CONTEXT.get("filename", "")
    since = DOWNLOAD_CONTEXT.get("since", time.time())
    snap = DOWNLOAD_CONTEXT.get("log_snapshot", {})
    expected = DOWNLOAD_CONTEXT.get("expected_path", "")

    path = None
    # 1) 对话框里读到的完整路径（最直接）
    if expected:
        deadline = time.time() + TEMPLATE_SAVE_TIMEOUT
        while time.time() < deadline:
            if os.path.isfile(expected):
                path = expected
                break
            # 同时看一眼应用日志：保存位置与预期不一致时
            # （例如被对话框记住的上次目录带偏）不必空等满超时
            for p in _extract_save_paths(_read_log_delta(snap)):
                if p.lower().endswith(".xlsx") and os.path.isfile(p):
                    path = p
                    break
            if path:
                break
            time.sleep(0.5)
        if path:
            print("[+] 已在保存位置找到文件: " + path)

    # 2) 从应用日志里解析 savePath
    if not path:
        deadline = time.time() + TEMPLATE_SAVE_TIMEOUT
        while time.time() < deadline:
            for p in _extract_save_paths(_read_log_delta(snap)):
                if p.lower().endswith(".xlsx") and os.path.isfile(p):
                    path = p
                    break
            if path:
                break
            time.sleep(0.5)
        if path:
            print("[+] 已从应用日志中解析到保存路径: " + path)

    # 3) 兜底：在候选目录中扫描
    if not path:
        path = _scan_for_xlsx(fname, since)
        if path:
            print("[*] 前面都没定位到，已在候选目录中扫描到文件: " + path)

    if not path:
        print("[!] 未能找到刚刚保存的 .xlsx 文件" + (("：" + fname) if fname else ""))
        return None

    try:
        st = os.stat(path)
        print("[+] 已找到刚保存的文件: " + path)
        print("    大小=%.1f KB，修改时间=%s" % (
            st.st_size / 1024.0,
            time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(st.st_mtime))))
    except OSError as e:
        print("[!] 读取文件信息失败: " + repr(e))

    _describe_excel(path)
    return path


def _move_template_into_areportfile(path):
    """兜底：若 xlsx 不在 %ProgramData%\\AreportFile 直下，把它搬回去。

    原生对话框会记住上次使用的目录（上一轮选过 OrigDoc 后，
    本轮 xlsx 可能落进 AreportFile\\OrigDoc）。xlsx 混在 OrigDoc 里
    会被应用当成「报告原件」一并导入，必须搬回 AreportFile。
    """
    if not path:
        return path
    program_data = os.environ.get("ProgramData") or r"C:\ProgramData"
    want_dir = os.path.join(program_data, "AreportFile")
    try:
        cur_dir = os.path.dirname(os.path.abspath(path))
    except Exception:
        return path
    if os.path.normcase(cur_dir) == os.path.normcase(os.path.abspath(want_dir)):
        return path
    target = os.path.join(want_dir, os.path.basename(path))
    try:
        os.makedirs(want_dir, exist_ok=True)
        os.replace(path, target)  # 同盘内为原子移动，目标已存在则覆盖
    except OSError as e:
        print("[!] 搬回 xlsx 到 %s 失败（沿用原路径）: %s" % (want_dir, repr(e)))
        return path
    print("[*] xlsx 原本落在 %s，已搬回: %s" % (cur_dir, target))
    return target


def _step_locate_template_file():
    """第三步：定位刚刚保存下来的 .xlsx 文件（供后续填写 / 打开使用）。"""
    path = _locate_template_file()
    if not path:
        return False
    path = _move_template_into_areportfile(path)
    DOWNLOAD_CONTEXT["path"] = path
    return True


# ===================== 模板 xlsx 自动填写数据 =====================
# 数据来源：先 POST /api/Login/CheckLogin 登录取得 res.data.token，
# 再带 Authorization 头 POST HashReport 接口（GetHashReportUploadedList，Status=1）获取报告列表，
# 每条数据按 ft / lt / qj 三种报告类型展开为 0~3 行写入 xlsx：
#   某类型的 {前缀}UploadStatus == 1（要上传）时，该类型生成一行；
#   三个状态都是 1 则一条数据生成三行，都不是 1 则跳过该条。
#
# 字段映射（以前缀 ft 为例，lt / qj 同理，仅字段前缀不同）：
#   报告编号      -> ftReportNo
#   授权签字人    -> ftReportQF
#   报告签发日期  -> ftReportDate（取日期部分 YYYY-MM-DD）
#   报告签发年份  -> ftReportDate 截取年份
#   哈希值        -> 留空（系统自动生成）
#   编号是否有效  -> 默认「是」
#   委托人        -> ftReportWT

# HashReport 接口配置（域名后续可改，暂用 http://ba.gdyjjc.com:2037/）
HASH_REPORT_BASE_URL = "http://ba.gdyjjc.com:2037/"
# 登录：POST + {"UserName": ..., "Password": ...} → res.data.token（登录成功 code=0）
HASH_REPORT_LOGIN_API = HASH_REPORT_BASE_URL + "/api/Login/CheckLogin"
# 列表：POST + Authorization: <token>（后端校验 header 原值，不能加 "Bearer " 前缀）
HASH_REPORT_API = HASH_REPORT_BASE_URL + "/api/HashReport/GetHashReportUploadedList"
# 全部流程完成后回写上传状态：POST + Authorization，只把 1（要上传）改为 2（已上传）
HASH_REPORT_MARK_UPLOADED_API = (
    HASH_REPORT_BASE_URL.rstrip("/") + "/api/HashReport/MarkHashReportUploaded"
)
# 默认账号密码；可在 gui_config.json 里用同名键覆盖（重启生效）
HASH_REPORT_USERNAME = "管理员2"
HASH_REPORT_PASSWORD = "rpa2025"
HASH_REPORT_BEGIN_DATE = ""   # 留空 = 不限起始日期，后续需要再填
HASH_REPORT_END_DATE = ""     # 留空 = 不限结束日期，后续需要再填
# 列表状态口径：1=要上传（条件齐备、等待上传），2=已上传
HASH_REPORT_STATUS = 1

# 报告 PDF 的下载域名（与 HashReport 接口同一台服务器）。
# 完整地址规则：
#   {域名}/Server/Upload/Formal/TestCheckOrder/{T_AutoID}/UploadPic/{ltReport}
REPORT_PDF_BASE_URL = HASH_REPORT_BASE_URL

# 模板填写相关常量
FILL_ORG = "广东优捷检测科技有限公司"  # 写入 C2（检验检测机构名称）
FILL_CLEAR_FROM_ROW = 4               # 从该行开始清空原有数据
FILL_START_ROW = 4                    # 数据写入起始行


# 三种报告类型的字段前缀；{前缀}UploadStatus == 1（要上传）表示该类报告需要填入模板
UPLOAD_FIELD_PREFIXES = ("ft", "lt", "qj")

# ---- 大批量流式处理配置 ----
# 目标：20 万条数据只产出「一个」全量 xlsx，但内存占用不再随总条数增长。
# 做法：接口逐页拉取 → 展开成扁平行写入磁盘 JSONL → 写表/PDF 下载各自流式读。
HASH_REPORT_PAGE_ROW = 1000        # 每页条数
HASH_REPORT_MAX_PAGES = 500        # 分页硬上限，防止后端不翻页导致死循环
ROWS_FILE_NAME = "auto_upload_rows.jsonl"   # 展开后的中转文件（放在系统临时目录）
PDF_DOWNLOAD_WORKERS = 12          # PDF 下载并发线程数
PDF_MAX_INFLIGHT = 32              # 内存中最多同时存在的待下载任务数
PDF_FAIL_TOLERANCE = 0             # 0 = 零容忍：任一份 PDF 下载失败即终止流程
LARGE_DATA_ROW_THRESHOLD = 10000  # 超过该行数时跳过「打开/关闭 xlsx」预览


# ===================== 填写 xlsx 前：按「已上传报告管理」列表去重 =====================
# 把接口拉到的全量数据，先与系统里「已上传报告管理 → 全部报告」列表逐页比对：
# 列表里已经出现过的报告编号不再写入模板 xlsx（PDF 下载复用同一份中转数据，
# 所以也会一并跳过）。该列表有很多分页，必须翻到最后一页才能拿到完整编号集合。
def _optional_config(name, default):
    """读取 gui_config.json 里的可选开关（不改动主配置结构与设置界面）。

    主配置 load_config() 只认识 DEFAULT_CONFIG 里的键，这里单独读一次，
    让用户可以直接在 gui_config.json 里加 FILTER_UPLOADED_REPORTS /
    STRICT_UPLOADED_FILTER 两个布尔值（改完重启程序生效）。
    """
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        v = data.get(name) if isinstance(data, dict) else None
        if isinstance(v, bool):
            return v
        if isinstance(v, str):
            return v.strip().lower() not in ("", "0", "false", "no", "off")
    except (OSError, ValueError):
        pass
    return default


FILTER_UPLOADED_REPORTS = _optional_config("FILTER_UPLOADED_REPORTS", True)
STRICT_UPLOADED_FILTER = _optional_config("STRICT_UPLOADED_FILTER", True)
UPLOADED_PAGE_MENU_TEXT = "已上传报告管理"
UPLOADED_PAGE_TAB_TEXT = "全部报告"
UPLOAD_PAGE_MENU_TEXT = "报告编号上传"
UPLOADED_REPORT_NO_COLUMN = "报告编号"
UPLOADED_LIST_MAX_PAGES = 5000      # 翻页硬上限，防止分页状态异常导致死循环
UPLOADED_LIST_PAGE_WAIT = 12        # 切换菜单 / tab / 翻页后等待列表渲染的最长秒数
UPLOADED_LIST_LOG_EVERY = 10        # 每多少页回报一次采集进度
UPLOADED_REPORT_NOS = set()         # 本轮采集到的已上传报告编号（规范化后，供写表时过滤）

# 列表单元格里的占位文本，不能当作报告编号
_UPLOADED_NOISE = {"", "-", "--", "暂无数据", "无数据", "NODATA", "EMPTY",
                   "报告编号", "报告编号:", "报告编号："}


def _norm_report_no(value):
    """报告编号规范化：去所有空白、统一大小写，便于页面列表与接口数据比对。"""
    return re.sub(r"\s+", "", str(value or "")).upper()


# 按文本定位菜单 / tab，返回坐标交给 CDP 做真实点击（__KEY__ / __SCOPE__ 由 Python 填入）
JS_LOCATE_MENU_OR_TAB = r"""
(function () {
    const KEY = __KEY__;
    const SCOPE = __SCOPE__;
    function norm(s) { return ((s || '') + '').replace(/\s+/g, ''); }
    function textOf(el) { return norm(el.innerText || el.textContent || ''); }
    function visible(el) { const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0; }
    function depth(el) { let d = 0, p = el.parentElement; while (p && d < 60) { d++; p = p.parentElement; } return d; }
    const K = norm(KEY);
    const MENU_SEL = '[role="menuitem"], [class*="menu-item"], [class*="menuItem"], aside a, nav a, .arco-menu a, [class*="sidebar"] a, [class*="sider"] a';
    const TAB_SEL = '[role="tab"], [class*="tabs-tab"], [class*="tabs-header-title"], [class*="tab-item"], [class*="tabItem"]';
    const ANY_SEL = 'a, button, li, span, div';
    let groups = [MENU_SEL, TAB_SEL, ANY_SEL];
    if (SCOPE === 'tab') groups = [TAB_SEL, MENU_SEL, ANY_SEL];

    function pick(pool) {
        const exact = pool.filter(e => textOf(e) === K);
        if (exact.length) return exact;
        return pool.filter(e => {
            const t = textOf(e);
            return t && t.indexOf(K) >= 0 && t.length <= K.length + 16;
        });
    }
    let hits = [];
    for (const sel of groups) {
        hits = pick(Array.from(document.querySelectorAll(sel)).filter(visible));
        if (hits.length) break;
    }
    if (!hits.length) return { ok: false, msg: '未找到「' + KEY + '」' };
    hits.sort((a, b) => depth(b) - depth(a));   // 取最内层文本元素，避免点到整个容器
    const el = hits[0];

    // 向上找到真正承载点击事件的祖先（菜单项 / tab / 链接 / 按钮）
    let target = el;
    for (let p = el, i = 0; p && i < 4; p = p.parentElement, i++) {
        const cls = (p.className || '') + '';
        if (/^(A|LI|BUTTON)$/i.test(p.tagName) || p.getAttribute('role') === 'menuitem' ||
            p.getAttribute('role') === 'tab' || /menu-item|tabs-tab|tab-item/i.test(cls)) {
            target = p;
            break;
        }
    }
    try { target.scrollIntoView({ block: 'center', inline: 'center' }); } catch (e) {}
    const r = target.getBoundingClientRect();
    if (!r.width || !r.height) return { ok: false, msg: '「' + KEY + '」不可见' };
    return {
        ok: true,
        x: Math.round(r.x + r.width / 2),
        y: Math.round(r.y + r.height / 2),
        tag: target.tagName,
        cls: (target.className || '') + '',
        text: textOf(target)
    };
})()
"""

# 读取「全部报告」列表的「报告编号」列 + 当前分页状态（一次求值拿全，减少 CDP 往返）
JS_READ_UPLOADED_LIST = r"""
(function () {
    const KEY = __KEY__;
    function norm(s) { return ((s || '') + '').replace(/\s+/g, ''); }
    function tOf(el) { return ((el.innerText || el.textContent || '') + '').trim(); }
    const K = norm(KEY);
    const TABLE_SEL = 'table, .arco-table, [role="table"]';
    const TH_SEL = 'th, [role="columnheader"], [class*="table-th"]';
    const TR_SEL = 'tbody tr, [role="row"], [class*="table-tr"]';
    const TD_SEL = 'td, [role="cell"], [class*="table-td"]';

    // 1) 全页找「报告编号」列索引：表头与表体可能被拆成两个 table，分开找更稳
    const tables = Array.from(document.querySelectorAll(TABLE_SEL));
    let colIdx = -1, headCount = 0, headers = [];
    for (const t of tables) {
        const heads = Array.from(t.querySelectorAll(TH_SEL))
            .map(h => norm(h.innerText || h.textContent || ''));
        let idx = heads.findIndex(h => h === K);
        if (idx < 0) idx = heads.findIndex(h => h && h.indexOf(K) >= 0);
        if (idx >= 0 && heads.length > headCount) {
            colIdx = idx; headCount = heads.length; headers = heads;
        }
    }
    if (colIdx < 0) return { ok: false, msg: '未找到含「' + KEY + '」列的表格' };

    // 2) 取数据行最多的表格（固定列 / 表头表体分离时页面上会同时存在多张 table）
    let best = null;
    for (const t of tables) {
        const bodyRows = Array.from(t.querySelectorAll(TR_SEL))
            .filter(tr => tr.querySelectorAll(TD_SEL).length > 0);
        if (!best || bodyRows.length > best.rowCount) {
            best = { bodyRows: bodyRows, rowCount: bodyRows.length, table: t };
        }
    }
    const bodyRows = best ? best.bodyRows : [];
    if (!bodyRows.length) {
        // 「暂无数据」是合法的空列表；其它情况说明还在异步加载，交给调用方继续等
        const anyEmpty = tables.some(t => /暂无数据|无数据|NoData/i.test(tOf(t)));
        if (!anyEmpty) return { ok: false, msg: '「' + KEY + '」列表数据行尚未渲染' };
    }

    const values = [];
    for (const tr of bodyRows) {
        const cells = Array.from(tr.querySelectorAll(TD_SEL));
        const cell = cells[colIdx];
        if (!cell) { values.push(''); continue; }
        // 内容过长时 Arco 会把完整值放在 title 里，两个都收
        const titled = cell.querySelector('[title]');
        const title = cell.getAttribute('title') || (titled ? titled.getAttribute('title') : '') || '';
        values.push((title ? title + '\n' : '') + tOf(cell));
    }
    const first = (values.find(v => v && v.trim()) || '').slice(0, 80);
    const last = values.length ? (values[values.length - 1] || '').slice(0, 80) : '';
    const tableRaw = best ? tOf(best.table) : '';
    const empty = /暂无数据|无数据|NoData/i.test(tableRaw) ||
                  /暂无数据|无数据|NoData/i.test(document.body ? (document.body.innerText || '') : '') ||
                  !values.some(v => v && v.trim());

    // ---- 分页状态 ----
    const pager = document.querySelector('.arco-pagination, .ant-pagination, [class*="pagination"], [class*="pager"]');
    const pinfo = { hasPager: !!pager, next: null, current: '', totalPages: null, buttons: [] };
    if (pager) {
        const act = pager.querySelector('.arco-pagination-item-active, .ant-pagination-item-active, [aria-current="page"], [class*="is-active"]');
        if (act) pinfo.current = norm(tOf(act));
        const m = tOf(pager).match(/共\s*(\d+)\s*页/);
        if (m) pinfo.totalPages = parseInt(m[1], 10);

        const NEXT_SEL = '.arco-pagination-item-next, .arco-pagination-next, .ant-pagination-next, [title="下一页"], [aria-label="下一页"], [aria-label="Next Page"], li[class*="next"], button[class*="next"]';
        let n = pager.querySelector(NEXT_SEL);
        if (!n) {
            n = Array.from(pager.querySelectorAll('li, button, a')).find(e => {
                const s = norm(tOf(e));
                return s === '下一页' || s === '>' || s === '›' || s === '»' || s === 'Next';
            });
        }
        if (n) {
            const r = n.getBoundingClientRect();
            const cls = (n.className || '') + '';
            const pcls = (n.parentElement && n.parentElement.className || '') + '';
            pinfo.next = {
                x: Math.round(r.x + r.width / 2),
                y: Math.round(r.y + r.height / 2),
                disabled: n.disabled === true || n.getAttribute('aria-disabled') === 'true' ||
                          /disabled/i.test(cls) || /disabled/i.test(pcls)
            };
        }
        // 兜底：某些分页组件没有「下一页」按钮，只能点页码
        const items = Array.from(pager.querySelectorAll('li, button, a'))
            .filter(e => /^\d+$/.test(norm(tOf(e))) && norm(tOf(e)).length <= 6);
        for (const e of items.slice(0, 40)) {
            const r = e.getBoundingClientRect();
            const cls = (e.className || '') + '';
            pinfo.buttons.push({
                text: norm(tOf(e)),
                x: Math.round(r.x + r.width / 2),
                y: Math.round(r.y + r.height / 2),
                active: /active/i.test(cls) || e.getAttribute('aria-current') === 'page'
            });
        }
    }
    return {
        ok: true,
        col: colIdx,
        headers: headers,
        rowCount: bodyRows.length,
        values: values,
        empty: empty,
        fingerprint: (pinfo.current || '') + '|' + bodyRows.length + '|' + first + '|' + last,
        pager: pinfo
    };
})()
"""

# 切回上传菜单后的确认：上传页特有的两个按钮文案
JS_CHECK_UPLOAD_PAGE = r"""
(function () {
    const t = document.body ? (document.body.innerText || '') : '';
    const hits = ['下载表格模板', '选择数据表格导入'].filter(k => t.indexOf(k) >= 0);
    return { ok: hits.length > 0, hits: hits, url: location.href };
})()
"""


def _js_locate_text(key, scope):
    return (JS_LOCATE_MENU_OR_TAB
            .replace("__KEY__", json.dumps(key, ensure_ascii=False))
            .replace("__SCOPE__", json.dumps(scope)))


def _js_read_uploaded_list():
    return JS_READ_UPLOADED_LIST.replace(
        "__KEY__", json.dumps(UPLOADED_REPORT_NO_COLUMN, ensure_ascii=False))


def _eval_js_retry(js, timeout, require_ok=True):
    """页面切换 / 异步加载期间反复求值，直到拿到有效结果或超时。"""
    assert cdp_client is not None, "cdp_client 未初始化"
    deadline = time.time() + timeout
    last = None
    while time.time() < deadline:
        try:
            last = cdp_client.eval_js(js)
        except Exception:
            last = None      # 渲染上下文重建时会抛异常，重试即可
        if isinstance(last, dict) and (not require_ok or last.get("ok")):
            return last
        time.sleep(0.35)
    return last if isinstance(last, dict) else None


def _click_text_element(key, scope, timeout=UPLOADED_LIST_PAGE_WAIT):
    """按文本定位菜单 / tab，再用 CDP 做一次真实点击（trusted click）。"""
    res = _eval_js_retry(_js_locate_text(key, scope), timeout)
    if not res or not res.get("ok"):
        print("[!] 未找到可点击的「%s」: %s" % (key, (res or {}).get("msg", "超时")))
        return False
    if not _cdp_click_at(cdp_client, int(res["x"]), int(res["y"])):
        return False
    print("[*] 已点击「%s」（%s.%s）"
          % (key, res.get("tag", ""), str(res.get("cls", ""))[:40]))
    return True


def _add_uploaded_values(values):
    """把「报告编号」列单元格文本拆行、规范化后并入全局编号集合。"""
    added = 0

    def add(n):
        nonlocal added
        if not n or n in _UPLOADED_NOISE:
            return
        if n not in UPLOADED_REPORT_NOS:
            added += 1
        UPLOADED_REPORT_NOS.add(n)

    for raw in values or []:
        line = str(raw or "")
        add(_norm_report_no(line))
        for piece in str(raw or "").splitlines():
            add(_norm_report_no(piece))
            # 兼容单元格里写成「报告编号：XXX」的情况
            if "：" in piece or ":" in piece:
                add(_norm_report_no(re.split(r"[：:]", piece)[-1]))
    return added


def _wait_uploaded_list_page_change(prev_fp, timeout=UPLOADED_LIST_PAGE_WAIT):
    """翻页后等表格真正刷新：连续两轮读到相同的新指纹才认为渲染稳定。"""
    assert cdp_client is not None, "cdp_client 未初始化"
    deadline = time.time() + timeout
    last_fp = None
    while time.time() < deadline:
        time.sleep(0.25)
        try:
            st = cdp_client.eval_js(_js_read_uploaded_list())
        except Exception:
            st = None
        if not st or not st.get("ok"):
            continue
        fp = st.get("fingerprint") or ""
        if not fp or fp == prev_fp:
            last_fp = None
            continue
        if fp == last_fp:
            return st
        last_fp = fp
    return None


def _switch_back_to_upload_menu():
    """切回「报告编号上传」菜单，并确认上传页按钮已重新渲染。"""
    for attempt in range(1, 4):
        if _click_text_element(UPLOAD_PAGE_MENU_TEXT, "menu", timeout=6):
            res = _eval_js_retry(JS_CHECK_UPLOAD_PAGE, 8)
            if res and res.get("ok"):
                print("[+] 已切回「%s」菜单（%s）"
                      % (UPLOAD_PAGE_MENU_TEXT, "、".join(res.get("hits") or [])))
                return True
        print("[*] 第 %d 次切回「%s」未确认到上传页，重试 ..."
              % (attempt, UPLOAD_PAGE_MENU_TEXT))
        time.sleep(0.8)
    return False


def _step_collect_uploaded_report_nos():
    """填写 xlsx 前：进入「已上传报告管理 → 全部报告」，翻遍所有分页采集报告编号。

    结果写入全局 UPLOADED_REPORT_NOS，写表步骤据此过滤已上传数据。
    无论采集成功与否，最后都必须切回「报告编号上传」菜单，否则后续导入/上传步骤会失效。
    """
    if not FILTER_UPLOADED_REPORTS:
        UPLOADED_REPORT_NOS.clear()
        print("[*] 已关闭「按已上传报告列表去重」，跳过页面采集。")
        return True

    assert cdp_client is not None, "cdp_client 未初始化"
    UPLOADED_REPORT_NOS.clear()
    # 采集结果必须参与本轮写表，避免复用上一轮残留的中转数据
    DOWNLOAD_CONTEXT.pop("rows_file", None)
    err = None
    page_no = 1
    try:
        print("[*] 切换到「%s」页面 ..." % UPLOADED_PAGE_MENU_TEXT)
        if not _click_text_element(UPLOADED_PAGE_MENU_TEXT, "menu"):
            raise RuntimeError("未能点击「%s」菜单" % UPLOADED_PAGE_MENU_TEXT)
        time.sleep(0.8)
        print("[*] 切换到「%s」tab ..." % UPLOADED_PAGE_TAB_TEXT)
        if not _click_text_element(UPLOADED_PAGE_TAB_TEXT, "tab"):
            raise RuntimeError("未能点击「%s」tab" % UPLOADED_PAGE_TAB_TEXT)

        state = _eval_js_retry(_js_read_uploaded_list(), UPLOADED_LIST_PAGE_WAIT)
        if not state or not state.get("ok"):
            raise RuntimeError("「%s」列表未渲染出含「%s」列的表格: %s" % (
                UPLOADED_PAGE_TAB_TEXT, UPLOADED_REPORT_NO_COLUMN,
                (state or {}).get("msg", "超时")))

        seen_fp = set()
        while True:
            fp = state.get("fingerprint")
            if fp in seen_fp:
                raise RuntimeError("第 %d 页内容与之前某页重复，疑似分页状态异常" % page_no)
            seen_fp.add(fp)

            added = _add_uploaded_values(state.get("values"))
            pager = state.get("pager") or {}
            if page_no == 1 or page_no % UPLOADED_LIST_LOG_EVERY == 0 or state.get("empty"):
                page_hint = ("（空列表）" if state.get("empty") else
                             "（第 %s/%s 页）" % (pager.get("current") or "?",
                                                pager.get("totalPages") or "?"))
                print("[*] 已上传报告列表 第 %d 页：%d 行，新增编号 %d 个，累计 %d 个%s"
                      % (page_no, state.get("rowCount", 0), added,
                         len(UPLOADED_REPORT_NOS), page_hint))

            # 分页显示「共 N 页」时，当前页已是最后一页就直接收尾，不再多点一次
            cur = pager.get("current") or ""
            tp = pager.get("totalPages")
            if tp and cur.isdigit() and int(cur) >= int(tp):
                break

            # 找「下一页」可点击的位置；没有下一页按钮时退化为点 active 页码的后一个页码
            nxt = pager.get("next")
            point = None
            if nxt and not nxt.get("disabled"):
                point = (int(nxt["x"]), int(nxt["y"]))
            elif not nxt:
                buttons = pager.get("buttons") or []
                idx = next((i for i, b in enumerate(buttons) if b.get("active")), -1)
                if idx >= 0 and idx + 1 < len(buttons):
                    b = buttons[idx + 1]
                    point = (int(b["x"]), int(b["y"]))

            if point is None:
                break   # 没有分页组件，或已到最后一页
            if page_no >= UPLOADED_LIST_MAX_PAGES:
                raise RuntimeError("分页超过 %d 页上限，请调大 UPLOADED_LIST_MAX_PAGES"
                                   % UPLOADED_LIST_MAX_PAGES)
            if not _cdp_click_at(cdp_client, point[0], point[1]):
                raise RuntimeError("第 %d 页点击「下一页」失败" % page_no)

            new_state = _wait_uploaded_list_page_change(fp, UPLOADED_LIST_PAGE_WAIT)
            if new_state is None:
                # 按钮可能其实已禁用（点击无效）：复核一次分页状态再决定是否结束
                check = _eval_js_retry(_js_read_uploaded_list(), 3)
                ck = ((check or {}).get("pager") or {}).get("next")
                if not ck or ck.get("disabled"):
                    break
                raise RuntimeError("第 %d 页点击「下一页」后列表内容未变化" % page_no)

            state = new_state
            page_no += 1

        print("[+] 已上传报告列表采集完成：共 %d 页、%d 个不重复报告编号"
              % (page_no, len(UPLOADED_REPORT_NOS)))
    except Exception as e:
        err = e
        print("[!] 采集已上传报告编号失败: " + repr(e))

    # ---- 无论成功失败，都必须切回「报告编号上传」菜单 ----
    if not _switch_back_to_upload_menu():
        print("[!] 未能切回「%s」菜单，后续导入/上传步骤无法执行，终止流程。"
              % UPLOAD_PAGE_MENU_TEXT)
        return False

    if err is not None:
        if STRICT_UPLOADED_FILTER:
            print("[!] 已上传报告编号采集失败，为避免重复上传，终止流程。")
            print("[!] 如需忽略该失败继续（本轮不过滤），可在 gui_config.json 里加 \"STRICT_UPLOADED_FILTER\": false 后重启程序。")
            return False
        print("[!] 忽略采集失败，本轮不做已上传报告过滤（可能重复上传）。")
        UPLOADED_REPORT_NOS.clear()
    return True


def _is_empty_ltqj(value) -> bool:
    """判断字段值是否为空（None / 空字符串 / NaN 均视为空）。"""
    if value is None:
        return True
    if isinstance(value, float) and value != value:  # NaN
        return True
    return str(value).strip() == ""


def _status_is_pending(value) -> bool:
    """判断 UploadStatus 值是否为 1（要上传），兼容 int / float / str 类型。"""
    if value is None:
        return False
    s = str(value).strip()
    return s == "1" or s == "1.0"


def _item_has_upload_report(item) -> bool:
    """该条数据是否至少有一类报告（ft / lt / qj）的 UploadStatus 为 1（要上传）。"""
    if not isinstance(item, dict):
        return False
    return any(_status_is_pending(item.get(p + "UploadStatus")) for p in UPLOAD_FIELD_PREFIXES)


# ---- HashReport 登录 / token ----
_HASH_REPORT_TOKEN = {"value": ""}
_HASH_REPORT_TOKEN_LOCK = threading.Lock()


def _optional_str_config(name, default):
    """读取 gui_config.json 里的可选字符串键（不改动主配置结构与设置界面）。"""
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        v = data.get(name) if isinstance(data, dict) else None
        if isinstance(v, str) and v.strip():
            return v.strip()
    except (OSError, ValueError):
        pass
    return default


def _parse_hash_api_payload(resp):
    """统一解析成 dict：登录/鉴权失败时服务端会把 JSON 再包一层字符串。"""
    data = resp.json()
    if isinstance(data, str):
        data = json.loads(data)
    if not isinstance(data, dict):
        raise RuntimeError("接口返回格式异常: %s" % str(data)[:200])
    return data


def _hash_report_login():
    """POST /api/Login/CheckLogin，取 res.data.token。"""
    body = {
        "UserName": _optional_str_config("HASH_REPORT_USERNAME", HASH_REPORT_USERNAME),
        "Password": _optional_str_config("HASH_REPORT_PASSWORD", HASH_REPORT_PASSWORD),
    }
    resp = None
    try:
        resp = requests.post(HASH_REPORT_LOGIN_API, json=body, timeout=30)
        resp.raise_for_status()
        payload = _parse_hash_api_payload(resp)
    except Exception as e:
        raise RuntimeError("HashReport 登录接口请求/解析失败: %s" % repr(e))
    finally:
        if resp is not None:
            try:
                resp.close()
            except Exception:
                pass
    data = payload.get("data")
    token = (data.get("token") if isinstance(data, dict) else None) or ""
    if not token:
        raise RuntimeError("HashReport 登录未取得 token（账号或密码可能有误）: %s"
                           % str(payload)[:200])
    return str(token)


def _hash_report_token(force_refresh=False):
    """同一轮流程复用 token；token 失效时强制重新登录。"""
    with _HASH_REPORT_TOKEN_LOCK:
        if not force_refresh and _HASH_REPORT_TOKEN["value"]:
            return _HASH_REPORT_TOKEN["value"]
        token = _hash_report_login()
        _HASH_REPORT_TOKEN["value"] = token
        print("[+] HashReport 登录成功，已取得 token")
        return token


def _hash_report_body(page):
    """列表接口改为 POST，查询条件放 JSON body；只取「要上传」状态的数据。"""
    return {
        "BeginDate": HASH_REPORT_BEGIN_DATE or "",
        "EndDate": HASH_REPORT_END_DATE or "",
        "Status": HASH_REPORT_STATUS,
        "NowPage": page,
        "PageRow": HASH_REPORT_PAGE_ROW,
    }


def _fetch_hash_report_page(page):
    """带 token POST 拉取一页；鉴权失效时重新登录并重试一次。"""
    body = _hash_report_body(page)
    for attempt in (1, 2):
        token = _hash_report_token(force_refresh=(attempt == 2))
        headers = {
            "Authorization": token,   # 后端校验 header 原值，不能加 "Bearer " 前缀
            "Content-Type": "application/json",
        }
        resp = None
        try:
            resp = requests.post(HASH_REPORT_API, json=body, headers=headers, timeout=30)
            if getattr(resp, "status_code", 200) == 401:
                payload = {"code": 401}
            else:
                resp.raise_for_status()
                payload = _parse_hash_api_payload(resp)
        except Exception as e:
            raise RuntimeError("HashReport 第 %d 页请求/解析失败: %s" % (page, repr(e)))
        finally:
            if resp is not None:
                try:
                    resp.close()
                except Exception:
                    pass
        if payload.get("code") == 401:
            if attempt == 1:
                print("[*] HashReport token 已失效，重新登录后重试第 %d 页" % page)
                continue
            raise RuntimeError("HashReport 第 %d 页鉴权失败（重新登录后仍被拒绝）" % page)
        return payload
    raise RuntimeError("HashReport 第 %d 页请求失败" % page)


def _iter_hash_report_pages():
    """① 逐页产出接口数据（生成器）：任意时刻只有「一页」在内存。

    - 先登录取 token，之后每页都带 Authorization 头 POST 请求；
    - 页数有硬上限，并检测「后端忽略 NowPage、一直返回第一页」的死循环；
    - 任一页失败直接抛异常，绝不静默返回残缺数据（残缺数据不能被全量上传）。
    """
    prev_first_id = None
    for page in range(1, HASH_REPORT_MAX_PAGES + 1):
        payload = _fetch_hash_report_page(page)

        if payload.get("code") != 200:
            raise RuntimeError("HashReport 第 %d 页返回异常: %s" % (page, str(payload)[:200]))
        data = payload.get("data")
        rows = (data.get("List") or []) if isinstance(data, dict) else []
        if not isinstance(rows, list):
            raise RuntimeError("HashReport 第 %d 页 List 不是列表: %s" % (page, type(rows)))
        if not rows:
            print("[*] HashReport 第 %d 页返回空，翻页结束。" % page)
            return

        first_id = str(rows[0].get("T_AutoID")) if isinstance(rows[0], dict) else ""
        if first_id and first_id == prev_first_id:
            raise RuntimeError("HashReport 第 %d 页与上一页首条数据相同（后端未翻页），已中止" % page)
        prev_first_id = first_id

        print("[*] HashReport 第 %d 页返回 %d 条" % (page, len(rows)))
        yield rows
        # 调用方消费完这一页、恢复执行时断开引用，让该页对象立刻可回收
        rows = None
        payload = None

    raise RuntimeError("HashReport 分页超过 %d 页上限，请缩小日期范围或调大 HASH_REPORT_MAX_PAGES"
                       % HASH_REPORT_MAX_PAGES)


def _rows_file_path():
    """JSONL 中转文件放在系统临时目录，避免打包后 exe 目录无写权限。"""
    return os.path.join(tempfile.gettempdir(), ROWS_FILE_NAME)


def _build_flat_row(item, prefix):
    """把一条接口数据的一类报告展开成「xlsx 7 列 + PDF 下载信息」的扁平行。

    字段映射与 _report_rows_for_item 完全一致，不改动业务规则。
    """
    if not _status_is_pending(item.get(prefix + "UploadStatus")):
        return None
    report_no = str(item.get(prefix + "ReportNo") or "").strip()
    if not report_no:
        return None
    date_str, year = _parse_report_date(item.get(prefix + "ReportDate"))
    client = item.get(prefix + "ReportWT")
    return {
        "no": report_no,
        "signer": item.get(prefix + "ReportQF") or "",
        "date": date_str,
        "year": year,
        "client": "/" if _is_empty_ltqj(client) else client,
        "pdf_name": _RE_WIN_INVALID.sub("_", report_no).strip(),
        "pdf_url": _report_pdf_url(item, prefix),
        # 供流程结束后回写 hashReport 上传状态使用（模板填写/PDF下载只读自己需要的键）
        "auto_id": item.get("T_AutoID"),
        "group": prefix,
    }


def _collect_rows_to_jsonl(rows_file, exclude_nos=None):
    """② 逐页拉取 → 展开 → 追加写 JSONL；只用磁盘做去重，内存不随总条数增长。

    exclude_nos: 已上传报告编号集合（规范化后），命中的行不落盘，
                 因而不会写进 xlsx，也不会被后面的 OrigDoc PDF 下载步骤消费。

    Returns: (row_count, bad_count, dup_count, filtered_count)
    """
    tmp = rows_file + ".part"
    db = rows_file + ".db"
    for p in (tmp, db):
        try:
            if os.path.isfile(p):
                os.remove(p)
        except OSError:
            pass

    # 两侧都用同一规则规范化，避免大小写/空格差异导致静默漏过滤
    exclude = frozenset(_norm_report_no(n) for n in (exclude_nos or ()))
    conn = sqlite3.connect(db)
    conn.execute("PRAGMA journal_mode=OFF")
    conn.execute("PRAGMA synchronous=OFF")
    conn.execute("CREATE TABLE seen (name TEXT PRIMARY KEY)")
    row_count = bad_count = dup_count = filtered_count = 0
    try:
        with open(tmp, "w", encoding="utf-8", newline="\n") as f:
            for page in _iter_hash_report_pages():
                for item in page:
                    if not isinstance(item, dict):
                        continue
                    for prefix in UPLOAD_FIELD_PREFIXES:
                        rec = _build_flat_row(item, prefix)
                        if rec is None:
                            if _status_is_pending(item.get(prefix + "UploadStatus")):
                                bad_count += 1
                            continue
                        # 「已上传报告管理 / 全部报告」列表里已有的编号直接跳过
                        if exclude and _norm_report_no(rec["no"]) in exclude:
                            filtered_count += 1
                            continue
                        if not rec["pdf_name"] or not rec["pdf_url"]:
                            # 仍写入该行：xlsx 必须包含这条数据；PDF 步骤会因缺 URL 判失败
                            bad_count += 1
                        elif conn.execute(
                                "INSERT OR IGNORE INTO seen(name) VALUES (?)",
                                (rec["pdf_name"],)).rowcount == 0:
                            dup_count += 1
                            continue
                        f.write(json.dumps(rec, ensure_ascii=False) + "\n")
                        row_count += 1
                del page                      # 一页处理完立刻释放
        os.replace(tmp, rows_file)
    finally:
        try:
            conn.close()
        except Exception:
            pass
        for p in (db, tmp):
            try:
                if os.path.isfile(p):
                    os.remove(p)
            except OSError:
                pass

    print("[+] 报告数据展开完成：%d 行，无效 %d 条，重复编号 %d 条，已上传过滤 %d 条 → %s"
          % (row_count, bad_count, dup_count, filtered_count, rows_file))
    return row_count, bad_count, dup_count, filtered_count


def _iter_jsonl(path):
    """顺序流式读 JSONL，任意时刻只有一行 dict 在内存。"""
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                yield json.loads(line)


def _count_jsonl(path):
    n = 0
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            if line.strip():
                n += 1
    return n


def _discard_rows_file():
    p = _rows_file_path()
    for f in (p, p + ".part", p + ".db"):
        try:
            if os.path.isfile(f):
                os.remove(f)
        except OSError:
            pass


def _reset_run_context():
    """重新执行时清空跨步骤上下文，避免上一轮的数据/句柄残留。"""
    global detected_target
    _discard_rows_file()
    DOWNLOAD_CONTEXT.clear()
    IMPORT_CONTEXT.clear()
    detected_target = None


def _parse_report_date(value):
    """从日期字段中解析出 (日期字符串 YYYY-MM-DD, 年份)。"""
    s = str(value or "").strip()
    if not s:
        return "", ""
    # 取 T 或空格之前的日期部分
    for sep in ("T", " "):
        if sep in s:
            s = s.split(sep, 1)[0]
            break
    year = ""
    if len(s) >= 4 and s[:4].isdigit():
        year = s[:4]
    return s, year


def _report_rows_for_item(item: dict) -> list:
    """把一条接口数据按 ft / lt / qj 类型展开为 0~3 行待填数据。

    某类型 {前缀}UploadStatus == 1（要上传）时生成一行，三个都是 1 则生成三行。
    字段映射（以 ft 为例，lt / qj 同理）：
        报告编号=ftReportNo，授权签字人=ftReportQF，
        签发日期/年份=ftReportDate，委托人=ftReportWT。
    """
    rows = []
    for prefix in UPLOAD_FIELD_PREFIXES:
        if not _status_is_pending(item.get(prefix + "UploadStatus")):
            continue
        report_no = str(item.get(prefix + "ReportNo") or "").strip()
        if not report_no:
            print("[!] 跳过一条 %sUploadStatus=1（要上传）但 %sReportNo 为空的数据: %s"
                  % (prefix, prefix, str(item)[:120]))
            continue
        date_str, year = _parse_report_date(item.get(prefix + "ReportDate"))
        client = item.get(prefix + "ReportWT")
        rows.append({
            "no": report_no,
            "signer": item.get(prefix + "ReportQF") or "",
            "date": date_str,
            "year": year,
            "client": "/" if _is_empty_ltqj(client) else client,
        })
    return rows


def _fill_template_rows(ws, report_list: list) -> int:
    """按接口返回的 report_list 把数据写入工作表 ws。

    每条数据按 ft / lt / qj 的 UploadStatus==1（要上传）展开为 0~3 行写入。
    返回填充行数。
    """
    clear_from = FILL_CLEAR_FROM_ROW
    start = FILL_START_ROW

    # 1) 清空原有数据行
    for row in ws.iter_rows(min_row=clear_from, max_row=ws.max_row):
        for c in row:
            c.value = None

    # 2) 检验检测机构名称写入 C2
    ws["C2"] = FILL_ORG

    # 3) 逐条展开并写入（一条数据最多展开 3 行：ft / lt / qj）
    count = 0
    for item in report_list:
        if not isinstance(item, dict):
            continue
        for row_data in _report_rows_for_item(item):
            r = start + count
            ws.cell(row=r, column=1, value=row_data["no"])
            ws.cell(row=r, column=2, value=row_data["signer"])
            ws.cell(row=r, column=3, value=row_data["date"])
            ws.cell(row=r, column=4, value=row_data["year"])
            # 第 5 列「哈希值」留空，系统自动生成
            ws.cell(row=r, column=6, value="是")
            ws.cell(row=r, column=7, value=row_data["client"])
            count += 1

    return count


# ===================== 流式写表（大批量内存恒定） =====================
# 说明：openpyxl 常规模式会把每个单元格物化成 Cell 对象，60 万行 × 7 列会
#       占 1~2GB 内存。这里改为「流式改写原 xlsx 的 sheet1.xml」：
#         - zip 里除 sheet1.xml 外的部件（styles / sharedStrings / 合并单元格 /
#           数据验证 / 列宽 / 条件格式）原样复制，表头与格式 100% 保留；
#         - sheetData 里保留第 1~3 行表头，跳过 r>=4 的旧数据行；
#         - C2 原地写入 FILL_ORG（等价于原来的 ws["C2"] = FILL_ORG）；
#         - 新数据行逐行写进 zip 流，内存里始终只有一行。
#       原 _fill_template_rows 保留作为小数据量回退/参考，不再被流程调用。
_SHEET_PART = "xl/worksheets/sheet1.xml"
_RE_DIMENSION = re.compile(r'(<dimension\b[^>]*\bref=")[^"]*(")')
_RE_ROW = re.compile(r'<row\b[^>]*>.*?</row>|<row\b[^>]*/\s*>', re.S)
_RE_ROW_NUM = re.compile(r'\br="(\d+)"')
_RE_CELL_C2 = re.compile(
    r'<c\b[^>]*\br="C2"[^>]*>.*?</c>|<c\b[^>]*\br="C2"[^>]*/\s*>', re.S)


def _xml_text(value):
    """XML 文本转义，并剔除会让 xlsx 失效的控制字符。"""
    s = str(value)
    s = "".join(ch for ch in s if ch >= " " or ch in "\t\n\r")
    return _xml_escape(s)


def _cell_inline_xml(ref, value, style=""):
    s = ' s="%s"' % style if style else ""
    return '<c r="%s"%s t="inlineStr"><is><t>%s</t></is></c>' % (
        ref, s, _xml_text(value))


def _sheet_row_xml(row_num, rec, style_map):
    vals = (rec["no"], rec["signer"], rec["date"], rec["year"], None,
            "是", rec["client"])
    cols = "ABCDEFG"
    parts = ['<row r="%d">' % row_num]
    for i, v in enumerate(vals):
        if v is None:
            continue                      # 第 5 列「哈希值」留空，系统自动生成
        parts.append(_cell_inline_xml("%s%d" % (cols[i], row_num), v,
                                     style_map.get(i, "")))
    parts.append("</row>")
    return "".join(parts)


def _ensure_c2_in_row(row_xml):
    """在第 2 行里原地写入机构名；模板原本没有 C2 时补一个 cell。"""
    c2 = _RE_CELL_C2.search(row_xml)
    if c2:
        old = c2.group(0)
        head_end = old.find(">")
        sm = re.search(r'\bs="(\d+)"', old[:head_end + 1])
        style = sm.group(1) if sm else ""
        return row_xml[:c2.start()] + _cell_inline_xml("C2", FILL_ORG, style) + row_xml[c2.end():]
    cell = _cell_inline_xml("C2", FILL_ORG)
    stripped = row_xml.rstrip()
    if stripped.endswith("/>"):
        return stripped[:-2] + ">" + cell + "</row>"
    return row_xml.replace("</row>", cell + "</row>")


def _fill_template_streaming(path, rows_file, row_count):
    """③ 流式改写 sheet1.xml，产出「一个」包含全量数据的 xlsx。"""
    tmp = path + ".stream.xlsx"
    try:
        with zipfile.ZipFile(path, "r") as zin:
            if _SHEET_PART not in zin.namelist():
                raise RuntimeError("模板里找不到 %s，无法流式写表" % _SHEET_PART)
            xml = zin.read(_SHEET_PART).decode("utf-8")

            self_close = re.search(r'<sheetData\b[^>]*/\s*>', xml)
            if self_close:
                head = xml[:self_close.start()] + "<sheetData>"
                body = ""
                tail = "</sheetData>" + xml[self_close.end():]
            else:
                m = re.search(r'<sheetData\b[^>]*>', xml)
                if not m:
                    raise RuntimeError("模板 sheet1.xml 里没有 <sheetData>")
                end = xml.find("</sheetData>", m.end())
                if end < 0:
                    raise RuntimeError("模板 sheet1.xml 的 <sheetData> 未闭合")
                head, body, tail = xml[:m.end()], xml[m.end():end], xml[end:]

            header_rows = []
            for rm in _RE_ROW.finditer(body):
                tag = rm.group(0)
                num = _RE_ROW_NUM.search(tag)
                if not num or int(num.group(1)) > 3:
                    continue
                if int(num.group(1)) == 2:
                    tag = _ensure_c2_in_row(tag)
                header_rows.append(tag)

            # 继承原模板第 4 行各列的样式索引，保证新数据行外观不变
            style_map = {}
            r4 = re.search(r'<row\b[^>]*\br="4"[^>]*>.*?</row>', body, re.S)
            if r4:
                for cm in re.finditer(r'<c\b([^>]*)\br="([A-G])4"', r4.group(0)):
                    sm = re.search(r'\bs="(\d+)"', cm.group(1))
                    if sm:
                        style_map["ABCDEFG".index(cm.group(2))] = sm.group(1)

            head = _RE_DIMENSION.sub(
                lambda mm: mm.group(1) + "A1:G%d" % (3 + row_count) + mm.group(2),
                head, count=1)

            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for info in zin.infolist():
                    if info.filename == _SHEET_PART:
                        zi = zipfile.ZipInfo(_SHEET_PART, date_time=info.date_time)
                        zi.compress_type = zipfile.ZIP_DEFLATED
                        zi.external_attr = info.external_attr
                        with zout.open(zi, "w") as w:
                            w.write(head.encode("utf-8"))
                            for tag in header_rows:
                                w.write(tag.encode("utf-8"))
                            r = FILL_START_ROW
                            for rec in _iter_jsonl(rows_file):
                                w.write(_sheet_row_xml(r, rec, style_map).encode("utf-8"))
                                r += 1
                            w.write(tail.encode("utf-8"))
                    else:
                        zout.writestr(info, zin.read(info.filename))
        os.replace(tmp, path)
    finally:
        try:
            if os.path.isfile(tmp):
                os.remove(tmp)
        except OSError:
            pass
    return row_count


def _verify_filled_template(path, expected_rows):
    """生成后自检：表头/C2/行数必须与改造前逻辑等价，否则终止流程。"""
    try:
        from openpyxl import load_workbook
    except ImportError:
        # 写表本身不依赖 openpyxl（直接流式改写 XML），只是自检需要它
        print("[*] 未安装 openpyxl，跳过模板保真自检（不影响已写好的文件）。")
        return
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        if ws is None:
            raise RuntimeError("模板里没有活动工作表，无法自检")
        head = list(ws.iter_rows(min_row=1, max_row=3, values_only=True))
        if len(head) < 2 or len(head[1]) < 3 or str(head[1][2] or "").strip() != FILL_ORG:
            raise RuntimeError("C2 机构名校验失败，模板表头可能被破坏")
        want = 3 + expected_rows
        if ws.max_row != want:
            # 模板可能没有 <dimension>，再用流式遍历确认一次
            actual = 3 + sum(1 for _ in ws.iter_rows(
                min_row=4, max_row=want, values_only=True))
            if actual != want:
                raise RuntimeError("行数校验失败：期望 %d，实际 %d" % (want, actual))
    finally:
        wb.close()
    print("[+] 模板保真自检通过：表头 3 行保留、C2=%s、共 %d 行数据"
          % (FILL_ORG, expected_rows))


def _step_fill_template_data():
    """在刚下载好的 .xlsx 模板中写入「过滤后」的数据（流式，内存恒定）。

    数据来源：HashReport 接口逐页拉取 → JSONL 磁盘中转。
    写表前 _step_collect_uploaded_report_nos 已从「已上传报告管理 / 全部报告」
    列表逐页采集报告编号，这里按编号排除已上传数据，只把剩余数据写入 xlsx。
    无论是否启用 AUTO_FILL_TEMPLATE 都先准备中转文件，供后面的 PDF 下载步骤
    复用，保证「只请求一次接口」且两次消费的数据完全一致。
    """
    path = DOWNLOAD_CONTEXT.get("path") or _locate_template_file()
    if not path:
        print("[!] 未定位到模板文件，终止流程。")
        return False
    DOWNLOAD_CONTEXT["path"] = path

    exclude_nos = UPLOADED_REPORT_NOS if FILTER_UPLOADED_REPORTS else None
    rows_file = _rows_file_path()
    if DOWNLOAD_CONTEXT.get("rows_file") == rows_file and os.path.isfile(rows_file):
        row_count = DOWNLOAD_CONTEXT.get("row_count") or _count_jsonl(rows_file)
        print("[*] 复用本轮已展开的报告数据中转文件：%d 行（已上传过滤 %d 条）"
              % (row_count, DOWNLOAD_CONTEXT.get("filtered_count") or 0))
    else:
        try:
            row_count, bad_count, dup_count, filtered_count = _collect_rows_to_jsonl(
                rows_file, exclude_nos)
        except Exception as e:
            print("[!] 拉取/展开报告数据失败，终止流程: " + repr(e))
            return False
        DOWNLOAD_CONTEXT["rows_file"] = rows_file
        DOWNLOAD_CONTEXT["row_count"] = row_count
        DOWNLOAD_CONTEXT["bad_count"] = bad_count
        DOWNLOAD_CONTEXT["dup_count"] = dup_count
        DOWNLOAD_CONTEXT["filtered_count"] = filtered_count
        print("[*] 数据规模：%d 行（无效 %d 条，重复编号 %d 条，已上传过滤 %d 条）"
              % (row_count, bad_count, dup_count, filtered_count))

    if row_count <= 0:
        # 没有「要上传」的数据属于正常情况（例如全部报告都已上传），
        # 不记为失败：跳过写表/打开/导入/上传，本轮流程正常结束。
        if exclude_nos:
            print("[*] 过滤已上传报告后没有需要填写的新数据（报告可能已全部上传）。")
        else:
            print("[*] 接口未返回任何「要上传」的报告数据。")
        print("[*] 本轮无需上传，跳过后续的写表、打开表格、导入与上传步骤。")
        return FLOW_STOP

    if not CONFIG.get("AUTO_FILL_TEMPLATE", True):
        print("[*] 已关闭「自动填写模板数据」，跳过写表；中转数据保留供 PDF 下载。")
        return True

    try:
        count = _fill_template_streaming(path, rows_file, row_count)
        _verify_filled_template(path, row_count)
    except Exception as e:
        print("[!] 流式写表或模板保真自检失败，终止流程: " + repr(e))
        return False

    print("[+] 已写入 %d 行数据到: %s（机构=%s，已排除 %d 条已上传报告）"
          % (count, path, FILL_ORG, DOWNLOAD_CONTEXT.get("filtered_count") or 0))
    return True


def _step_open_template_file():
    """第五步：用系统默认程序打开刚刚填写好的 .xlsx 文件。"""
    path = DOWNLOAD_CONTEXT.get("path") or _locate_template_file()
    if not path:
        print("[!] 未定位到模板文件，无法打开。")
        return False
    DOWNLOAD_CONTEXT["path"] = path

    row_count = int(DOWNLOAD_CONTEXT.get("row_count") or 0)
    if row_count >= LARGE_DATA_ROW_THRESHOLD:
        print("[*] 全量数据 %d 行：跳过「打开 xlsx」预览，"
              "避免 Excel 加载超时与文件锁导致后续导入失败。" % row_count)
        return True

    if not CONFIG.get("AUTO_OPEN_TEMPLATE", True):
        print("[*] 已关闭「自动打开 xlsx」，仅完成填写。文件位置: " + path)
        return True

    try:
        os.startfile(path)
        print("[+] 已用系统默认程序打开该 .xlsx 文件: " + path)
    except Exception as e:
        print("[!] 打开文件失败: " + repr(e))
    return True


# ===================== 按报告编号下载 OrigDoc PDF =====================
# 报告原件放在上传系统服务器上：
#   {域名}/Server/Upload/Formal/TestCheckOrder/{T_AutoID}/UploadPic/{ltReport}
# 这里按 HashReport 接口返回的每一条数据，用 T_AutoID + ltReport 拼地址，
# 把 PDF 下载到 %ProgramData%\AreportFile\OrigDoc，文件名用「报告编号」。
# 下载前会先清空 OrigDoc 里的旧文件，避免上一轮遗留的旧原件被一起导入上传。
_RE_WIN_INVALID = re.compile(r'[\\/:*?"<>|]')

# [保留备用] 旧版是本地拼装的占位 PDF；现在改为从服务器下载真实报告，
# 这两个函数流程里不再调用，保留以便离线演示 / 排障时使用。
def _pdf_escape_text(text: str) -> str:
    """把文本转成可写入 PDF 字符串的内容：转义 ( ) \\，并剔除非 latin-1 字符。"""
    out = []
    for ch in str(text or ""):
        if ch in "()\\":
            out.append("\\" + ch)
            continue
        try:
            ch.encode("latin-1")
        except UnicodeEncodeError:
            continue
        out.append(ch)
    return "".join(out)


def _make_simple_pdf(path: str, text: str) -> bool:
    """手工构造一个最简单的合法单页 PDF（PDF 1.4），页面上只画一行文本。

    不依赖任何第三方库：自行拼装 Catalog / Pages / Page / Contents / Font
    五个对象，并按真实字节偏移生成 xref 表、trailer、startxref 与 %%EOF。

    Args:
        path: 输出 PDF 的完整路径（已存在则覆盖）。
        text: 要绘制的文本（Helvetica 24pt，位于页面左上角区域）。

    Returns:
        bool: 写入成功返回 True，失败返回 False。
    """
    content = ("BT /F1 24 Tf 72 720 Td (%s) Tj ET\n" % _pdf_escape_text(text)).encode("latin-1")
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
        b"/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
        b"<< /Length " + str(len(content)).encode("ascii") + b" >>\nstream\n" + content + b"endstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
    ]

    buf = bytearray()
    buf += b"%PDF-1.4\n"
    offsets = []
    for idx, body in enumerate(objs, start=1):
        offsets.append(len(buf))
        buf += str(idx).encode("ascii") + b" 0 obj\n" + body + b"\nendobj\n"

    xref_pos = len(buf)
    size = len(objs) + 1
    buf += b"xref\n0 " + str(size).encode("ascii") + b"\n"
    buf += b"0000000000 65535 f \n"
    for off in offsets:
        buf += ("%010d 00000 n \n" % off).encode("ascii")
    buf += b"trailer\n<< /Size " + str(size).encode("ascii") + b" /Root 1 0 R >>\n"
    buf += b"startxref\n" + str(xref_pos).encode("ascii") + b"\n%%EOF\n"

    try:
        with open(path, "wb") as f:
            f.write(bytes(buf))
    except OSError as e:
        print("[!] 写入 PDF 失败: %s -> %s" % (path, repr(e)))
        return False
    return True


def _remove_file_force(path):
    """删除单个文件（只读属性导致的失败先去属性再删一次），失败返回 False。"""
    try:
        os.remove(path)
        return True
    except PermissionError:
        pass
    except OSError as e:
        print("[!] 删除失败（可能被占用）: %s -> %s" % (path, repr(e)))
        return False
    # 只读属性也会让 os.remove 抛 PermissionError，去掉属性后再试一次
    try:
        os.chmod(path, 0o666)
        os.remove(path)
        return True
    except OSError as e:
        print("[!] 删除失败（可能被占用）: %s -> %s" % (path, repr(e)))
        return False


def _clear_origdoc_dir(out_dir):
    """清空 OrigDoc 里的旧文件（只删内容，保留目录本身）。

    最后一步会把 OrigDoc 整个文件夹当作「报告电子原件」导入：上一轮遗留的
    旧 PDF / OFD 会跟着一起上传，页面上的「N 个文件」也会对不上，
    所以生成新 PDF 之前先把旧文件清掉。

    Returns:
        (deleted, failed): 删除成功个数 / 删不掉的名字列表（多数是被程序占用）。
    """
    # 流式删除：直接迭代 scandir 并立即删，不 sorted、不物化 DirEntry 列表，
    # 否则几十万个文件时会一次性占满内存。NTFS 边遍历边删可能跳项，
    # 所以最多扫 3 轮，直到某一轮没有删掉任何条目为止。
    deleted, failed = 0, []
    for _round in range(1, 4):
        round_deleted = 0
        try:
            with os.scandir(out_dir) as it:
                for ent in it:
                    try:
                        is_dir = ent.is_dir(follow_symlinks=False)
                    except OSError:
                        is_dir = False
                    if is_dir:
                        try:
                            shutil.rmtree(ent.path)
                            deleted += 1
                            round_deleted += 1
                        except OSError as e:
                            failed.append(ent.name)
                            print("[!] 删除子目录失败（可能被占用）: %s -> %s"
                                  % (ent.path, repr(e)))
                    elif _remove_file_force(ent.path):
                        deleted += 1
                        round_deleted += 1
                    else:
                        failed.append(ent.name)
        except OSError as e:
            print("[!] 读取 OrigDoc 目录失败: " + repr(e))
            return deleted, list(dict.fromkeys(failed))
        if round_deleted == 0:
            break
    return deleted, list(dict.fromkeys(failed))


def _report_pdf_url(item, prefix="lt"):
    """按 T_AutoID + {前缀}Report 拼出该类报告 PDF 的下载地址；字段缺失返回空串。

    Args:
        item: HashReport 接口返回的单条数据。
        prefix: 报告类型前缀（"ft" / "lt" / "qj"），对应取 ftReport / ltReport / qjReport 字段。

    地址规则：
        {域名}/Server/Upload/Formal/TestCheckOrder/{T_AutoID}/UploadPic/{ftReport|ltReport|qjReport}
    例http://ba.gdyjjc.com:2037/Server/Upload/Formal/TestCheckOrder/351635/UploadPic/mc20261104143458aft.pdf
    """
    auto_id = item.get("T_AutoID")
    lt_report = str(item.get(prefix + "Report") or "").strip()
    if auto_id is None or lt_report == "":
        return ""
    sid = str(auto_id).strip()
    # 接口里的 T_AutoID 是浮点数（如 351635.0），URL 里必须写成 351635
    try:
        num = float(sid)
        if num.is_integer():
            sid = str(int(num))
    except (TypeError, ValueError):
        pass
    if not sid:
        return ""
    return "%s/Server/Upload/Formal/TestCheckOrder/%s/UploadPic/%s" % (
        REPORT_PDF_BASE_URL.rstrip("/"), quote(sid, safe=""), quote(lt_report))


def _download_report_pdf(session, url, target):
    """把 url 指向的报告 PDF 下载到 target。

    Returns:
        (ok, reason): 成功 (True, "")；失败 (False, 失败原因)。
    """
    reason = ""
    for _attempt in (1, 2):  # 网络抖动重试一次；HTTP 404 这种不重试
        try:
            resp = session.get(url, timeout=(10, 120))
        except requests.RequestException as e:
            reason = "请求失败: " + repr(e)
            time.sleep(0.5)
            continue
        if resp.status_code != 200:
            return False, "HTTP %d" % resp.status_code
        content = resp.content or b""
        if not content.startswith(b"%PDF"):
            ctype = (resp.headers.get("Content-Type") or "未知").split(";")[0]
            return False, "返回的不是 PDF（Content-Type=%s，%d 字节）" % (ctype, len(content))
        tmp = target + ".part"
        try:
            with open(tmp, "wb") as f:
                f.write(content)
            os.replace(tmp, target)  # 先写临时文件再替换，避免中断留下半截 PDF
        except OSError as e:
            try:
                os.remove(tmp)
            except OSError:
                pass
            return False, "写入失败: " + repr(e)
        return True, ""
    return False, reason or "请求失败"


def _fmt_duration(seconds) -> str:
    """把秒数格式化成紧凑中文时长：X小时XX分XX秒 / X分XX秒 / X秒 / 0.35秒 / 19毫秒。

    不足 1 秒时不能再用 %d秒 取整，否则本机回环并发下载这类几十毫秒的耗时
    会一律显示成「总耗时0秒」。
    """
    try:
        value = max(0.0, float(seconds))
    except (TypeError, ValueError):
        return "未知"
    if value <= 0:
        return "0秒"
    if value < 1:
        if value >= 0.01:
            return "%.2f秒" % value
        return "%d毫秒" % max(1, int(round(value * 1000)))
    h, rem = divmod(int(round(value)), 3600)
    m, s = divmod(rem, 60)
    if h:
        return "%d小时%02d分%02d秒" % (h, m, s)
    if m:
        return "%d分%02d秒" % (m, s)
    return "%d秒" % s


def _render_progress_bar(done: int, total: int, width: int = 24) -> str:
    """生成纯文本进度条，例如 [██████████████░░░░░░░░░░░░] 50.0%。

    用 Unicode 方块字符，在 GUI 日志框（Tk Text）里显示正常；
    不依赖 tqdm 等第三方库，也不用 \\r 原地刷新（Tk 日志框不适合）。
    """
    total = max(1, int(total))
    ratio = max(0.0, min(1.0, done / total))
    filled = int(round(ratio * width))
    return "[" + "█" * filled + "░" * (width - filled) + "] %5.1f%%" % (ratio * 100)


def _step_generate_orig_pdfs():
    """按报告编号把服务器上的报告 PDF 下载到 OrigDoc 目录。

    流程：
      1) 创建 %ProgramData%\\AreportFile\\OrigDoc 目录；
      2) 直接复用填表步骤写好的 JSONL 中转文件，流式逐行读，内存不驻留全量数据；
      3) 下载前清空 OrigDoc 里的旧文件（保持平铺，不改应用选文件夹逻辑）；
      4) 有界并发下载（每线程独占 Session，在途任务数有上限），文件名 = 报告编号.pdf；
      5) 零容忍：任一份下载失败即停止提交新任务，并终止后续导入 / 上传。

    Returns:
        bool: 无数据 / 旧文件清不掉 / 任一份 PDF 下载失败时返回 False（阻断流程）。
    """
    program_data = os.environ.get("ProgramData") or r"C:\ProgramData"
    out_dir = os.path.join(program_data, "AreportFile", "OrigDoc")
    try:
        os.makedirs(out_dir, exist_ok=True)
    except OSError as e:
        print("[!] 创建 OrigDoc 目录失败，终止流程: %s -> %s" % (out_dir, repr(e)))
        return False

    # 1) 直接复用填表步骤写好的 JSONL 中转文件：只请求一次接口，
    #    且下载阶段内存里不再出现「全量 items 列表 / 全量去重 set」。
    rows_file = DOWNLOAD_CONTEXT.get("rows_file") or _rows_file_path()
    if not os.path.isfile(rows_file):
        print("[!] 未找到报告数据中转文件（填表步骤未成功完成），终止流程。")
        return False
    total = DOWNLOAD_CONTEXT.get("row_count") or _count_jsonl(rows_file)
    if total <= 0:
        print("[!] 中转文件里没有可下载的数据，终止流程。")
        return False
    bad_count = int(DOWNLOAD_CONTEXT.get("bad_count") or 0)
    dup_count = int(DOWNLOAD_CONTEXT.get("dup_count") or 0)
    data_source = "本轮 JSONL 中转文件"

    # 3) 下载前先清空 OrigDoc 里的旧文件（保持平铺，不改应用选文件夹的逻辑）
    deleted, failed = _clear_origdoc_dir(out_dir)
    if failed:
        print("[!] 有 %d 个旧文件删不掉（可能被占用），零容忍下终止流程: %s"
              % (len(failed), "，".join(failed[:10])))
        return False

    # 4) 有界并发下载：内存里最多 PDF_MAX_INFLIGHT 个在途任务，
    #    不构造全量 items 列表；失败按 PDF_FAIL_TOLERANCE 判定（0 = 零容忍）。
    state: dict[str, Any] = {"ok": 0, "fail": 0, "first": []}
    lock = threading.Lock()
    abort = threading.Event()
    inflight = threading.BoundedSemaphore(PDF_MAX_INFLIGHT)
    session_pool = queue.Queue()      # 每线程独占一个 Session，避免并发读写同一连接池
    for _ in range(PDF_DOWNLOAD_WORKERS):
        s = requests.Session()
        s.headers.update({"User-Agent": "auto-upload-demo/1.0"})
        session_pool.put(s)
    # perf_counter 精度高且单调：本机回环下载常常只有几十毫秒，
    # 用 time.time() 容易被 15.6ms 的时钟分辨率抹平成 0 秒。
    t0 = time.perf_counter()
    last_progress_ts = [0.0]

    def worker(rec):
        s = session_pool.get()
        try:
            if not rec.get("pdf_url") or not rec.get("pdf_name"):
                return False, "缺少报告编号或下载地址", rec.get("no", "")
            target = os.path.join(out_dir, rec["pdf_name"] + ".pdf")
            ok, reason = _download_report_pdf(s, rec["pdf_url"], target)
            return ok, reason, rec.get("no", "")
        except Exception as e:
            return False, repr(e), rec.get("no", "")
        finally:
            session_pool.put(s)

    def on_done(fut):
        try:
            ok, reason, code = fut.result()
        except Exception as e:
            ok, reason, code = False, repr(e), "?"
        with lock:
            if ok:
                state["ok"] += 1
            else:
                state["fail"] += 1
                if len(state["first"]) < 5:
                    state["first"].append((code, reason))
                abort.set()                # 零容忍：立即停止提交新任务
            done = state["ok"] + state["fail"]
        inflight.release()
        now = time.perf_counter()
        if done >= total or now - last_progress_ts[0] >= 0.2:
            last_progress_ts[0] = now
            elapsed = now - t0
            speed = done / elapsed if elapsed > 0 else 0.0  # 只用于估算剩余时间，不再显示
            eta_sec = (total - done) / speed if speed > 0 else 0.0
            finish_at = (time.strftime("%H:%M:%S", time.localtime(time.time() + eta_sec))
                         if eta_sec > 0 else "--:--:--")
            set_progress("[*] PDF下载 %s %d/%d | 成功%d 失败%d | 已用%s 预计剩余%s | 预计%s完成"
                         % (_render_progress_bar(done, total), done, total,
                            state["ok"], state["fail"], _fmt_duration(elapsed),
                            _fmt_duration(eta_sec), finish_at))

    bad_note = ("，无效数据%d条" % bad_count) if bad_count else ""
    print("[*] 开始下载报告 PDF：共 %d 份 → %s（数据%s，并发%d，清理旧文件%d项，跳过重复编号%d条%s）"
          % (total, out_dir, data_source, PDF_DOWNLOAD_WORKERS, deleted, dup_count, bad_note))
    set_progress("[*] PDF下载 准备中 ... 0/%d" % total)

    try:
        with ThreadPoolExecutor(max_workers=PDF_DOWNLOAD_WORKERS) as ex:
            for rec in _iter_jsonl(rows_file):
                if abort.is_set():
                    break
                inflight.acquire()
                try:
                    fut = ex.submit(worker, rec)
                    fut.add_done_callback(on_done)
                except Exception as e:
                    inflight.release()
                    with lock:
                        state["fail"] += 1
                        state["first"].append(("提交任务", repr(e)))
                    abort.set()
                    break
            # with 退出时等待已提交任务收尾（在途最多 PDF_MAX_INFLIGHT 个）
    finally:
        while True:
            try:
                session_pool.get_nowait().close()
            except queue.Empty:
                break
            except Exception:
                pass

    total_cost = time.perf_counter() - t0
    done = state["ok"] + state["fail"]
    set_progress("")

    summary = ("[+] PDF下载 %s %d/%d | 成功%d 失败%d | 总耗时%s | 目录: %s"
               % (_render_progress_bar(done, total), done, total,
                  state["ok"], state["fail"], _fmt_duration(total_cost), out_dir))
    if state["first"]:
        head = "；".join("%s(%s)" % (c, r) for c, r in state["first"])
        summary += " | 失败明细(前%d条): %s" % (len(state["first"]), head)
    print(summary)

    if state["fail"] > PDF_FAIL_TOLERANCE:
        print("[!] 零容忍：本轮有 %d 份 PDF 下载失败，已终止后续「选文件夹 / 数据上传」。" % state["fail"])
        print("[!] OrigDoc 中已下载的文件不完整（仅供排查），下一轮启动会自动清空，请勿手动上传。")
        return False
    return True


# ===================== 选择「电子原件文件夹」并导入 OrigDoc =====================
# 业务背景：上传系统页面是 3 步向导
#           （① 选择数据表格 → ② 选择电子原件 → ③ 数据上传）。
# 第 ② 步的 UI 有两种形态：
#   - 未选过：一个蓝色按钮「选择报告电子原件文件夹导入」；
#   - 已选过：变成一张卡片「已选择文件夹: OrigDoc (15 个文件)」+ 右侧「重新选择」按钮。
# 点这个按钮会弹出 Windows 原生的「选择报告文件夹」对话框
# （Vista/IFileDialog 风格，class "#32770"，底部有「文件夹:」输入框 +
#  「选择文件夹」/「取消」按钮）。
# 本节负责三件事：
#   1) 在页面里找到并点开这个按钮（拿到原生对话框）；
#   2) 在原生对话框里选中 %ProgramData%\AreportFile\OrigDoc 并点「选择文件夹」；
#   3) 回到页面校验是否显示「已选择文件夹: OrigDoc」。
ELEC_FOLDER_BTN_TIMEOUT = 15      # 等待页面上目标按钮出现的最长秒数
ELEC_FOLDER_DIALOG_TIMEOUT = 20   # 等待原生「选择文件夹」对话框弹出的最长秒数
ELEC_FOLDER_PICK_TIMEOUT = 8      # 点击「选择文件夹」后等待对话框关闭的最长秒数
ELEC_FOLDER_VERIFY_TIMEOUT = 15   # 等待页面刷新出「已选择文件夹」的最长秒数

# 在页面里定位第 ② 步的按钮，返回它的坐标（不直接点击，交给 CDP 做可信点击）。
# 说明：JS 合成的 .click() 在 Electron 里不一定能拉起原生对话框，
#       所以这里只负责「找」和「算坐标」。
JS_CLICK_SELECT_ELECTRONIC_FOLDER = r"""
(function () {
    const PICK_KEY = '选择报告电子原件文件夹导入';
    const RESELECT_KEY = '重新选择';
    const CARD_KEY = '已选择文件夹';
    function norm(s) { return ((s || '') + '').replace(/\s+/g, ''); }
    function textOf(el) { return norm(el.innerText || el.textContent || ''); }
    function isClickable(el) {
        return /^(BUTTON|A)$/i.test(el.tagName) ||
               el.getAttribute('role') === 'button' ||
               /arco-btn|btn/i.test((el.className || '') + '');
    }
    // 元素在 DOM 里的深度：同样长度的文本里，更深的那个才是真正的按钮
    // （外层容器 div / span 往往与按钮文本完全一样长）
    function depthOf(el) {
        let d = 0, p = el.parentElement;
        while (p && d < 50) { d++; p = p.parentElement; }
        return d;
    }

    const all = Array.from(document.querySelectorAll('button, a, span, div, i'));
    let el = null;
    let kind = 'pick';

    // 1) 未选过文件夹：页面上有「选择报告电子原件文件夹导入」按钮
    const hits = all.filter(e => textOf(e).indexOf(norm(PICK_KEY)) >= 0);
    if (hits.length) {
        // 优先选真正可点击的元素（button / a / .arco-btn），
        // 否则会点到外层容器 div，点击不会有任何反应
        const clickables = hits.filter(isClickable);
        const pool = clickables.length ? clickables : hits;
        pool.sort((a, b) => (textOf(a).length - textOf(b).length) ||
                            (depthOf(b) - depthOf(a)));
        el = pool[0];
    } else {
        // 2) 已选过文件夹：按钮变成「重新选择」。
        //    页面上第 ① 步区域可能也有一个「重新选择」，必须区分开：
        //    只有位于「已选择文件夹」卡片内的那个才算命中。
        function collectReselect(exactOnly) {
            const out = [];
            all.forEach(e => {
                const t = textOf(e);
                const matched = exactOnly ? (t === norm(RESELECT_KEY))
                                         : (t.indexOf(norm(RESELECT_KEY)) >= 0);
                if (!matched) return;
                let p = e.parentElement;
                let up = 0;
                while (p && up < 8) {
                    if (textOf(p).indexOf(norm(CARD_KEY)) >= 0) {
                        out.push({ el: e, depth: up });
                        break;
                    }
                    p = p.parentElement;
                    up++;
                }
            });
            return out;
        }
        let cands = collectReselect(true);
        if (!cands.length) cands = collectReselect(false);
        if (cands.length) {
            const clickables = cands.filter(c => isClickable(c.el));
            const pool = clickables.length ? clickables : cands;
            // depth 越小 = 离「已选择文件夹」卡片越近，越可能是真正的那个按钮
            pool.sort((a, b) => (a.depth - b.depth) ||
                                (textOf(a.el).length - textOf(b.el).length) ||
                                (depthOf(b.el) - depthOf(a.el)));
            el = pool[0].el;
            kind = 'reselect';
        }
    }

    if (!el) {
        return { ok: false, msg: '未找到「' + PICK_KEY + '」/「' + RESELECT_KEY + '」按钮' };
    }
    if (el.disabled || el.getAttribute('aria-disabled') === 'true') {
        return { ok: false, msg: '按钮当前不可用' };
    }
    try { el.scrollIntoView({ block: 'center' }); } catch (e) {}
    const r = el.getBoundingClientRect();
    return {
        ok: true,
        kind: kind,
        x: Math.round(r.x + r.width / 2),
        y: Math.round(r.y + r.height / 2),
        w: Math.round(r.width),
        h: Math.round(r.height),
        tag: el.tagName,
        cls: (el.className || '') + '',
        text: textOf(el)
    };
})()
"""

# 读页面上是否已显示「已选择文件夹: <名字>（<n> 个文件）」。
JS_CHECK_SELECTED_FOLDER = r"""
(function () {
    const txt = (document.body && (document.body.innerText || document.body.textContent)) || '';
    const nameM = txt.match(/已选择文件夹\s*[：:]?\s*([^\s（(，,]*)/);
    const cntM = txt.match(/(\d+)\s*个文件/);
    return {
        found: !!nameM,
        name: nameM ? (nameM[1] || '') : '',
        fileCount: cntM ? parseInt(cntM[1], 10) : 0,
        hasOrigDoc: txt.indexOf('OrigDoc') >= 0
    };
})()
"""


def _win_find_folder_dialog():
    """查找原生「选择文件夹」对话框，返回 (hwnd, title, button_text)。

    目标窗口是 Windows 原生的「选择报告文件夹」对话框（class "#32770"），
    底部有「文件夹:」输入框和「选择文件夹」按钮。

    优先级（从高到低）：
      1) 标题命中且属于「上传系统」进程（CONFIG["EXE_PATH"]）；
      2) 标题命中（「选择报告文件夹」/「选择文件夹」/「Select Folder」/「Folder」）；
      3) 属于「上传系统」进程（标题文案不匹配时兜底）；
      4) 任意带「选择文件夹」按钮的 #32770 对话框。

    找不到时返回 (0, "", "")——句柄为 0 即表示「当前没有这个对话框」。
    """
    want_exe = os.path.normcase(CONFIG.get("EXE_PATH") or "")
    best = None          # (rank, hwnd, title, btn_text)
    for hwnd, cls, title in _win_top_windows():
        if cls != "#32770":
            continue
        btn, txt = _win_find_button(hwnd, ("选择文件夹", "Select Folder"))
        if not btn:
            continue
        title_ok = any(k in title for k in
                       ("选择报告文件夹", "选择文件夹", "Select Folder", "Folder"))
        proc_ok = False
        if want_exe:
            proc_exe = _win_process_exe(_win_window_pid(hwnd))
            proc_ok = bool(proc_exe) and os.path.normcase(proc_exe) == want_exe
        if title_ok and proc_ok:
            rank = 0
        elif title_ok:
            rank = 1
        elif proc_ok:
            rank = 2
        else:
            rank = 3
        if best is None or rank < best[0]:
            best = (rank, hwnd, title, txt)
            if rank == 0:
                break
    if best is None:
        return 0, "", ""
    return best[1], best[2], best[3]


def _win_find_folder_name_edit(dlg_hwnd):
    """定位「文件夹:」输入框，返回 (hwnd, 当前文本)。

    Vista 风格「选择文件夹」对话框底部有一个「文件夹:」Static 标签 + 一个 Edit。
    优先按该标签定位（取其右侧、垂直位置接近的 Edit 里水平距离最近的那个），
    这样即使对话框里还有别的 Edit（如顶部「搜索」框）也不会选错。

    找不到标签时依次兜底：
      1) 父级是 ComboBox / ComboBoxEx32 的 Edit（取最靠下的）；
      2) 任意非空 Edit（取最靠下的）；
      3) 第一个 Edit。
    """
    if not dlg_hwnd:
        return 0, ""

    edits = []
    for h, cls, txt in _win_all_descendants(dlg_hwnd):
        if cls != "Edit":
            continue
        # 跨进程读 Edit 文本要用 WM_GETTEXT（GetWindowTextW 读不到别的进程的控件）
        real = _win_get_text_msg(h) or txt or ""
        rect = _win_get_rect(h)
        parent = 0
        try:
            # GetParent 的 restype 是 wintypes.HWND（c_void_p），
            # 直接 int() 会抛 ValueError，必须取 .value
            praw = _user32.GetParent(wintypes.HWND(h))
            parent = int(getattr(praw, "value", 0) or 0)
        except Exception:
            parent = 0
        edits.append({
            "hwnd": h,
            "txt": real.strip(),
            "left": rect[0], "top": rect[1], "right": rect[2],
            "pcls": _win_class(parent) if parent else "",
        })
    if not edits:
        return 0, ""

    # 1) 用「文件夹:」Static 标签定位
    labels = []
    for h, cls, txt in _win_all_descendants(dlg_hwnd):
        if cls != "Static":
            continue
        label = _win_get_text_msg(h) or txt or ""
        if "文件夹" not in label:
            continue
        rect = _win_get_rect(h)
        labels.append({"left": rect[0], "top": rect[1], "right": rect[2]})

    best = None
    for lb in labels:
        for e in edits:
            if e["left"] < lb["left"]:
                continue                      # 必须在标签右侧
            if abs(e["top"] - lb["top"]) >= 40:
                continue                      # 垂直位置必须接近
            gap = max(0, e["left"] - lb["right"])
            key = (gap, abs(e["top"] - lb["top"]))
            if best is None or key < best[0]:
                best = (key, e)
    if best is not None:
        return best[1]["hwnd"], best[1]["txt"]

    # 2) 兜底：父级是 ComboBox / ComboBoxEx32 的 Edit（取最靠下的）
    combos = [e for e in edits if e["pcls"] in ("ComboBox", "ComboBoxEx32")]
    if combos:
        combos.sort(key=lambda e: -e["top"])
        return combos[0]["hwnd"], combos[0]["txt"]

    # 3) 兜底：任意非空 Edit（取最靠下的）
    nonempty = [e for e in edits if e["txt"]]
    if nonempty:
        nonempty.sort(key=lambda e: -e["top"])
        return nonempty[0]["hwnd"], nonempty[0]["txt"]

    # 4) 最后兜底：第一个 Edit
    return edits[0]["hwnd"], edits[0]["txt"]


def _win_folder_dialog_pick(dlg, text, timeout=ELEC_FOLDER_PICK_TIMEOUT):
    """在「选择文件夹」对话框里写入路径并点击「选择文件夹」。

    Args:
        dlg: 对话框窗口句柄。
        text: 要写进「文件夹:」输入框的内容（一般是目录路径）。
        timeout: 点击后等待对话框关闭的最长秒数。

    Returns:
        bool: 对话框已关闭返回 True，否则返回 False。
    """
    if not dlg:
        print("[!] 未传入「选择文件夹」对话框句柄。")
        return False

    edit, cur = _win_find_folder_name_edit(dlg)
    if not edit:
        print("[!] 未找到「文件夹:」输入框。")
        return False
    print("[*] 「文件夹:」输入框 hwnd=" + str(edit) + "，当前内容: " + (cur or "(空)"))

    if not _win_set_text(edit, text):
        print("[!] 写入「文件夹:」输入框失败: " + str(text))
        return False
    time.sleep(0.3)

    # 回读校验：外壳有时会把内容规范化显示，不一致只告警、不中断
    got = _win_get_text_msg(edit) or ""
    if got.strip() != str(text).strip():
        print("[*] 「文件夹:」输入框回读为: " + (got or "(空)")
              + "（期望: " + str(text) + "）")

    btn, _btxt = _win_find_button(dlg, ("选择文件夹", "Select Folder"))
    if not btn:
        # 兜底：老式「浏览文件夹」对话框的确认按钮叫「选择」/「确定」
        btn, _btxt = _win_find_button(dlg, ("选择", "Select", "确定", "OK"))
    if not btn:
        print("[!] 未找到「选择文件夹」按钮。")
        return False

    print("[*] 点击「选择文件夹」按钮（内容: " + (got.strip() or str(text)) + "） ...")
    _win_click_button(btn)

    deadline = time.time() + max(1, timeout)
    while time.time() < deadline:
        if not _win_find_folder_dialog()[0]:
            print("[+] 「选择文件夹」对话框已关闭。")
            return True
        time.sleep(0.3)

    print("[!] 点击「选择文件夹」后对话框仍未关闭。")
    return False


def _step_click_select_electronic_folder():
    """点击页面第 ② 步的「选择报告电子原件文件夹导入」/「重新选择」按钮。

    只负责把原生「选择文件夹」对话框点出来，具体选哪个目录由下一步完成。

    找不到按钮即终止流程（返回 False）：对话框弹不出来就无法选中 OrigDoc，
    继续走下去只会上传不完整的电子原件。
    """
    assert cdp_client is not None, "cdp_client 未初始化"

    deadline = time.time() + ELEC_FOLDER_BTN_TIMEOUT
    res = None
    while time.time() < deadline:
        try:
            res = cdp_client.eval_js(JS_CLICK_SELECT_ELECTRONIC_FOLDER)
        except Exception as e:
            print("[!] 定位「选择报告电子原件文件夹导入」按钮异常: " + repr(e))
            res = None
        if res and res.get("ok") and res.get("x"):
            break
        if res and res.get("msg") and "不可用" in str(res.get("msg")):
            print("[!] " + str(res.get("msg")))
            break
        time.sleep(0.5)

    if not res or not res.get("ok"):
        print("[!] 未找到「选择报告电子原件文件夹导入」/「重新选择」按钮，终止流程: "
              + str((res or {}).get("msg")))
        return False

    btn_name = "重新选择" if res.get("kind") == "reselect" else "选择报告电子原件文件夹导入"
    print("[*] 模拟真实点击「%s」（%s，%s，坐标 %d,%d） ..."
          % (btn_name, str(res.get("tag")), str(res.get("cls")),
             int(res.get("x", 0)), int(res.get("y", 0))))
    _cdp_click_at(cdp_client, int(res.get("x", 0)), int(res.get("y", 0)))
    time.sleep(1)
    print("[+] 已点击「%s」，等待原生「选择文件夹」对话框弹出 ..." % btn_name)
    return True


def _step_select_origdoc_folder():
    """在原生「选择文件夹」对话框中选中 OrigDoc 目录并点「选择文件夹」。

    目标目录：%ProgramData%\\AreportFile\\OrigDoc（上一环节刚生成的 PDF 所在处）。

    选目录的方案（按顺序尝试）：
      A) 往「文件夹:」框里写完整路径 → 点「选择文件夹」；
         —— 直接告诉外壳「我要这个目录」，成功率最高。
      B) A 没能让对话框关闭（外壳可能拒绝了完整路径）时：
         B1) 写「父目录 + 结尾反斜杠」，让对话框先「进入」该目录；
         B2) 再只写目录名 OrigDoc。
    最后回到页面校验是否显示「已选择文件夹: OrigDoc」。
    """
    assert cdp_client is not None, "cdp_client 未初始化"
    if not _WIN32 or _user32 is None:
        print("[!] 当前环境不支持原生对话框自动化，无法选择文件夹。")
        return False

    program_data = os.environ.get("ProgramData") or r"C:\ProgramData"
    target = os.path.join(program_data, "AreportFile", "OrigDoc")
    parent = os.path.dirname(target)
    if not os.path.isdir(target):
        print("[!] OrigDoc 目录不存在，无法选择: " + target)
        return False
    print("[*] 目标文件夹: " + target)

    # 记录操作前的页面状态：原生对话框弹出时渲染进程会被阻塞，
    # 这里多半读不到（before_name 为空），拿它做「页面是否真的变了」的对比基准
    before = _eval_js_quick(cdp_client, JS_CHECK_SELECTED_FOLDER, timeout=3) or {}
    before_name = str(before.get("name") or "")

    # 1) 等待原生「选择文件夹」对话框弹出
    dlg, title, btn_text = 0, "", ""
    deadline = time.time() + ELEC_FOLDER_DIALOG_TIMEOUT
    while time.time() < deadline:
        dlg, title, btn_text = _win_find_folder_dialog()
        if dlg:
            break
        time.sleep(0.3)
    if not dlg:
        print("[!] 等待「选择文件夹」对话框超时（%d 秒）。" % ELEC_FOLDER_DIALOG_TIMEOUT)
        return False
    print("[+] 已定位「选择文件夹」对话框：标题=" + title + "，按钮=" + btn_text)

    # 2) 方案 A：直接写完整路径
    closed = _win_folder_dialog_pick(dlg, target)

    # 3) 方案 A 未让对话框关闭 → 清掉残留错误框后改用方案 B
    if not closed:
        _win_dismiss_error_boxes()
        dlg2, _t2, _b2 = _win_find_folder_dialog()
        if dlg2:
            print("[*] 方案 A 未生效，改用方案 B1：先进入父目录 " + parent + os.sep)
            closed = _win_folder_dialog_pick(dlg2, parent + os.sep)
        if not closed:
            dlg3, _t3, _b3 = _win_find_folder_dialog()
            if dlg3:
                print("[*] 方案 B1 未生效，改用方案 B2：只写目录名 "
                      + os.path.basename(target))
                closed = _win_folder_dialog_pick(dlg3, os.path.basename(target))

    if not closed:
        print("[!] 无法通过「选择文件夹」对话框选中 OrigDoc 目录。")
        return False

    # 4) 回页面校验：等页面刷新出「已选择文件夹」
    deadline = time.time() + ELEC_FOLDER_VERIFY_TIMEOUT
    info = {}
    while time.time() < deadline:
        info = _eval_js_quick(cdp_client, JS_CHECK_SELECTED_FOLDER, timeout=5) or {}
        name = str(info.get("name") or "")
        if info.get("hasOrigDoc"):
            print("[+] 页面已显示已选择的文件夹: %s（%d 个文件）"
                  % (name or os.path.basename(target), int(info.get("fileCount") or 0)))
            return True
        if name and name != before_name:
            # 页面已经刷成了别的目录名 → 选错了
            print("[!] 选中的文件夹不是 OrigDoc（当前: " + name + "）")
            return False
        time.sleep(0.5)

    name = str(info.get("name") or "")
    if name and not info.get("hasOrigDoc"):
        print("[!] 选中的文件夹不是 OrigDoc（当前: " + name + "）")
        return False
    print("[*] 未在页面上确认到文件夹信息（流程继续）。")
    return True


# ===================== 点击「数据上传」按钮 =====================
# 业务背景：3 步向导的最后一步（③ 数据上传），页面中部有一个蓝色
# 「数据上传」按钮。顶部步骤条第 ③ 项文案也含「数据上传」，
# 所以定位时要求元素的归一化文本「精确等于」数据上传（步骤条文案
# 带序号「3数据上传」，会被排除），并优先取真正的按钮元素。
UPLOAD_BTN_TIMEOUT = 15      # 等待页面上「数据上传」按钮出现的最长秒数

# 在页面里定位第 ③ 步的蓝色「数据上传」按钮，返回坐标（不直接点击，
# 交给 CDP 做可信点击）。
JS_CLICK_DATA_UPLOAD = r"""
(function () {
    const KEY = '数据上传';
    function norm(s) { return ((s || '') + '').replace(/\s+/g, ''); }
    function textOf(el) { return norm(el.innerText || el.textContent || ''); }
    function isClickable(el) {
        return /^(BUTTON|A)$/i.test(el.tagName) ||
               el.getAttribute('role') === 'button' ||
               /arco-btn|btn/i.test((el.className || '') + '');
    }
    function depthOf(el) {
        let d = 0, p = el.parentElement;
        while (p && d < 50) { d++; p = p.parentElement; }
        return d;
    }

    const all = Array.from(document.querySelectorAll('button, a, span, div'));
    // 精确匹配：排除步骤条（文案为「3数据上传」这类带序号的形式）
    const hits = all.filter(e => textOf(e) === norm(KEY));
    if (!hits.length) {
        return { ok: false, msg: '未找到「' + KEY + '」按钮' };
    }
    const clickables = hits.filter(isClickable);
    const pool = clickables.length ? clickables : hits;
    pool.sort((a, b) => (depthOf(b) - depthOf(a)));
    const el = pool[0];

    if (el.disabled || el.getAttribute('aria-disabled') === 'true') {
        return { ok: false, msg: '「' + KEY + '」按钮当前不可用' };
    }
    try { el.scrollIntoView({ block: 'center' }); } catch (e) {}
    const r = el.getBoundingClientRect();
    if (!r.width || !r.height) {
        return { ok: false, msg: '「' + KEY + '」按钮不可见' };
    }
    return {
        ok: true,
        x: Math.round(r.x + r.width / 2),
        y: Math.round(r.y + r.height / 2),
        tag: el.tagName,
        cls: (el.className || '') + '',
        text: textOf(el)
    };
})()
"""


def _step_click_data_upload():
    """点击页面第 ③ 步的「数据上传」按钮（向导最后一步）。

    找不到按钮即终止流程（返回 False）。
    注意：本步骤只负责点开「数据上传」确认弹窗；弹窗里的「数据上传」按钮
          由下一步在等待 10 秒后点击（上传后无法撤销）。
    """
    assert cdp_client is not None, "cdp_client 未初始化"

    deadline = time.time() + UPLOAD_BTN_TIMEOUT
    res = None
    while time.time() < deadline:
        try:
            res = cdp_client.eval_js(JS_CLICK_DATA_UPLOAD)
        except Exception as e:
            print("[!] 定位「数据上传」按钮异常: " + repr(e))
            res = None
        if res and res.get("ok") and res.get("x"):
            break
        time.sleep(0.5)

    if not res or not res.get("ok"):
        print("[!] 未找到「数据上传」按钮，终止流程: " + str((res or {}).get("msg")))
        return False

    print("[*] 模拟真实点击「数据上传」（%s，%s，坐标 %d,%d） ..."
          % (str(res.get("tag")), str(res.get("cls")),
             int(res.get("x", 0)), int(res.get("y", 0))))
    _cdp_click_at(cdp_client, int(res.get("x", 0)), int(res.get("y", 0)))
    time.sleep(1)
    print("[+] 已点击「数据上传」按钮。")
    return True


# ============ 等待 10 秒后点击「数据上传」确认弹窗里的按钮 ============
# 点第 ③ 步的「数据上传」后会弹出确认弹窗：
#   标题「数据上传」+ 红字「请确认核对数据，上传后无法撤销」+
#   底部按钮「数据上传（5S）」（带倒计时，倒计时结束后变成「数据上传」）。
# 这里先固定等待 UPLOAD_CONFIRM_DELAY 秒，再用 CDP 做一次可信点击提交上传。
UPLOAD_CONFIRM_DELAY = 10           # 弹窗出现后先等待的秒数（等倒计时走完 / 页面稳定）
UPLOAD_CONFIRM_TIMEOUT = 60         # 等待弹窗按钮出现且可用（倒计时结束）的最长秒数
UPLOAD_CONFIRM_CLOSE_TIMEOUT = 5    # 点击后等待弹窗关闭的最长秒数（仅做日志提示）

# 只在可见的确认弹窗范围内找按钮，避免误点到被遮住的页面第 ③ 步「数据上传」按钮。
JS_FIND_UPLOAD_CONFIRM_BTN = r"""
(function () {
    function norm(s) { return ((s || '') + '').replace(/\s+/g, ''); }
    function visible(el) {
        const r = el.getBoundingClientRect();
        if (r.width <= 0 || r.height <= 0) return false;
        const st = window.getComputedStyle(el);
        return st.display !== 'none' && st.visibility !== 'hidden' &&
               parseFloat(st.opacity || '1') > 0;
    }
    function depthOf(el) {
        let d = 0, p = el.parentElement;
        while (p && d < 50) { d++; p = p.parentElement; }
        return d;
    }
    function pickFrom(scope) {
        const out = [];
        const els = Array.from(scope.querySelectorAll('button, a, [role="button"], .arco-btn'));
        for (const b of els) {
            if (!visible(b)) continue;
            const t = norm(b.innerText || b.textContent || '');
            // 文案形如：数据上传 / 数据上传（5S） / 数据上传(55)
            if (!/^数据上传([（(][0-9a-zA-Z秒]*[）)])?$/.test(t)) continue;
            const r = b.getBoundingClientRect();
            const cls = (b.className || '') + '';
            out.push({
                x: Math.round(r.x + r.width / 2),
                y: Math.round(r.y + r.height / 2),
                text: t,
                tag: b.tagName,
                cls: cls,
                countdown: /[（(]/.test(t),
                disabled: !!(b.disabled || b.getAttribute('aria-disabled') === 'true' ||
                             /disabled/i.test(cls)),
                depth: depthOf(b),
                rank: /^(BUTTON|A)$/i.test(b.tagName) ? 0 : 1
            });
        }
        return out;
    }

    let hits = [];
    // 1) 常规：Arco 弹窗容器（可能同时命中 .arco-modal 和 .arco-modal-wrapper）
    const scopes = Array.from(document.querySelectorAll(
        '.arco-modal, .arco-modal-wrapper, [role="dialog"]')).filter(visible);
    for (const scope of scopes) {
        const txt = scope.innerText || '';
        if (txt.indexOf('数据上传') < 0) continue;
        hits = hits.concat(pickFrom(scope));
    }

    // 2) 兜底：弹窗不是 Arco 时，取「包含红色提示文案的最小可见容器」
    if (!hits.length) {
        let box = null;
        for (const d of document.querySelectorAll('div, section')) {
            if (!visible(d)) continue;
            const t = d.innerText || '';
            if (t.indexOf('上传后无法撤销') < 0 && t.indexOf('请确认核对数据') < 0) continue;
            const r = d.getBoundingClientRect();
            const area = r.width * r.height;
            if (!box || area < box.area) box = { el: d, area: area };
        }
        if (box) {
            const got = pickFrom(box.el);
            const cd = got.filter(h => h.countdown);
            hits = cd.length ? cd : got;   // 优先带倒计时的那个，排除页面上的同名按钮
        }
    }

    if (!hits.length) {
        return { ok: false, msg: '未找到弹窗里的「数据上传」按钮' };
    }

    // 同一按钮可能被多个弹窗容器重复收录，按坐标去重
    const seen = {}, uniq = [];
    for (const h of hits) {
        const k = h.x + ',' + h.y;
        if (seen[k]) continue;
        seen[k] = 1;
        uniq.push(h);
    }
    uniq.sort((a, b) => (a.disabled ? 1 : 0) - (b.disabled ? 1 : 0) ||
                        a.rank - b.rank || b.depth - a.depth ||
                        a.text.length - b.text.length);
    return {
        ok: true,
        x: uniq[0].x,
        y: uniq[0].y,
        text: uniq[0].text,
        tag: uniq[0].tag,
        cls: uniq[0].cls,
        disabled: uniq[0].disabled
    };
})()
"""

# 只用于点击后的提示性校验：确认弹窗是否还开着（不做任何重试）。
# 注意：Arco 弹窗关闭后 DOM 里的文字仍在，所以必须用「可见性」判断。
JS_CHECK_UPLOAD_CONFIRM_MODAL = r"""
(function () {
    const scopes = Array.from(document.querySelectorAll(
        '.arco-modal, .arco-modal-wrapper, [role="dialog"]'));
    for (const scope of scopes) {
        const r = scope.getBoundingClientRect();
        if (r.width <= 0 || r.height <= 0) continue;
        const st = window.getComputedStyle(scope);
        if (st.display === 'none' || st.visibility === 'hidden') continue;
        const t = scope.innerText || '';
        if (t.indexOf('上传后无法撤销') >= 0 || t.indexOf('请确认核对数据') >= 0) {
            return { open: true };
        }
    }
    return { open: false };
})()
"""


def _step_click_upload_confirm_modal():
    """流程最后一步：等待 10 秒后，再次模拟点击弹窗里的「数据上传」按钮。

    上一步点第 ③ 步的「数据上传」只是弹出确认弹窗，本步骤才真正提交上传。
    按钮带倒计时（如「数据上传（5S）」）时，等待时间结束后再校验它是否已
    可用；不可用则继续等到倒计时结束。

    重要：上传后无法撤销，因此这里只点一次，绝不做「没反应就重试」。
    """
    assert cdp_client is not None, "cdp_client 未初始化"

    print("[*] 已点开「数据上传」确认弹窗，先等待 %d 秒 ..." % UPLOAD_CONFIRM_DELAY)
    time.sleep(UPLOAD_CONFIRM_DELAY)

    deadline = time.time() + UPLOAD_CONFIRM_TIMEOUT
    not_found_since = time.time()
    found_once = False
    last_log = 0.0
    res = None
    while time.time() < deadline:
        try:
            res = cdp_client.eval_js(JS_FIND_UPLOAD_CONFIRM_BTN)
        except Exception as e:
            print("[!] 定位弹窗「数据上传」按钮异常: " + repr(e))
            res = None
        if res and res.get("ok"):
            found_once = True
            if not res.get("disabled"):
                break
            if time.time() - last_log >= 5:
                print("[*] 弹窗按钮仍在倒计时/禁用中（%s），继续等待 ..."
                      % str(res.get("text")))
                last_log = time.time()
        elif not found_once and time.time() - not_found_since >= 15:
            break                     # 弹窗一直没出来，没必要等满 60 秒
        time.sleep(0.5)

    if not res or not res.get("ok"):
        print("[!] 未找到弹窗里的「数据上传」按钮，终止流程: "
              + str((res or {}).get("msg")))
        return False
    if res.get("disabled"):
        print("[!] 弹窗「数据上传」按钮始终禁用（%s），未点击，终止流程。"
              % str(res.get("text")))
        return False

    print("[*] 等待结束，模拟真实点击弹窗里的「%s」（%s，%s，坐标 %d,%d） ..."
          % (str(res.get("text")), str(res.get("tag")), str(res.get("cls")),
             int(res.get("x", 0)), int(res.get("y", 0))))
    if not _cdp_click_at(cdp_client, int(res.get("x", 0)), int(res.get("y", 0))):
        print("[!] 模拟点击弹窗「数据上传」失败，终止流程。")
        return False
    print("[+] 已点击弹窗里的「数据上传」按钮（不重试，避免重复上传）。")

    # 只做提示性校验：等弹窗关闭；没关也不再点第二次
    close_deadline = time.time() + UPLOAD_CONFIRM_CLOSE_TIMEOUT
    while time.time() < close_deadline:
        time.sleep(0.5)
        info = _eval_js_quick(cdp_client, JS_CHECK_UPLOAD_CONFIRM_MODAL,
                              timeout=3, quiet=True)
        if info is not None and not info.get("open"):
            print("[+] 「数据上传」确认弹窗已关闭。")
            return True
    print("[*] 未确认弹窗是否关闭（可能正在上传中），流程结束，不再重复点击。")
    return True


# ===================== 全部流程结束后：回写 hashReport 上传状态 =====================
# 新后端接口：POST /HashReport/MarkHashReportUploaded
# 只把当前状态=1（要上传）的分组改为 2（已上传），重复调用幂等。
MARK_UPLOADED_BATCH = 500


def _collect_mark_uploaded_items():
    """从 JSONL 中转文件里取出本轮实际写入模板的 (AutoID, 分组) 组合。"""
    path = _rows_file_path()
    if not os.path.isfile(path):
        return []

    items = {}
    for rec in _iter_jsonl(path):
        try:
            auto_id = int(float(rec.get("auto_id")))
        except (TypeError, ValueError):
            continue
        group = str(rec.get("group") or "").strip().lower()
        if auto_id <= 0 or group not in ("ft", "lt", "qj"):
            continue
        items.setdefault(auto_id, set()).add(group)

    return [{"AutoID": k, "Groups": sorted(v)} for k, v in items.items()]


def _step_mark_hash_report_uploaded():
    """全部流程执行完毕后，把本轮报告在 hashReport 中改回「已上传」(2)。

    作为最后一个步骤执行；仍必须保证在 run_login_flow 的 finally 删除
    JSONL 中转文件之前运行。
    """
    items = _collect_mark_uploaded_items()
    if not items:
        print("[!] 没有可回写的报告数据（JSONL 为空或缺少 auto_id / group），跳过。")
        return True

    print("[*] 开始回写 hashReport 上传状态：%d 个订单，每批 %d ..."
          % (len(items), MARK_UPLOADED_BATCH))

    total_updated = 0
    for start in range(0, len(items), MARK_UPLOADED_BATCH):
        batch = items[start:start + MARK_UPLOADED_BATCH]
        body = {
            "Items": batch,
            "OnlyPending": True,
            "PublishEvent": True,
        }

        payload = None
        for attempt in (1, 2):
            token = _hash_report_token(force_refresh=(attempt == 2))
            headers = {
                "Authorization": token,
                "Content-Type": "application/json",
            }
            resp = None
            try:
                resp = requests.post(
                    HASH_REPORT_MARK_UPLOADED_API,
                    json=body,
                    headers=headers,
                    timeout=60,
                )
                resp.raise_for_status()
                payload = _parse_hash_api_payload(resp)
            except Exception as e:
                raise RuntimeError(
                    "回写上传状态第 %d 批失败: %s"
                    % (start // MARK_UPLOADED_BATCH + 1, repr(e))
                )
            finally:
                if resp is not None:
                    try:
                        resp.close()
                    except Exception:
                        pass

            if payload.get("code") == 401:
                if attempt == 1:
                    print("[*] 回写 token 已失效，重新登录后重试")
                    continue
                raise RuntimeError("回写上传状态鉴权失败")
            if payload.get("code") != 200:
                raise RuntimeError("回写上传状态失败: %s" % str(payload)[:300])

            data = payload.get("data") or {}
            updated = int(data.get("Updated") or 0)
            total_updated += updated
            print("[+] 回写批次 %d：请求 %d，更新 %d（ft=%s lt=%s qj=%s），未找到 %s"
                  % (start // MARK_UPLOADED_BATCH + 1,
                     data.get("Requested", 0),
                     updated,
                     (data.get("Groups") or {}).get("ft", 0),
                     (data.get("Groups") or {}).get("lt", 0),
                     (data.get("Groups") or {}).get("qj", 0),
                     data.get("NotFound") or []))
            break

    print("[+] hashReport 上传状态回写完成，共更新 %d 个分组。" % total_updated)
    return True


# ===================== 登录后的后续逻辑 =====================
# [已注释] 原「重载代码」按钮的行为：
#   1) 重新读取并加载本文件的最新代码（热重载，运行期状态全部保留）；
#   2) 只执行本节定义的「登录后逻辑」（run_after_login），
#      不会重新打开应用、也不会重新走登录流程。
#   ——该功能已停用，相关实现见下方「[已注释]」代码块。
#
# 登录流程结束后也会自动执行这些步骤（见 run_login_flow）。
#
# 步骤写成 (名称, 函数) 的形式，日志里会显示「第几步、这一步叫什么」。

POST_LOGIN_STEPS = [
    ("点击「下载表格模板」按钮", _step_click_download_template),
    ("在保存对话框中点击「保存」按钮", _step_handle_save_dialog),
    ("定位刚刚下载好的 .xlsx 文件", _step_locate_template_file),
    ("采集「已上传报告管理 / 全部报告」所有分页的报告编号并切回上传菜单",
     _step_collect_uploaded_report_nos),
    ("在模板 xlsx 中自动填写数据（排除已上传报告编号）", _step_fill_template_data),
    ("用系统默认程序打开该 .xlsx 文件", _step_open_template_file),
    ("等待几秒后关闭该 .xlsx 文件", _step_close_xlsx_file),
]

# 「选择数据表格导入」相关步骤（可用配置项 AUTO_IMPORT_TEMPLATE 关闭）
IMPORT_STEPS = [
    ("点击「选择数据表格导入」按钮", _step_click_import_table),
    ("在「打开」对话框中选中该 .xlsx 并点「打开」", _step_select_file_in_open_dialog),
    ("处理导入后的「编号失效确认」弹窗", _step_handle_invalid_confirm_modal),
    ("按报告编号生成 OrigDoc PDF 文件", _step_generate_orig_pdfs),
    ("确认页面上已完成数据表格导入", _step_check_import_result),
    ("点击「选择报告电子原件文件夹导入」按钮", _step_click_select_electronic_folder),
    ("在文件夹对话框中选中 OrigDoc 并点「选择文件夹」", _step_select_origdoc_folder),
    ("点击「数据上传」按钮", _step_click_data_upload),
    ("等待 %d 秒后点击弹窗里的「数据上传」确认按钮" % UPLOAD_CONFIRM_DELAY,
     _step_click_upload_confirm_modal),
    ("全部流程结束后，把 hashReport 状态回写为「已上传」",
     _step_mark_hash_report_uploaded),
]


def _post_login_steps():
    """完整的登录后步骤（含是否导入数据表格的开关）。"""
    steps = list(POST_LOGIN_STEPS)
    if CONFIG.get("AUTO_IMPORT_TEMPLATE", True):
        steps += list(IMPORT_STEPS)
    return steps


def run_after_login():
    """登录后后续逻辑入口（单独执行 POST_LOGIN_STEPS）。

    [已注释] 原热重载说明：热重载后会执行本函数的「最新版本」，
    因此可以边改代码边点「重载代码」验证（该功能已停用）。
    「重载代码」停用后本函数保留备用，可由自定义线程直接调用。
    """
    steps = _post_login_steps()
    flow_begin("登录后流程", [name for name, _ in steps])
    for name, fn in steps:
        _wait_until_resumed()
        res = run_step(name, fn)
        if res is FLOW_STOP:
            flow_summary(True, "本轮无需上传，流程正常结束")
            print("[+] 本轮没有需要上传的数据，后续步骤已跳过。")
            return True
        if not res:
            raise RuntimeError("后续步骤执行失败: " + name)
    flow_summary(True)
    print("[+] 登录后的后续逻辑执行完毕。")
    return True


# ---------------------------------------------------------------------
# [已注释] 以下为「重载代码」功能（热重载）相关实现，暂不启用。
# 取消注释即可恢复热重载能力（同时需要恢复 GUI 中的「重载代码」按钮
# 与 on_reload / run_post_login_flow）。
# ---------------------------------------------------------------------
# def _hot_reload_source_path() -> str:
#     """定位用于热重载的源码文件。"""
#     env = os.environ.get("HOT_RELOAD_SOURCE")
#     if env and os.path.isfile(env):
#         return os.path.abspath(env)
#     if getattr(sys, "frozen", False):
#         cand = os.path.join(
#             os.path.dirname(sys.executable),
#             os.path.splitext(os.path.basename(sys.executable))[0] + ".py",
#         )
#         if os.path.isfile(cand):
#             return cand
#         raise RuntimeError(
#             "打包成 exe 后找不到源码文件：请把同名 .py 源码放到 exe 同目录，"
#             "或设置环境变量 HOT_RELOAD_SOURCE 指向源码路径。"
#         )
#     return os.path.abspath(__file__)
#
#
# def hot_reload_module() -> None:
#     """热重载本文件的最新代码。
#
#     实现方式：重新读取源码并 exec 到当前模块的命名空间中，因此新增/修改的
#     函数、JS 脚本、常量立即生效；同时保留运行期状态
#     （current_state / app_proc / cdp_client / OCR 引擎 / CONFIG），
#     不会重新打开应用，也不会重新登录。
#     """
#     src_path = _hot_reload_source_path()
#     with open(src_path, "r", encoding="utf-8") as f:
#         source = f.read()
#
#     # 先编译：语法错误时不会污染当前正在运行的代码
#     code = compile(source, src_path, "exec")
#
#     g = globals()
#     keep = {k: g.get(k) for k in ("current_state", "app_proc", "cdp_client", "_ocr")}
#     old_cfg = g.get("CONFIG")
#     old_name = g.get("__name__")
#
#     g["__name__"] = "__hot_reload__"      # 防止再次触发 main()
#     try:
#         exec(code, g)
#     finally:
#         g["__name__"] = old_name
#         for k, v in keep.items():
#             g[k] = v
#         new_cfg = g.get("CONFIG")
#         if isinstance(new_cfg, dict) and isinstance(old_cfg, dict):
#             merged = dict(new_cfg)
#             merged.update(old_cfg)         # 运行期配置优先，同时保留新代码新增的键
#             g["CONFIG"] = merged
#
#     print("[+] 代码已重载: " + src_path)
# =====================================================================


def _ensure_cdp_client():
    """确保 CDP 连接可用（登录后的逻辑需要操作应用页面）。

    已连接则直接复用；未连接但应用仍在运行时，重新连接页面 target。
    """
    global cdp_client
    if cdp_client is not None:
        return True
    if not is_app_running(use_cache=False):
        print("[!] 应用未在运行，无法建立 CDP 连接。")
        return False
    try:
        target = find_login_page(prefer_logged_in=True)
        cdp_client = CDPClient(target["webSocketDebuggerUrl"])
    except Exception as e:
        print("[!] 建立 CDP 连接失败: " + repr(e))
        return False
    print("[+] 已重新连接页面: " + str(target.get("url", "")))
    return True


# ---------------------------------------------------------------------
# [已注释] 「重载代码」功能专用：后台线程执行登录后的后续逻辑。
# 取消注释即可恢复（同时需要恢复 hot_reload_module 与 GUI 的「重载代码」按钮）。
# ---------------------------------------------------------------------
# def run_post_login_flow():
#     """后台线程：执行登录后的后续逻辑（使用刚重载的最新代码）。"""
#     set_state(RUNNING)
#     try:
#         _ensure_cdp_client()
#         run_after_login()
#         set_state(FINISHED)
#     except Exception as e:
#         set_state(ERROR)
#         print("[!] 登录后后续逻辑执行出错: " + repr(e))
# =====================================================================


# ===================== GUI 日志管道（线程安全） =====================
# 旧实现里，后台线程的每次 print() 都会调用 widget.after(0, ...) 把日志塞进
# Tk 事件队列。Tk/Tcl 解释器只允许主线程访问，工作线程并发调用 after() 轻则
# 抛 "main thread is not in main loop"，重则整个窗口进程直接崩溃；下载几百份
# PDF 时进度/失败日志刷屏，风险被放大，日志框还会无限增长直到界面卡死。
#
# 现在：后台线程只往线程安全队列里投「整行文本」，绝不接触 tkinter；
#       主线程每 _GUI_LOG_POLL_MS 毫秒批量取出、一次 insert 写入日志框，
#       并裁掉过旧的行；进行中的进度走单独通道，始终只显示一条。
_GUI_LOG_QUEUE = queue.Queue()
_GUI_LOG_BUFFER = []                 # print 的一次输出会被拆成多次 write，先拼成整行
_GUI_LOG_LOCK = threading.Lock()
_GUI_LOG_POLL_MS = 100               # 主线程批量刷新日志的间隔
_GUI_LOG_MAX_LINES = 2000            # 日志框最多保留的行数（防止无限增长）
# 用字典存状态（与 FLOW_STATE / _APP_RUNNING_CACHE 同风格），避免 global 重赋值
_GUI_LOG_STATE = {"alive": True}     # 窗口关闭后后台线程不再堆积日志


def _gui_log_line(text):
    """写一条普通日志（任意线程可调用，只入队、不碰 Tk 控件）。"""
    if text and _GUI_LOG_STATE["alive"]:
        _GUI_LOG_QUEUE.put(("line", text))


def set_progress(text):
    """刷新界面上唯一的一条进度；text 为空表示清除。

    长任务（如批量下载 PDF）用它显示实时进度：刷新多少次都只占一行，
    不会往日志框追加内容，也就不会把 GUI 刷爆。
    """
    if _GUI_LOG_STATE["alive"]:
        _GUI_LOG_QUEUE.put(("progress", text))


class _GuiStream:
    """把 print()/traceback 重定向到 GUI 日志框（线程安全）。

    write() 只做行缓冲 + 入队，不调用任何 tkinter 方法；
    真正的控件写入统一由 AutoLoginApp._poll_log_queue() 在主线程完成。
    """
    def __init__(self, widget=None):
        self.widget = widget

    def write(self, text):
        if not text or not _GUI_LOG_STATE["alive"]:
            return
        with _GUI_LOG_LOCK:
            _GUI_LOG_BUFFER.append(text)
            if "\n" not in text:
                return
            parts = "".join(_GUI_LOG_BUFFER).split("\n")
            tail = parts.pop()               # 末尾半行留到下一次 write
            _GUI_LOG_BUFFER[:] = [tail] if tail else []
            for line in parts:
                line = line.rstrip("\r")
                if line:
                    _GUI_LOG_QUEUE.put(("line", line))

    def flush(self):
        # 有调用方写完半行就 flush；把残留也投出去，避免丢日志
        with _GUI_LOG_LOCK:
            if _GUI_LOG_BUFFER:
                line = "".join(_GUI_LOG_BUFFER).rstrip("\r\n")
                _GUI_LOG_BUFFER.clear()
                if line and _GUI_LOG_STATE["alive"]:
                    _GUI_LOG_QUEUE.put(("line", line))

    # 兼容第三方库对 sys.stdout / sys.stderr 的常见探测
    encoding = "utf-8"

    def isatty(self):
        return False

    def writelines(self, lines):
        for line in lines:
            self.write(line)


class _EyeToggle(tk.Frame):
    """密码输入框旁的“眼睛”图标按钮：点击切换明文 / 掩码显示。

    图标用 Canvas 绘制，不依赖任何外部图片资源：
        - 密码隐藏（默认）：灰色带斜杠的眼睛，表示“当前不可见，点击显示”；
        - 密码显示：蓝色睁开的眼睛，表示“当前可见，点击隐藏”。
    """

    ICON_W = 26
    ICON_H = 18

    def __init__(self, master, entry):
        super().__init__(master, bg=master.cget("bg"))
        self.entry = entry
        self.revealed = False
        self.hover = False
        self.canvas = tk.Canvas(
            self, width=self.ICON_W, height=self.ICON_H,
            bg=master.cget("bg"), highlightthickness=0, bd=0, cursor="hand2")
        self.canvas.pack()
        self.canvas.bind("<Button-1>", self._toggle)
        self.canvas.bind("<Enter>", self._on_enter)
        self.canvas.bind("<Leave>", self._on_leave)
        self._draw()

    # ---------- 事件 ----------
    def _toggle(self, _event=None):
        self.revealed = not self.revealed
        self.entry.configure(show="" if self.revealed else "*")
        self._draw()

    def _on_enter(self, _event=None):
        self.hover = True
        self._draw()

    def _on_leave(self, _event=None):
        self.hover = False
        self._draw()

    # ---------- 绘制 ----------
    def _draw(self):
        c = self.canvas
        c.delete("all")
        cx = self.ICON_W / 2
        cy = self.ICON_H / 2
        rx = 10.0
        ry = 6.5
        box = (cx - rx, cy - ry, cx + rx, cy + ry)
        if self.revealed:
            color = "#2f6fdb"
        else:
            color = "#2f6fdb" if self.hover else "#666666"

        # 上下眼睑：两条圆弧拼成杏仁形
        c.create_arc(*box, start=20, extent=140, style=tk.ARC, outline=color, width=2)
        c.create_arc(*box, start=200, extent=140, style=tk.ARC, outline=color, width=2)
        # 瞳孔
        c.create_oval(cx - 2.6, cy - 2.6, cx + 2.6, cy + 2.6, outline=color, width=2)

        if not self.revealed:
            # 隐藏状态：加一条斜杠，表示“密码不可见”
            c.create_line(cx - rx - 1, cy + ry, cx + rx + 1, cy - ry, fill=color, width=2)


class SettingsDialog(tk.Toplevel):
    """软件路径等配置对话框。"""
    def __init__(self, parent_app):
        super().__init__(parent_app.root)
        self.app = parent_app
        self.title("设置")
        # 功能开关已不显示，窗口高度相应缩小
        self.geometry("640x250")
        self.resizable(False, False)
        self.transient(parent_app.root)
        self.grab_set()

        body = tk.Frame(self, padx=14, pady=14)
        body.pack(fill=tk.BOTH, expand=True)

        tk.Label(body, text="软件路径 (exe)：", anchor="w").grid(
            row=0, column=0, sticky="w", pady=4)
        self.path_var = tk.StringVar(value=CONFIG["EXE_PATH"])
        path_entry = tk.Entry(body, textvariable=self.path_var, width=50)
        path_entry.grid(row=0, column=1, padx=(6, 4), pady=4)
        tk.Button(body, text="浏览...", width=8, command=self._browse).grid(
            row=0, column=2, pady=4)

        tk.Label(body, text="用户名：", anchor="w").grid(
            row=1, column=0, sticky="w", pady=4)
        self.user_var = tk.StringVar(value=CONFIG["USERNAME"])
        tk.Entry(body, textvariable=self.user_var, width=50).grid(
            row=1, column=1, columnspan=2, padx=6, pady=4, sticky="w")

        tk.Label(body, text="密码：", anchor="w").grid(
            row=2, column=0, sticky="w", pady=4)
        self.pwd_var = tk.StringVar(value=CONFIG["PASSWORD"])
        self.pwd_entry = tk.Entry(body, textvariable=self.pwd_var, width=50, show="*")
        self.pwd_entry.grid(row=2, column=1, padx=(6, 4), pady=4, sticky="w")
        self.pwd_eye = _EyeToggle(body, self.pwd_entry)
        self.pwd_eye.grid(row=2, column=2, padx=(0, 4), pady=4)

        tk.Label(body, text="调试端口：", anchor="w").grid(
            row=3, column=0, sticky="w", pady=4)
        self.port_var = tk.StringVar(value=str(CONFIG["DEBUG_PORT"]))
        tk.Entry(body, textvariable=self.port_var, width=50).grid(
            row=3, column=1, columnspan=2, padx=6, pady=4, sticky="w")

        # ---------------------------------------------------------------
        # 功能开关不再让用户勾选，统一在后台按默认值启用（保留代码以便需要时恢复）。
        # ---------------------------------------------------------------
        # self.keep_app_var = tk.BooleanVar(value=bool(CONFIG["KEEP_APP_OPEN"]))
        # tk.Checkbutton(
        #     body, text="登录成功后保留应用（不自动关闭）", variable=self.keep_app_var
        # ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(6, 0))
        #
        # self.auto_dl_var = tk.BooleanVar(value=bool(CONFIG.get("AUTO_DOWNLOAD_TEMPLATE", True)))
        # tk.Checkbutton(
        #     body, text="登录进入首页后自动下载表格模板（点模板按钮→保存→打开 xlsx）",
        #     variable=self.auto_dl_var
        # ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(2, 0))
        #
        # self.auto_open_var = tk.BooleanVar(value=bool(CONFIG.get("AUTO_OPEN_TEMPLATE", True)))
        # tk.Checkbutton(
        #     body, text="模板下载完成后自动打开该 .xlsx 文件",
        #     variable=self.auto_open_var
        # ).grid(row=6, column=0, columnspan=3, sticky="w", pady=(2, 0))
        #
        # self.auto_fill_var = tk.BooleanVar(value=bool(CONFIG.get("AUTO_FILL_TEMPLATE", True)))
        # tk.Checkbutton(
        #     body, text="打开该 .xlsx 前自动按规则填写数据（单位名称 / 报告编号 / 日期 等）",
        #     variable=self.auto_fill_var
        # ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(2, 0))
        #
        # self.close_xlsx_var = tk.BooleanVar(value=bool(CONFIG.get("CLOSE_XLSX_AFTER_OPEN", True)))
        # tk.Checkbutton(
        #     body, text="打开该 .xlsx 后等待几秒自动关闭它（释放文件锁）",
        #     variable=self.close_xlsx_var
        # ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(2, 0))
        #
        # self.auto_import_var = tk.BooleanVar(value=bool(CONFIG.get("AUTO_IMPORT_TEMPLATE", True)))
        # tk.Checkbutton(
        #     body, text="关闭后自动点「选择数据表格导入」并选中该 xlsx 点「打开」",
        #     variable=self.auto_import_var
        # ).grid(row=9, column=0, columnspan=3, sticky="w", pady=(2, 0))
        #
        # 后台统一启用的功能（用户无需关心）：
        #   登录后保留应用 / 自动下载模板 / 自动打开 xlsx / 自动填写数据 / 自动关闭 xlsx / 自动导入数据表格
        # 直接取 DEFAULT_CONFIG 中的默认值（当前均为 True），以后调整默认值此处自动跟随。
        self.fixed_options = {
            "keep_app_open": bool(DEFAULT_CONFIG["KEEP_APP_OPEN"]),
            "auto_download_template": bool(DEFAULT_CONFIG["AUTO_DOWNLOAD_TEMPLATE"]),
            "auto_open_template": bool(DEFAULT_CONFIG["AUTO_OPEN_TEMPLATE"]),
            "auto_fill_template": bool(DEFAULT_CONFIG["AUTO_FILL_TEMPLATE"]),
            "close_xlsx_after_open": bool(DEFAULT_CONFIG["CLOSE_XLSX_AFTER_OPEN"]),
            "auto_import_template": bool(DEFAULT_CONFIG["AUTO_IMPORT_TEMPLATE"]),
        }

        btn_f = tk.Frame(body)
        btn_f.grid(row=4, column=0, columnspan=3, pady=(18, 0))
        tk.Button(btn_f, text="保存", width=12, command=self._save).pack(side=tk.LEFT, padx=8)
        tk.Button(btn_f, text="取消", width=12, command=self.destroy).pack(side=tk.LEFT, padx=8)

        self._center()

    def _center(self):
        """把弹窗居中显示在屏幕正中间（高度按实际内容自适应）。"""
        self.update_idletasks()
        w = self.winfo_width() or 640
        if w <= 1:
            w = 640
        # 功能开关已去掉，按控件实际需要的高度收窄窗口，兼容不同 DPI 缩放
        h = max(230, self.winfo_reqheight())
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _browse(self):
        path = filedialog.askopenfilename(
            title="选择上传系统 exe",
            filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")]
        )
        if path:
            self.path_var.set(path)

    def _save(self):
        path = self.path_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先设置软件路径 (exe)，否则无法启动。")
            return
        user = self.user_var.get().strip()
        pwd = self.pwd_var.get()
        try:
            port = int(self.port_var.get().strip())
        except ValueError:
            messagebox.showwarning("提示", "调试端口必须是一个整数。")
            return
        # 功能开关不再由界面提供，统一使用后台固定值（原勾选代码已注释，见上方）。
        # self.app.apply_settings(path, user, pwd, port, self.keep_app_var.get(),
        #                         self.auto_dl_var.get(), self.auto_open_var.get(),
        #                         self.auto_fill_var.get(), self.close_xlsx_var.get(),
        #                         self.auto_import_var.get())
        self.app.apply_settings(
            path, user, pwd, port,
            self.fixed_options["keep_app_open"],
            self.fixed_options["auto_download_template"],
            self.fixed_options["auto_open_template"],
            self.fixed_options["auto_fill_template"],
            self.fixed_options["close_xlsx_after_open"],
            self.fixed_options["auto_import_template"],
        )
        self.destroy()


class AutoLoginApp:
    def __init__(self, root):
        self.root = root
        self.root.title("检验检测报告编号上传自动助手")
        self.root.geometry("760x500")

        self.status_var = tk.StringVar(value="当前状态: 就绪")
        self.progress_var = tk.StringVar(value="")   # 长任务进度，始终只有一条
        self.log_text = scrolledtext.ScrolledText(
            root, height=22, wrap=tk.WORD, font=("Consolas", 10)
        )
        self.log_text.configure(state=tk.DISABLED)

        self._last_state = None  # 用于 _update_ui 仅在状态变化时写日志
        self._build_widgets()

        self._append_log("[*] 应用已启动，当前状态: 就绪。请点击「启动」。")
        if _OCR_AVAILABLE:
            self._append_log("[+] 图形验证码自动识别已启用（ddddocr），将自动填写验证码。")
        else:
            self._append_log("[!] 未安装 ddddocr，图形验证码需要人工输入（pip install ddddocr）。")
        self.refresh_exe_path_label()
        if not is_exe_path_set():
            self._append_log("[!] 当前未配置软件路径，请点击「设置」选择上传系统 exe 后再启动。")

        self._center_root()

    def _center_root(self):
        """把主窗口居中显示在屏幕正中间。"""
        self.root.update_idletasks()
        w = self.root.winfo_width() or 760
        h = self.root.winfo_height() or 500
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _build_widgets(self):
        # 标题
        title = tk.Label(
            self.root,
            text="检验检测报告编号上传系统自动化助手",
            font=("Microsoft YaHei", 14, "bold")
        )
        title.pack(pady=(10, 5))

        # 状态
        status_label = tk.Label(
            self.root,
            textvariable=self.status_var,
            font=("Microsoft YaHei", 12, "bold")
        )
        status_label.pack(pady=(0, 5))

        # 当前软件路径显示
        self.exe_path_var = tk.StringVar(value="软件路径: （未设置，请先点「设置」）")
        self.exe_path_label = tk.Label(
            self.root, textvariable=self.exe_path_var,
            font=("Microsoft YaHei", 10), fg="#AA5500"
        )
        self.exe_path_label.pack(pady=(0, 5))

        # 流程进度（显示当前进行到第几步）
        self.flow_var = tk.StringVar(value="流程进度: 尚未开始")
        self.flow_label = tk.Label(
            self.root, textvariable=self.flow_var,
            font=("Microsoft YaHei", 10), fg="#1f5faa"
        )
        self.flow_label.pack(pady=(0, 5))

        # 按钮
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=8)

        self.start_btn = tk.Button(
            btn_frame, text="启动", width=12, command=self.on_start
        )
        self.start_btn.grid(row=0, column=0, padx=6)

        self.pause_btn = tk.Button(
            btn_frame, text="暂停", width=12, command=self.on_pause, state=tk.DISABLED
        )
        self.pause_btn.grid(row=0, column=1, padx=6)

        self.continue_btn = tk.Button(
            btn_frame, text="继续", width=12, command=self.on_continue, state=tk.DISABLED
        )
        self.continue_btn.grid(row=0, column=2, padx=6)

        self.retry_btn = tk.Button(
            btn_frame, text="重新执行", width=12, command=self.on_retry, state=tk.DISABLED
        )
        self.retry_btn.grid(row=0, column=3, padx=6)

        # [已注释] 「重载代码」按钮（热重载功能已停用，取消注释可恢复）
        # self.reload_btn = tk.Button(
        #     btn_frame, text="重载代码", width=12, command=self.on_reload
        # )
        # self.reload_btn.grid(row=0, column=4, padx=6)

        self.settings_btn = tk.Button(
            btn_frame, text="设置", width=12, command=self.on_settings
        )
        self.settings_btn.grid(row=0, column=5, padx=6)

        # 长任务进度：独立一行原地刷新，不往日志框追加行
        self.progress_label = tk.Label(
            self.root, textvariable=self.progress_var, anchor="w",
            font=("Consolas", 10), fg="#1f5faa"
        )
        self.progress_label.pack(fill=tk.X, padx=8, pady=(4, 0))

        # 日志
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

    def _append_log(self, message):
        # 统一走日志队列，保证所有日志按同一顺序在一个地方写入控件
        _gui_log_line(message)

    def _poll_log_queue(self):
        """主线程：批量把后台线程的日志写入控件（每 100ms 一次）。"""
        try:
            if not self.root.winfo_exists():
                return
        except tk.TclError:
            return

        lines, progress = [], None
        for _ in range(1000):        # 单批上限，避免日志洪峰时长时间占用主线程
            try:
                kind, text = _GUI_LOG_QUEUE.get_nowait()
            except queue.Empty:
                break
            if kind == "progress":
                progress = text      # 同一批里只保留最后一条进度
            elif text:
                lines.append(text)

        try:
            if progress is not None:
                self.progress_var.set(progress)
            if lines:
                # 用户手动往上翻看日志时，不强制拉回底部
                at_bottom = self.log_text.yview()[1] >= 0.99
                self.log_text.configure(state=tk.NORMAL)
                self.log_text.insert(tk.END, "\n".join(lines) + "\n")
                self._trim_log()
                if at_bottom:
                    self.log_text.see(tk.END)
                self.log_text.configure(state=tk.DISABLED)
        except tk.TclError:
            return                   # 窗口正在销毁：停止刷新，不把异常抛进 Tk
        except Exception:
            # 单批日志出问题也不能中断整个刷新循环
            try:
                self.log_text.configure(state=tk.DISABLED)
            except Exception:
                pass
        self.root.after(_GUI_LOG_POLL_MS, self._poll_log_queue)

    def _trim_log(self):
        """日志框只保留最近 _GUI_LOG_MAX_LINES 行，避免长时间运行后无限膨胀。"""
        try:
            last = int(self.log_text.index("end-1c").split(".")[0])
        except Exception:
            return
        if last > _GUI_LOG_MAX_LINES:
            self.log_text.delete("1.0", "%d.0" % (last - _GUI_LOG_MAX_LINES + 1))

    def _update_ui(self):
        state = current_state

        # 流程进度：显示当前进行到第几步
        idx = FLOW_STATE.get("index", 0)
        if idx:
            total = max(FLOW_STATE.get("total", 0), idx)
            self.flow_var.set("流程进度: 第 %d/%d 步 — %s" % (
                idx, total, FLOW_STATE.get("current", "")))
        else:
            self.flow_var.set("流程进度: 尚未开始")

        # 仅在状态变化时记录一次日志，避免每 250ms 刷屏
        if state != self._last_state:
            self._last_state = state
            if state in (IDLE, FINISHED, ERROR):
                set_progress("")   # 流程结束 / 出错 / 重置时清掉残留的进度行
            if state == IDLE:
                self._append_log("[*] 当前状态: 就绪。等待点击「启动」。")
            elif state == RUNNING:
                self._append_log("[+] 当前状态: 运行中")
            elif state == PAUSED:
                self._append_log("[!] 当前状态: 已暂停。点击「继续」恢复。")
            elif state == STATE_WAIT_CAPTCHA:
                self._append_log("[!] 当前状态: 等待你输入验证码。请在应用窗口中完成后点击「继续」。")
            elif state == STATE_SOLVE_CAPTCHA:
                self._append_log("[*] 当前状态: 正在自动识别图形验证码 ...")
            elif state == STATE_WAIT_EMAIL_VERIFY:
                self._append_log("[!] 检测到邮箱安全验证：请在浏览器（应用窗口）完成验证，"
                                 "完成后会自动继续，无需点「继续」。")
            elif state == FINISHED:
                self._append_log("[+] 当前状态: 已完成。可点击「重新执行」。")
            elif state == ERROR:
                self._append_log("[!] 当前状态: 出错。请点击「重新执行」重试。")

        # 按钮与状态文本
        if state == IDLE:
            # 只读后台探测线程维护的缓存，绝不在 Tk 主线程里发网络请求
            app_running = bool(_APP_RUNNING_CACHE.get("value"))
            self.start_btn.config(
                state=tk.NORMAL if (is_exe_path_set() or app_running) else tk.DISABLED)
            self.pause_btn.config(state=tk.DISABLED)
            self.continue_btn.config(state=tk.DISABLED)
            self.retry_btn.config(state=tk.DISABLED)
            if app_running:
                self.status_var.set("当前状态: 就绪（检测到应用已打开，启动时将直接续跑）")
            elif not is_exe_path_set():
                self.status_var.set("当前状态: 就绪（未配置路径则无法启动）")
            else:
                self.status_var.set("当前状态: 就绪")
        elif state == RUNNING:
            self.start_btn.config(state=tk.DISABLED)
            self.pause_btn.config(state=tk.NORMAL)
            self.continue_btn.config(state=tk.DISABLED)
            self.retry_btn.config(state=tk.DISABLED)
            self.status_var.set("当前状态: 运行中")
        elif state == PAUSED:
            self.start_btn.config(state=tk.DISABLED)
            self.pause_btn.config(state=tk.DISABLED)
            self.continue_btn.config(state=tk.NORMAL)
            self.retry_btn.config(state=tk.DISABLED)
            self.status_var.set("当前状态: 已暂停")
        elif state == STATE_WAIT_CAPTCHA:
            self.start_btn.config(state=tk.DISABLED)
            self.pause_btn.config(state=tk.NORMAL)
            self.continue_btn.config(state=tk.DISABLED)
            self.retry_btn.config(state=tk.DISABLED)
            self.status_var.set("当前状态: 等待你输入验证码")
        elif state == STATE_SOLVE_CAPTCHA:
            self.start_btn.config(state=tk.DISABLED)
            self.pause_btn.config(state=tk.NORMAL)
            self.continue_btn.config(state=tk.DISABLED)
            self.retry_btn.config(state=tk.DISABLED)
            self.status_var.set("当前状态: 正在自动识别验证码")
        elif state == STATE_WAIT_EMAIL_VERIFY:
            # 等待期间：自动化已暂停点击/输入，用户可在应用窗口手工验证；
            # 状态栏同步显示已等待时长（EMAIL_VERIFY_WAIT 由后台线程维护）
            self.start_btn.config(state=tk.DISABLED)
            self.pause_btn.config(state=tk.NORMAL)
            self.continue_btn.config(state=tk.DISABLED)
            self.retry_btn.config(state=tk.DISABLED)
            waited = int(time.time() - EMAIL_VERIFY_WAIT.get("t0", time.time())) \
                if EMAIL_VERIFY_WAIT.get("active") else 0
            self.status_var.set("当前状态: 等待邮箱验证（请在浏览器完成验证，已等待 %d 秒 / 上限 %d 秒）"
                                % (waited, EMAIL_VERIFY_TIMEOUT))
        elif state in (FINISHED, ERROR):
            self.start_btn.config(state=tk.DISABLED)
            self.pause_btn.config(state=tk.DISABLED)
            self.continue_btn.config(state=tk.DISABLED)
            self.retry_btn.config(state=tk.NORMAL)
            self.status_var.set("当前状态: 已完成" if state == FINISHED else "当前状态: 出错")

        # [已注释] 「重载代码」始终可用（不限制状态）——功能已停用
        # self.reload_btn.config(state=tk.NORMAL)

        # 关键：250ms 后再次刷新，否则界面只更新一次就“卡死”
        self.root.after(250, self._update_ui)

    def on_start(self):
        # 用锁 + _FLOW_ACTIVE 双保险防竞态：
        # 避免极快双击「启动」导致两个 run_login_flow 同时跑起来。
        if not _START_LOCK.acquire(blocking=False):
            self._append_log("[!] 已有启动流程在执行，请稍候再点。")
            return
        try:
            global _FLOW_ACTIVE
            if _FLOW_ACTIVE:
                self._append_log("[!] 已有启动流程在执行，请稍候再点。")
                return
            if current_state != IDLE:
                self._append_log("[!] 当前状态不允许启动：" + str(current_state))
                return

            # 先检测本应用是否已经打开：已打开就直接复用（不再启动、不再登录）。
            # 用 _is_our_app_on_port 而非 is_app_running，避免端口被其它程序占用时误判。
            if _is_our_app_on_port(CONFIG["DEBUG_PORT"]):
                self._append_log("[*] 检测到上传系统应用已打开，将直接复用该实例"
                                 "（跳过打开应用；若已登录则直接执行登录后的逻辑）。")
            else:
                if not is_exe_path_set():
                    self._append_log("[!] 尚未设置软件路径，请先点「设置」选择上传系统 exe。")
                    return
                if not os.path.isfile(CONFIG["EXE_PATH"]):
                    self._append_log("[!] 软件路径不存在: " + CONFIG["EXE_PATH"]
                                     + "，请重新点「设置」选择。")
                    return
                self._append_log("[*] 点击「启动」后开始执行自动化流程。")
        finally:
            _START_LOCK.release()

        thread = threading.Thread(target=run_login_flow, daemon=True)
        thread.start()

    def on_pause(self):
        if current_state not in (RUNNING, STATE_WAIT_CAPTCHA, STATE_SOLVE_CAPTCHA, STATE_WAIT_EMAIL_VERIFY):
            self._append_log("[!] 当前状态不允许暂停：" + str(current_state))
            return

        set_state(PAUSED)
        self._append_log("[!] 已暂停自动化流程。")

    def on_continue(self):
        if current_state != PAUSED:
            self._append_log("[!] 当前状态不允许继续：" + str(current_state))
            return

        set_state(RUNNING)
        self._append_log("[+] 已继续。")

    def on_retry(self):
        if current_state not in (FINISHED, ERROR):
            self._append_log("[!] 当前状态不允许重新执行：" + str(current_state))
            return
        if _FLOW_ACTIVE:
            self._append_log("[!] 上一轮流程仍在收尾，请稍候再点「重新执行」。")
            return

        _reset_run_context()          # 清掉上一轮 JSONL / 上下文 / 页面 target
        set_state(IDLE)
        self._append_log("[*] 已重置为就绪状态（上一轮中转数据已清理）。点击「启动」重新执行。")

    # [已注释] 「重载代码」功能（热重载）已停用，取消注释可恢复
    # def on_reload(self):
    #     """重载代码：加载本文件最新代码，只执行登录之后的后续逻辑（按钮始终可用）。"""
    #     self._append_log("[*] 正在重载代码（不会重新打开应用、不会重新登录） ...")
    #     try:
    #         hot_reload_module()
    #     except Exception as e:
    #         self._append_log("[!] 重载代码失败: " + repr(e))
    #         return
    #
    #     if current_state in (RUNNING, PAUSED, STATE_WAIT_CAPTCHA, STATE_SOLVE_CAPTCHA):
    #         # 已有自动化流程在跑，本次只重载代码，避免两个流程并发互相干扰
    #         self._append_log("[!] 当前有自动化流程正在运行，本次仅重载了代码；"
    #                          "流程结束后再点「重载代码」即可执行登录后的新逻辑。")
    #         return
    #
    #     self._append_log("[+] 代码重载完成，开始执行登录后的后续逻辑。")
    #     threading.Thread(target=run_post_login_flow, daemon=True).start()

    def refresh_exe_path_label(self):
        if is_exe_path_set():
            self.exe_path_var.set("软件路径: " + CONFIG["EXE_PATH"])
            self.exe_path_label.config(fg="#225522")
        else:
            self.exe_path_var.set("软件路径: （未设置或文件不存在，请点「设置」）")
            self.exe_path_label.config(fg="#AA5500")

    def on_settings(self):
        SettingsDialog(self)

    def apply_settings(self, exe_path, username, password, debug_port, keep_app_open=True,
                       auto_download_template=True, auto_open_template=True,
                       auto_fill_template=True, close_xlsx_after_open=True,
                       auto_import_template=True):
        CONFIG["EXE_PATH"] = exe_path
        CONFIG["USERNAME"] = username
        CONFIG["PASSWORD"] = password
        CONFIG["DEBUG_PORT"] = debug_port
        CONFIG["KEEP_APP_OPEN"] = bool(keep_app_open)
        CONFIG["AUTO_DOWNLOAD_TEMPLATE"] = bool(auto_download_template)
        CONFIG["AUTO_OPEN_TEMPLATE"] = bool(auto_open_template)
        CONFIG["AUTO_FILL_TEMPLATE"] = bool(auto_fill_template)
        CONFIG["CLOSE_XLSX_AFTER_OPEN"] = bool(close_xlsx_after_open)
        CONFIG["AUTO_IMPORT_TEMPLATE"] = bool(auto_import_template)
        save_config(CONFIG)
        self.refresh_exe_path_label()
        self._append_log("[+] 已更新设置：软件路径=" + CONFIG["EXE_PATH"] + "，调试端口=" + str(CONFIG["DEBUG_PORT"])
                          + "，登录后保留应用=" + ("是" if CONFIG["KEEP_APP_OPEN"] else "否")
                          + "，自动下载模板=" + ("是" if CONFIG["AUTO_DOWNLOAD_TEMPLATE"] else "否")
                          + "，自动打开文件=" + ("是" if CONFIG["AUTO_OPEN_TEMPLATE"] else "否")
                          + "，自动填写数据=" + ("是" if CONFIG["AUTO_FILL_TEMPLATE"] else "否")
                          + "，自动关闭xlsx=" + ("是" if CONFIG["CLOSE_XLSX_AFTER_OPEN"] else "否")
                          + "，自动导入数据表格=" + ("是" if CONFIG["AUTO_IMPORT_TEMPLATE"] else "否"))

    def run(self):
        # 首次进入若未配置软件路径，自动弹出设置
        if not is_exe_path_set():
            self._append_log("[!] 尚未配置软件路径，已自动打开「设置」，请先选择 exe。")
            self.root.after(300, self.on_settings)
        # 定时刷新：日志批量落屏 + UI 状态 + 应用运行探测
        self.root.after(_GUI_LOG_POLL_MS, self._poll_log_queue)
        self.root.after(250, self._update_ui)
        _start_app_probe_worker()
        try:
            self.root.mainloop()
        finally:
            # 窗口关闭后后台 daemon 线程仍可能 print：让日志流静默并恢复原始终端
            _GUI_LOG_STATE["alive"] = False
            if sys.__stdout__ is not None:
                sys.stdout = sys.__stdout__
            if sys.__stderr__ is not None:
                sys.stderr = sys.__stderr__


def main():
    # 环境检查
    if CONFIG["USERNAME"].startswith("在这里") or CONFIG["PASSWORD"].startswith("在这里"):
        print("[!] 请先在脚本配置区填写 USERNAME 和 PASSWORD")
        sys.exit(1)

    root = tk.Tk()
    app = AutoLoginApp(root)
    # 把 print()/traceback 重定向到 GUI 日志框，
    # 这样打包成无控制台（--windowed）的 exe 也能看到运行信息。
    sys.stdout = _GuiStream(app.log_text)
    sys.stderr = _GuiStream(app.log_text)
    app.run()


if __name__ == "__main__":
    main()


