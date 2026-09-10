# -*- coding: utf-8 -*-
"""默认打印机隔离(Windows): 启动办公引擎期间临时改用本地虚拟打印机。

为什么需要: WPS / LibreOffice / Chrome 启动时都会初始化 Windows 打印子系统并读取
"默认打印机"能力(实测 `soffice.bin` 加载 WINSPOOL.DRV + PrintConfig.dll)。若默认打印机
是 WSD / IP 网络打印机,系统就会去连接它 —— 表现为启动变慢、系统托盘出现"连接打印机",
而 office-cli 的任何命令都不需要打印。

做法: 引擎调用前后把当前用户的默认打印机临时切到本地虚拟打印机(缺省
"Microsoft Print to PDF"),用完立即恢复原值;默认打印机本身已是本地端口时不做任何
操作(零副作用),非 Windows 平台 no-op(macOS/Linux 不走 winspool)。

开关:
- `OFFICE_PRINTER_ISOLATE=0`  关闭隔离(默认开启)
- `OFFICE_VIRTUAL_PRINTER=名称` 指定替代用的本地虚拟打印机

实现只用标准库(winreg / subprocess),装了 pywin32 时走更快更静的
`win32print.SetDefaultPrinter`,否则退回 `rundll32 printui.dll,PrintUIEntry /y`。
"""

from __future__ import annotations

import os
import subprocess
import sys

ENV_ENABLE = "OFFICE_PRINTER_ISOLATE"
ENV_VIRTUAL = "OFFICE_VIRTUAL_PRINTER"

_WIN_KEY = r"Software\Microsoft\Windows NT\CurrentVersion\Windows"
_PRINTERS_KEY = r"SYSTEM\CurrentControlSet\Control\Print\Printers"

# 候选虚拟打印机(顺序即优先级;端口必须是本地的)
_VIRTUAL_CANDIDATES = (
    "Microsoft Print to PDF",
    "Microsoft XPS Document Writer",
    "输出为 WPS PDF",
    "WPS PDF",
    "Adobe PDF",
)

# 网络端口前缀(WSD 发现 / TCP-IP 直连 / UNC 共享 / LPD / IPP)
_NET_PORT_PREFIXES = ("wsd-", "ip_", "\\\\", "lpr", "tcp", "http", "ipp", "snmp")


def enabled() -> bool:
    """隔离是否启用(默认启用;OFFICE_PRINTER_ISOLATE=0 关闭)。"""
    return (os.environ.get(ENV_ENABLE) or "1").strip().lower() not in (
        "0", "false", "no", "off")


def _hkey(name: str):
    """字符串 → winreg 根键常量。"""
    import winreg

    return {"HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
            "HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE}[name]


def _open_key(root: str, path: str):
    import winreg

    return winreg.OpenKey(_hkey(root), path)


def _read_value(root: str, path: str, name: str) -> str | None:
    try:
        import winreg

        with _open_key(root, path) as k:
            val, _ = winreg.QueryValueEx(k, name)
        return str(val)
    except Exception:  # noqa: BLE001 注册表不可读/键不存在 → 视为未知
        return None


def current_default() -> str | None:
    """当前用户默认打印机名(HKCU\\...\\Windows\\Device 的第一段);非 Windows 返回 None。"""
    if os.name != "nt":
        return None
    raw = _read_value("HKEY_CURRENT_USER", _WIN_KEY, "Device")
    if not raw:
        return None
    return raw.split(",")[0].strip() or None


def list_printers() -> list[str]:
    """本机已安装打印机名列表(HKLM ...\\Print\\Printers 的子键名)。"""
    if os.name != "nt":
        return []
    names: list[str] = []
    try:
        import winreg

        with _open_key("HKEY_LOCAL_MACHINE", _PRINTERS_KEY) as k:
            i = 0
            while True:
                try:
                    names.append(winreg.EnumKey(k, i))
                except OSError:
                    break
                i += 1
    except Exception:  # noqa: BLE001 无权限/键不存在
        return names
    return names


def port_of(name: str | None) -> str | None:
    """打印机使用的端口名(如 PORTPROMPT:、WSD-xxx、IP_192.168.1.5)。"""
    if not name or os.name != "nt":
        return None
    return _read_value("HKEY_LOCAL_MACHINE", _PRINTERS_KEY + "\\" + name, "Port")


def is_network_port(port: str | None) -> bool:
    """端口是否指向网络打印机(true 时才有隔离价值)。"""
    p = (port or "").strip().lower()
    return bool(p) and p.startswith(_NET_PORT_PREFIXES)


def isolate_target() -> str | None:
    """替代用的本地虚拟打印机名(环境变量优先;不存在则返回 None)。"""
    want = (os.environ.get(ENV_VIRTUAL) or "").strip()
    installed = {n.lower(): n for n in list_printers()}
    if want:
        return installed.get(want.lower())
    for cand in _VIRTUAL_CANDIDATES:
        if cand.lower() in installed:
            return installed[cand.lower()]
    return None


def needs_isolation() -> bool:
    """默认打印机是否是网络打印机(需要隔离)。"""
    if os.name != "nt" or not enabled():
        return False
    cur = current_default()
    return bool(cur) and is_network_port(port_of(cur))


def set_default(name: str) -> bool:
    """设置当前用户默认打印机(成功返回 True)。"""
    if not name or os.name != "nt":
        return False
    try:  # pywin32 可用时最快最静默
        import win32print

        win32print.SetDefaultPrinter(name)
        if (current_default() or "").lower() == name.lower():
            return True
    except Exception:  # noqa: BLE001 未装 pywin32 / 调用失败 → 退回 rundll32
        pass
    try:
        subprocess.run(
            ["rundll32", "printui.dll,PrintUIEntry", "/y", "/n", name],
            capture_output=True, timeout=30, check=False,
        )
    except Exception:  # noqa: BLE001 命令不可用
        return False
    return (current_default() or "").lower() == name.lower()


class isolate:
    """上下文管理器: 期间把默认打印机切到本地虚拟打印机,退出时恢复。

    以下情况不做任何操作(零副作用): 非 Windows / 隔离被关闭 / 默认打印机已是本地端口 /
    找不到可用的本地虚拟打印机。恢复在 finally 中执行,且只恢复"确实被我们改过"的情况。
    """

    def __init__(self, force: bool = False) -> None:
        self.force = force          # force=True 时不判断端口类型,一律切换
        self._original: str | None = None
        self._switched_to: str | None = None

    def __enter__(self) -> "isolate":
        if os.name != "nt" or not enabled():
            return self
        try:
            cur = current_default()
            if not cur:
                return self
            if not self.force and not is_network_port(port_of(cur)):
                return self                    # 已经是本地端口,不必动它
            target = isolate_target()
            if not target or target.lower() == cur.lower():
                return self
            if set_default(target):
                self._original, self._switched_to = cur, target
        except Exception:  # noqa: BLE001 隔离失败不影响主流程(引擎照跑)
            if os.environ.get("OFFICE_DEBUG"):
                raise
        return self

    def __exit__(self, exc_type, exc, tb) -> bool:
        if self._original:
            try:
                if not set_default(self._original):
                    # 恢复失败再试一次(rundll32 偶发被占用)
                    set_default(self._original)
            except Exception:  # noqa: BLE001 恢复失败不掩盖业务异常
                pass
            self._original = None
        return False

    @property
    def switched(self) -> bool:
        """本次是否真的切换过(供日志/诊断)。"""
        return self._switched_to is not None


def info() -> dict:
    """诊断信息(供 `office info`)。"""
    cur = current_default()
    port = port_of(cur)
    return {
        "isolate_enabled": enabled() and sys.platform == "win32",
        "default": cur,
        "default_port": port,
        "default_is_network": is_network_port(port),
        "virtual_target": isolate_target(),
    }
