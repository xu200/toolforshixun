import base64
import ctypes
import hashlib
import json
import os
import re
import subprocess
import time
from ctypes import wintypes
from dataclasses import dataclass
from typing import Any, Dict, Optional, Tuple

from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey


PRODUCT_CODE = "shixun_report_tool"
SCHEMA_VERSION = 1
LICENSE_KEY_ENV_PUBLIC_KEY = "SHIXUN_REPORT_TOOL_PUBLIC_KEY"

PUBLIC_KEY_B64URL = "GcsaO5OZtoyo0FkJYc9z26_BO5U2izfgcPDQJveBHJ8"

_MACHINE_CODE_CACHE: Optional[str] = None


def _b64url_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode("ascii").rstrip("=")


def _b64url_decode(text: str) -> bytes:
    text = text.strip()
    padding = "=" * ((4 - len(text) % 4) % 4)
    return base64.urlsafe_b64decode((text + padding).encode("ascii"))


def _canonical_json_bytes(payload: Dict[str, Any]) -> bytes:
    return json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")


class _DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_byte)),
    ]


def _dpapi_bind():
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32

    crypt32.CryptProtectData.argtypes = [
        ctypes.POINTER(_DATA_BLOB),
        wintypes.LPCWSTR,
        ctypes.POINTER(_DATA_BLOB),
        wintypes.LPVOID,
        wintypes.LPVOID,
        wintypes.DWORD,
        ctypes.POINTER(_DATA_BLOB),
    ]
    crypt32.CryptProtectData.restype = wintypes.BOOL

    crypt32.CryptUnprotectData.argtypes = [
        ctypes.POINTER(_DATA_BLOB),
        ctypes.POINTER(wintypes.LPWSTR),
        ctypes.POINTER(_DATA_BLOB),
        wintypes.LPVOID,
        wintypes.LPVOID,
        wintypes.DWORD,
        ctypes.POINTER(_DATA_BLOB),
    ]
    crypt32.CryptUnprotectData.restype = wintypes.BOOL

    kernel32.LocalFree.argtypes = [wintypes.HLOCAL]
    kernel32.LocalFree.restype = wintypes.HLOCAL

    return crypt32, kernel32


def _dpapi_encrypt(plain: bytes) -> Optional[bytes]:
    if os.name != "nt":
        return None
    try:
        if not plain:
            return b""
        buffer = ctypes.create_string_buffer(plain)
        in_blob = _DATA_BLOB(len(plain), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
        out_blob = _DATA_BLOB()
        crypt32, kernel32 = _dpapi_bind()
        ok = crypt32.CryptProtectData(
            ctypes.byref(in_blob),
            None,
            None,
            None,
            None,
            0,
            ctypes.byref(out_blob),
        )
        if not ok:
            return None
        try:
            return ctypes.string_at(out_blob.pbData, out_blob.cbData)
        finally:
            kernel32.LocalFree(ctypes.cast(out_blob.pbData, ctypes.c_void_p))
    except Exception:
        return None


def _dpapi_decrypt(protected: bytes) -> Optional[bytes]:
    if os.name != "nt":
        return None
    try:
        if protected == b"":
            return b""
        buffer = ctypes.create_string_buffer(protected)
        in_blob = _DATA_BLOB(len(protected), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
        out_blob = _DATA_BLOB()
        crypt32, kernel32 = _dpapi_bind()
        ok = crypt32.CryptUnprotectData(
            ctypes.byref(in_blob),
            None,
            None,
            None,
            None,
            0,
            ctypes.byref(out_blob),
        )
        if not ok:
            return None
        try:
            return ctypes.string_at(out_blob.pbData, out_blob.cbData)
        finally:
            kernel32.LocalFree(ctypes.cast(out_blob.pbData, ctypes.c_void_p))
    except Exception:
        return None


def _run_powershell(command: str, timeout: float = 3.0) -> str:
    out = subprocess.check_output(
        ["powershell", "-NoProfile", "-Command", command],
        stderr=subprocess.STDOUT,
        timeout=timeout,
        text=True,
        encoding="utf-8",
        errors="ignore",
    )
    return out


def _wmic_value(wmic_args: list[str], key: str) -> Optional[str]:
    try:
        out = subprocess.check_output(
            ["wmic", *wmic_args],
            stderr=subprocess.STDOUT,
            timeout=3.0,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
    except Exception:
        return None

    for line in out.splitlines():
        line = line.strip()
        if not line:
            continue
        if line.lower().startswith(key.lower() + "="):
            return line.split("=", 1)[1].strip() or None
    return None


def _first_non_empty_line(text: str) -> Optional[str]:
    for line in text.splitlines():
        v = line.strip()
        if v:
            return v
    return None


def _normalize_hw_value(value: Optional[str]) -> str:
    if value is None:
        return "UNKNOWN"
    v = value.strip()
    if not v:
        return "UNKNOWN"
    if v.strip().lower() == "to be filled by o.e.m.":
        return "UNKNOWN"
    return v


def _get_cpu_id() -> str:
    v = _wmic_value(["cpu", "get", "ProcessorId", "/value"], "ProcessorId")
    nv = _normalize_hw_value(v)
    if nv != "UNKNOWN":
        return nv
    try:
        out = _run_powershell("(Get-CimInstance Win32_Processor | Select-Object -First 1 -ExpandProperty ProcessorId)")
        vv = _first_non_empty_line(out)
        return _normalize_hw_value(vv)
    except Exception:
        return "UNKNOWN"


def _get_board_serial() -> str:
    v = _wmic_value(["baseboard", "get", "SerialNumber", "/value"], "SerialNumber")
    nv = _normalize_hw_value(v)
    if nv != "UNKNOWN":
        return nv
    try:
        out = _run_powershell("(Get-CimInstance Win32_BaseBoard | Select-Object -First 1 -ExpandProperty SerialNumber)")
        vv = _first_non_empty_line(out)
        return _normalize_hw_value(vv)
    except Exception:
        return "UNKNOWN"


def _get_system_uuid() -> str:
    v = _wmic_value(["csproduct", "get", "UUID", "/value"], "UUID")
    nv = _normalize_hw_value(v)
    if nv != "UNKNOWN":
        return nv
    try:
        out = _run_powershell("(Get-CimInstance Win32_ComputerSystemProduct | Select-Object -First 1 -ExpandProperty UUID)")
        vv = _first_non_empty_line(out)
        return _normalize_hw_value(vv)
    except Exception:
        return "UNKNOWN"


def get_machine_code() -> str:
    global _MACHINE_CODE_CACHE
    if _MACHINE_CODE_CACHE:
        return _MACHINE_CODE_CACHE
    cpu_id = _get_cpu_id()
    board = _get_board_serial()
    uuid = _get_system_uuid()

    raw_fingerprint = f"{cpu_id}|{board}|{uuid}"
    _MACHINE_CODE_CACHE = hashlib.sha256(raw_fingerprint.encode("utf-8")).hexdigest()
    return _MACHINE_CODE_CACHE


def get_machine_code_display() -> str:
    code = get_machine_code().upper()
    parts = [code[i : i + 8] for i in range(0, 64, 8)]
    return "-".join(parts)


@dataclass
class LicenseCheckResult:
    ok: bool
    message: str
    payload: Optional[Dict[str, Any]] = None


def encode_license_key(payload: Dict[str, Any], signature: bytes) -> str:
    payload_b64 = _b64url_encode(_canonical_json_bytes(payload))
    sig_b64 = _b64url_encode(signature)
    return f"{payload_b64}.{sig_b64}"


def decode_license_key(license_key: str) -> Tuple[Dict[str, Any], bytes]:
    parts = [p for p in license_key.strip().split(".") if p]
    if len(parts) != 2:
        raise ValueError("invalid_license_key_format")

    payload_bytes = _b64url_decode(parts[0])
    sig = _b64url_decode(parts[1])

    payload = json.loads(payload_bytes.decode("utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("invalid_payload")
    return payload, sig


def _load_public_key_bytes() -> bytes:
    key_text = (os.getenv(LICENSE_KEY_ENV_PUBLIC_KEY) or "").strip()
    if not key_text:
        key_text = (PUBLIC_KEY_B64URL or "").strip()
    if not key_text:
        raise RuntimeError("public_key_not_configured")

    key_text = re.sub(r"\s+", "", key_text)
    return _b64url_decode(key_text)


def verify_license_key(license_key: str, now_ts: Optional[int] = None) -> LicenseCheckResult:
    if now_ts is None:
        now_ts = int(time.time())

    try:
        payload, sig = decode_license_key(license_key)
    except Exception:
        return LicenseCheckResult(False, "激活码格式不正确")

    try:
        public_key_bytes = _load_public_key_bytes()
        pub = Ed25519PublicKey.from_public_bytes(public_key_bytes)
        payload_bytes = _canonical_json_bytes(payload)
        pub.verify(sig, payload_bytes)
    except RuntimeError as e:
        if str(e) == "public_key_not_configured":
            return LicenseCheckResult(False, "客户端未配置公钥，无法校验")
        return LicenseCheckResult(False, "验签失败")
    except Exception:
        return LicenseCheckResult(False, "验签失败")

    if payload.get("schema_version") not in (None, SCHEMA_VERSION):
        return LicenseCheckResult(False, "激活码版本不兼容")

    if payload.get("product_code") != PRODUCT_CODE:
        return LicenseCheckResult(False, "产品不匹配")

    local_machine = get_machine_code()
    payload_machine = str(payload.get("machine_code") or "").strip()
    payload_machine = re.sub(r"[\s-]+", "", payload_machine).lower()
    if payload_machine != local_machine:
        return LicenseCheckResult(False, "机器码不匹配")

    try:
        expire_at = int(payload.get("expire_at"))
    except Exception:
        return LicenseCheckResult(False, "到期时间无效")

    if now_ts > expire_at:
        return LicenseCheckResult(False, "授权已过期")

    license_type = payload.get("license_type")
    if license_type not in {"trial", "edu", "personal", "pro"}:
        return LicenseCheckResult(False, "授权类型无效")

    return LicenseCheckResult(True, "激活成功", payload)


class LicenseStore:
    def __init__(self):
        base = os.getenv("APPDATA")
        if not base:
            base = os.path.expanduser("~\\AppData\\Roaming")
        self.dir_path = os.path.join(base, PRODUCT_CODE)
        self.file_path = os.path.join(self.dir_path, "license.json")

    def clear(self) -> bool:
        try:
            if os.path.exists(self.file_path):
                os.remove(self.file_path)
            return not os.path.exists(self.file_path)
        except Exception:
            return False

    def load(self) -> Optional[Dict[str, Any]]:
        try:
            if not os.path.exists(self.file_path):
                return None
            with open(self.file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                return None

            if (not (data.get("license_key") or "").strip()) and (data.get("license_key_protected") is not None):
                enc = str(data.get("license_key_protected") or "").strip()
                if enc:
                    protected = _b64url_decode(enc)
                    plain = _dpapi_decrypt(protected)
                    if plain is None:
                        return None
                    data["license_key"] = plain.decode("utf-8", errors="ignore")
            return data
        except Exception:
            return None

    def save(self, data: Dict[str, Any]) -> None:
        os.makedirs(self.dir_path, exist_ok=True)
        tmp = self.file_path + ".tmp"

        record = dict(data)
        record.pop("payload", None)
        license_key = (record.get("license_key") or "").strip()
        if license_key:
            protected = _dpapi_encrypt(license_key.encode("utf-8"))
            if protected is not None:
                record["license_key_protected"] = _b64url_encode(protected)
                record["protection"] = "dpapi"
                record.pop("license_key", None)

        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(record, f, ensure_ascii=False, indent=2)
        os.replace(tmp, self.file_path)


class LicenseManager:
    def __init__(self):
        self.store = LicenseStore()

    def deactivate(self) -> bool:
        try:
            return bool(self.store.clear())
        except Exception:
            return False

    def check(self) -> LicenseCheckResult:
        record = self.store.load()
        if not record:
            return LicenseCheckResult(False, "未激活")

        license_key = (record.get("license_key") or "").strip()
        if not license_key:
            return LicenseCheckResult(False, "未激活")

        now_ts = int(time.time())
        try:
            last_verified_at = int(record.get("last_verified_at")) if record.get("last_verified_at") is not None else None
        except Exception:
            last_verified_at = None

        if last_verified_at is not None:
            if now_ts + 600 < last_verified_at:
                return LicenseCheckResult(False, "系统时间异常")

        result = verify_license_key(license_key, now_ts=now_ts)
        if result.ok:
            record["last_verified_at"] = now_ts
            try:
                self.store.save(record)
            except Exception:
                pass
        return result

    def activate(self, license_key: str) -> LicenseCheckResult:
        now_ts = int(time.time())
        result = verify_license_key(license_key, now_ts=now_ts)
        if not result.ok:
            return result

        record = {
            "license_key": license_key.strip(),
            "payload": result.payload or {},
            "activated_at": now_ts,
            "last_verified_at": now_ts,
        }
        try:
            self.store.save(record)
        except Exception:
            return LicenseCheckResult(False, "授权保存失败")

        return result
