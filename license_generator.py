import argparse
import base64
import json
import re
import time
from typing import Any, Dict

from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
from cryptography.hazmat.primitives.serialization import Encoding, PrivateFormat, PublicFormat, NoEncryption


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


def _build_payload(machine_code: str, product_code: str, license_type: str, expire_at: int, schema_version: int) -> Dict[str, Any]:
    return {
        "schema_version": schema_version,
        "machine_code": machine_code,
        "product_code": product_code,
        "license_type": license_type,
        "expire_at": int(expire_at),
    }


def cmd_gen_keys(_: argparse.Namespace) -> int:
    private_key = Ed25519PrivateKey.generate()
    private_bytes = private_key.private_bytes(
        encoding=Encoding.Raw,
        format=PrivateFormat.Raw,
        encryption_algorithm=NoEncryption(),
    )
    public_bytes = private_key.public_key().public_bytes(
        encoding=Encoding.Raw,
        format=PublicFormat.Raw,
    )

    print("private_key_b64url=" + _b64url_encode(private_bytes))
    print("public_key_b64url=" + _b64url_encode(public_bytes))
    return 0


def cmd_gen_license(args: argparse.Namespace) -> int:
    private_key_bytes = _b64url_decode(args.private_key_b64url)
    if len(private_key_bytes) != 32:
        raise SystemExit("private_key_b64url 必须为 32 字节 Ed25519 私钥")

    private_key = Ed25519PrivateKey.from_private_bytes(private_key_bytes)

    now_ts = int(time.time())
    if getattr(args, "expire_days", None) is not None:
        expire_at = now_ts + int(args.expire_days) * 86400
    else:
        expire_at = int(args.expire_at)
        if (not getattr(args, "allow_expired", False)) and (expire_at <= now_ts):
            raise SystemExit(
                "expire_at 必须是未来的 Unix 时间戳(秒)。\n"
                "你当前传入的 expire_at 看起来不是时间戳，可能误把“天数”当成了时间戳。\n"
                "建议用: python license_generator.py expire-days <天数> 先算出时间戳，\n"
                "或直接用: python license_generator.py gen-license --expire-days <天数>"
            )
    machine_code = re.sub(r"[\s-]+", "", args.machine_code).lower()
    if not re.fullmatch(r"[0-9a-f]{64}", machine_code):
        raise SystemExit("machine_code 格式应为 64 位十六进制字符串")

    payload = _build_payload(
        machine_code=machine_code,
        product_code=args.product_code.strip(),
        license_type=args.license_type.strip(),
        expire_at=expire_at,
        schema_version=int(args.schema_version),
    )

    payload_bytes = _canonical_json_bytes(payload)
    sig = private_key.sign(payload_bytes)

    license_key = _b64url_encode(payload_bytes) + "." + _b64url_encode(sig)
    print(license_key)
    return 0


def cmd_expire_days(args: argparse.Namespace) -> int:
    now_ts = int(time.time())
    expire_at = now_ts + int(args.days) * 86400
    print(expire_at)
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(prog="license_generator")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_keys = sub.add_parser("gen-keys")
    p_keys.set_defaults(func=cmd_gen_keys)

    p_exp = sub.add_parser("expire-days")
    p_exp.add_argument("days", type=int)
    p_exp.set_defaults(func=cmd_expire_days)

    p_gen = sub.add_parser("gen-license")
    p_gen.add_argument("--private-key-b64url", required=True)
    p_gen.add_argument("--machine-code", required=True)
    p_gen.add_argument("--license-type", required=True, choices=["trial", "edu", "personal", "pro"])
    exp_group = p_gen.add_mutually_exclusive_group(required=True)
    exp_group.add_argument("--expire-at", help="Unix 时间戳（秒）")
    exp_group.add_argument("--expire-days", type=int, help="从现在起的有效天数（例如 365）")
    p_gen.add_argument("--allow-expired", action="store_true", help="允许生成已过期授权（仅用于测试）")
    p_gen.add_argument("--product-code", default="shixun_report_tool")
    p_gen.add_argument("--schema-version", default=1, type=int)
    p_gen.set_defaults(func=cmd_gen_license)

    args = parser.parse_args()
    return int(args.func(args))


if __name__ == "__main__":
    raise SystemExit(main())
