#!/usr/bin/env python3
"""Upload a generated report directory to Aliyun OSS and print its public URL."""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
import mimetypes
import sys
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from email.utils import formatdate
from pathlib import Path


DEFAULT_CONFIG = Path(__file__).resolve().parents[1] / "config" / "oss_config.json"


@dataclass
class OssConfig:
    access_key_id: str
    access_key_secret: str
    endpoint: str
    bucket: str
    public_base_url: str
    remote_prefix: str
    public_read: bool
    signed_url_expires_days: int


def fail(message: str) -> None:
    print(f"error: {message}", file=sys.stderr)
    raise SystemExit(1)


def load_config(path: Path) -> OssConfig:
    if not path.exists():
        fail(f"missing OSS config: {path}")
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        fail(f"invalid OSS config JSON: {exc}")

    required = ["access_key_id", "access_key_secret", "endpoint", "bucket"]
    missing = [key for key in required if not raw.get(key)]
    if missing:
        fail(f"missing OSS config keys: {', '.join(missing)}")

    return OssConfig(
        access_key_id=raw["access_key_id"],
        access_key_secret=raw["access_key_secret"],
        endpoint=raw["endpoint"].rstrip("/"),
        bucket=raw["bucket"],
        public_base_url=raw.get("public_base_url", "").rstrip("/"),
        remote_prefix=raw.get("remote_prefix", ""),
        public_read=bool(raw.get("public_read", True)),
        signed_url_expires_days=int(raw.get("signed_url_expires_days", 3)),
    )


def endpoint_host(endpoint: str) -> str:
    parsed = urllib.parse.urlparse(endpoint if "://" in endpoint else f"https://{endpoint}")
    return parsed.netloc or parsed.path


def object_key(prefix: str, relative_path: Path) -> str:
    normalized_prefix = prefix.strip("/")
    relative = "/".join(relative_path.parts)
    return f"{normalized_prefix}/{relative}" if normalized_prefix else relative


def quote_key(key: str) -> str:
    return "/".join(urllib.parse.quote(part) for part in key.split("/"))


def content_type(path: Path) -> str:
    guessed, _encoding = mimetypes.guess_type(path.name)
    if guessed:
        if guessed.startswith("text/") or guessed in {"application/javascript", "application/json"}:
            return f"{guessed}; charset=utf-8"
        return guessed
    return "application/octet-stream"


def sign(config: OssConfig, method: str, key: str, headers: dict[str, str], content_md5: str, ctype: str) -> str:
    oss_headers = {
        name.lower(): value.strip()
        for name, value in headers.items()
        if name.lower().startswith("x-oss-")
    }
    canonical_oss_headers = "".join(
        f"{name}:{oss_headers[name]}\n" for name in sorted(oss_headers)
    )
    canonical_resource = f"/{config.bucket}/{key}"
    string_to_sign = (
        f"{method}\n{content_md5}\n{ctype}\n{headers['Date']}\n"
        f"{canonical_oss_headers}{canonical_resource}"
    )
    digest = hmac.new(
        config.access_key_secret.encode("utf-8"),
        string_to_sign.encode("utf-8"),
        hashlib.sha1,
    ).digest()
    return base64.b64encode(digest).decode("ascii")


def put_object(config: OssConfig, local_path: Path, key: str) -> None:
    body = local_path.read_bytes()
    ctype = content_type(local_path)
    content_md5 = base64.b64encode(hashlib.md5(body).digest()).decode("ascii")
    headers = {
        "Date": formatdate(usegmt=True),
        "Content-Type": ctype,
        "Content-MD5": content_md5,
        "Content-Disposition": "inline",
    }
    if config.public_read:
        headers["x-oss-object-acl"] = "public-read"

    signature = sign(config, "PUT", key, headers, content_md5, ctype)
    headers["Authorization"] = f"OSS {config.access_key_id}:{signature}"

    host = f"{config.bucket}.{endpoint_host(config.endpoint)}"
    url = f"https://{host}/{quote_key(key)}"
    request = urllib.request.Request(url, data=body, headers=headers, method="PUT")
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            if response.status not in {200, 201}:
                fail(f"upload failed for {local_path}: HTTP {response.status}")
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        fail(f"upload failed for {local_path}: HTTP {exc.code}\n{details}")
    except urllib.error.URLError as exc:
        fail(f"upload failed for {local_path}: {exc.reason}")


def signed_get_url(config: OssConfig, key: str) -> str:
    expires = int(
        (datetime.now(timezone.utc) + timedelta(days=config.signed_url_expires_days)).timestamp()
    )
    response_headers = {
        "response-content-disposition": "inline",
    }
    canonicalized_resource = (
        f"/{config.bucket}/{key}?"
        + "&".join(f"{name}={response_headers[name]}" for name in sorted(response_headers))
    )
    string_to_sign = f"GET\n\n\n{expires}\n{canonicalized_resource}"
    digest = hmac.new(
        config.access_key_secret.encode("utf-8"),
        string_to_sign.encode("utf-8"),
        hashlib.sha1,
    ).digest()
    signature = base64.b64encode(digest).decode("ascii")
    query = {
        "OSSAccessKeyId": config.access_key_id,
        "Expires": str(expires),
        "Signature": signature,
        **response_headers,
    }
    if config.public_base_url:
        base_url = config.public_base_url
    else:
        base_url = f"https://{config.bucket}.{endpoint_host(config.endpoint)}"
    return f"{base_url}/{quote_key(key)}?{urllib.parse.urlencode(query)}"


def iter_files(report_dir: Path) -> list[Path]:
    return sorted(path for path in report_dir.rglob("*") if path.is_file())


def upload(report_dir: Path, config_path: Path = DEFAULT_CONFIG) -> str:
    if not report_dir.exists() or not report_dir.is_dir():
        fail(f"report directory not found: {report_dir}")
    if not (report_dir / "index.html").exists():
        fail(f"report directory has no index.html: {report_dir}")

    config = load_config(config_path)
    prefix = config.remote_prefix.strip("/") or report_dir.name
    files = iter_files(report_dir)
    if not files:
        fail(f"report directory is empty: {report_dir}")

    for path in files:
        key = object_key(prefix, path.relative_to(report_dir))
        put_object(config, path, key)
        print(f"uploaded {path.relative_to(report_dir)} -> oss://{config.bucket}/{key}")

    index_key = object_key(prefix, Path("index.html"))
    if config.public_base_url:
        public_base_url = config.public_base_url
    else:
        public_base_url = f"https://{config.bucket}.{endpoint_host(config.endpoint)}"
    public_url = f"{public_base_url}/{quote_key(index_key)}"
    inline_url = signed_get_url(config, index_key)
    print(f"url: {public_url}")
    print(f"inline_url: {inline_url}")
    return inline_url


def main(argv: list[str]) -> None:
    if len(argv) not in {2, 3}:
        fail("usage: upload_to_oss.py report-folder [config-json]")
    report_dir = Path(argv[1]).expanduser().resolve()
    config_path = Path(argv[2]).expanduser().resolve() if len(argv) == 3 else DEFAULT_CONFIG
    upload(report_dir, config_path)


if __name__ == "__main__":
    main(sys.argv)
