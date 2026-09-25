"""Publish the products in products/catalog.json to Polar.

POLAR_TOKEN=... python tools/polar_publish.py            # production
POLAR_API=https://sandbox-api.polar.sh POLAR_TOKEN=... python tools/polar_publish.py

For each product: upload the download file and the cover, create the product,
a "downloadables" benefit, attach it, and create a checkout link. Products that
already exist (same name) are skipped. Results go to products/published.json.
"""

import base64
import hashlib
import json
import mimetypes
import os
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
API = os.environ.get("POLAR_API", "https://api.polar.sh")
TOKEN = os.environ["POLAR_TOKEN"]
UA = {"User-Agent": "moneybot-polar-publisher/1.0 (+https://github.com/Someone006/MoneyBot)"}
mimetypes.add_type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx")


def call(method, path, body=None):
    req = urllib.request.Request(
        API + path,
        data=json.dumps(body).encode() if body is not None else None,
        method=method,
        headers={"Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json", "Accept": "application/json", **UA},
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            raw = r.read()
            return json.loads(raw) if raw else None
    except urllib.error.HTTPError as e:
        raise SystemExit(f"{method} {path} -> {e.code}: {e.read().decode()[:1500]}")


def upload(path: Path, service: str) -> str:
    data = path.read_bytes()
    sha = base64.b64encode(hashlib.sha256(data).digest()).decode()
    f = call("POST", "/v1/files/", {
        "name": path.name,
        "mime_type": mimetypes.guess_type(path.name)[0] or "application/octet-stream",
        "size": len(data),
        "checksum_sha256_base64": sha,
        "service": service,
        "upload": {"parts": [{"number": 1, "chunk_start": 0, "chunk_end": len(data), "checksum_sha256_base64": sha}]},
    })
    part = f["upload"]["parts"][0]
    put = urllib.request.Request(part["url"], data=data, method="PUT", headers=part.get("headers") or {})
    with urllib.request.urlopen(put, timeout=120) as r:
        etag = r.headers["ETag"]
    call("POST", f"/v1/files/{f['id']}/uploaded", {
        "id": f["upload"]["id"],
        "path": f["upload"]["path"],
        "parts": [{"number": 1, "checksum_etag": etag, "checksum_sha256_base64": sha}],
    })
    return f["id"]


def publish(item: dict, existing: dict) -> dict:
    if item["name"] in existing:
        print("skip (exists):", item["name"])
        return existing[item["name"]]
    file_ids = [upload(ROOT / f, "downloadable") for f in item.get("files") or [item["file"]]]
    media_id = upload(ROOT / item["cover"], "product_media")
    product = call("POST", "/v1/products/", {
        "name": item["name"],
        "description": item["description"],
        "prices": [
            {"amount_type": "fixed", "price_amount": item["prices"][cur], "price_currency": cur}
            for cur in dict.fromkeys(["chf", DEFAULT_CURRENCY])
        ],
        "medias": [media_id],
    })
    benefit = call("POST", "/v1/benefits/", {
        "type": "downloadables",
        "description": item["benefit"],
        "properties": {"files": file_ids},
    })
    call("POST", f"/v1/products/{product['id']}/benefits", {"benefits": [benefit["id"]]})
    link = call("POST", "/v1/checkout-links/", {
        "payment_processor": "stripe",
        "products": [product["id"]],
        "allow_discount_codes": True,
    })
    print("published:", item["name"], "->", link["url"])
    return {"product_id": product["id"], "checkout_url": link["url"]}


DEFAULT_CURRENCY = "usd"


def main():
    global DEFAULT_CURRENCY
    org = call("GET", "/v1/organizations/")["items"][0]
    DEFAULT_CURRENCY = org.get("default_presentment_currency") or "usd"
    catalog = json.loads((ROOT / "products/catalog.json").read_text())
    existing = {p["name"]: {"product_id": p["id"]} for p in call("GET", "/v1/products/?limit=100")["items"]}
    out_path = ROOT / "products/published.json"
    results = json.loads(out_path.read_text()) if out_path.exists() else {}
    env = "sandbox" if "sandbox" in API else "production"
    results.setdefault(env, {})
    for item in catalog:
        results[env][item["name"]] = {**results[env].get(item["name"], {}), **publish(item, existing)}
    out_path.write_text(json.dumps(results, indent=2, ensure_ascii=False) + "\n")


if __name__ == "__main__":
    main()
