# -*- coding: utf-8 -*-
"""编制依据运行时联网核验。

按《运行时联网核验规范.md》的口径做“可执行的最小核验”：
- 只对库内提供原文URL的条目发起直连请求，不臆造检索结果；
- 页面可达且文号/标准号与名称关键词同时命中，记“一致”；
- 页面可达但字段未命中，记“内容有更新或页面不完整”；
- 网络不可达、超时、非HTML，记“未能联网”；
- 结果写入 依据核验记录.json 与 依据核对表.md，供成文脚本与内部件使用。

本脚本不修改依据库原文件；发现更新时只写“更新建议”，由库维护方回灌。
"""
from __future__ import annotations

import argparse
import json
import re
import ssl
import sys
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

PENDING_IDS = {"C02", "C04"}
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
OFFICIAL_HINTS = (".gov.cn", "gov.cn", "pkulaw.com", "npc.gov.cn", "ndrc.gov.cn",
                  "mohurd.gov.cn", "example.gov.cn", "example.gov.cn")


def _norm(s: str) -> str:
    s = s or ""
    s = s.replace("〔", "[").replace("〕", "]").replace("［", "[").replace("］", "]")
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"\s+", "", s)
    return s


def _decode(raw: bytes, charset: str | None) -> str:
    for enc in [charset, "utf-8", "gb18030"]:
        if not enc:
            continue
        try:
            return raw.decode(enc, errors="ignore")
        except Exception:
            continue
    return raw.decode("utf-8", errors="ignore")


def _name_fragments(name: str) -> list[str]:
    name = re.sub(r"[《》]", "", name or "")
    name = re.sub(r"^.*?关于", "", name)
    name = re.sub(r"的通知$|的指导意见$|的实施意见$|的批复$", "", name)
    frags = []
    if len(name) >= 6:
        frags.append(name[:6])
    if len(name) >= 10:
        frags.append(name[-8:])
    if len(name) >= 4:
        frags.append(name[:4])
    return [x for x in dict.fromkeys(frags) if len(x) >= 4]


def check_entry(entry: dict, timeout: float) -> dict:
    eid = entry.get("id")
    name = entry.get("名称") or ""
    url = (entry.get("原文URL") or entry.get("核验来源") or "").strip()
    no = (entry.get("文号") or "").strip()
    base = {"id": eid, "name": name, "url": url, "number": no,
            "checked_at": datetime.now().astimezone().isoformat(timespec="seconds")}
    if not url:
        return {**base, "status": "未能联网", "http": None, "matched": [],
                "note": "库内无原文URL，未发起联网核验"}
    host = urllib.parse.urlparse(url).netloc.lower()
    if not any(h in host for h in OFFICIAL_HINTS):
        return {**base, "status": "未能联网", "http": None, "matched": [],
                "note": f"非官方域名（{host}），不在本次直连核验范围"}
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "text/html,*/*"})
    ctx = ssl.create_default_context()
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
            http = getattr(resp, "status", 200)
            ctype = (resp.headers.get("Content-Type") or "").lower()
            raw = resp.read(2_000_000)
    except urllib.error.HTTPError as e:
        return {**base, "status": "未能联网", "http": e.code, "matched": [],
                "note": f"HTTP {e.code}"}
    except Exception as e:
        return {**base, "status": "未能联网", "http": None, "matched": [],
                "note": f"{type(e).__name__}: {str(e)[:120]}"}
    if "pdf" in ctype or url.lower().endswith(".pdf"):
        return {**base, "status": "未能联网", "http": http, "matched": [],
                "note": "原文为PDF，未做逐字比对"}
    text = _decode(raw, None)
    nt = _norm(text)
    matched = []
    no_ok = bool(no) and not no.startswith(("未", "未见")) and _norm(no) in nt
    if no_ok:
        matched.append(_norm(no))
    frag_ok = [f for f in _name_fragments(name) if _norm(f) in nt]
    if frag_ok:
        matched.append(frag_ok[0])
    if http == 200 and no_ok and frag_ok:
        status, note = "一致", "直连原文页，文号与名称关键词均命中"
    elif http == 200 and (no_ok or frag_ok):
        status, note = "内容有更新或页面不完整", "仅部分字段命中，需人工复核"
    elif http == 200:
        status, note = "未能联网", "页面可达但正文未命中，按未复核处理"
    else:
        status, note = "未能联网", f"HTTP {http}"
    return {**base, "status": status, "http": http, "matched": matched, "note": note}


def write_md(ledger: dict, out: Path) -> None:
    lines = ["# 编制依据核对表", "",
             f"- 核验时间：{ledger.get('checked_at')}",
             f"- 核验方式：{ledger.get('method')}",
             f"- 依据库：{ledger.get('library')}",
             f"- 汇总：{json.dumps(ledger.get('summary'), ensure_ascii=False)}", "",
             "| 编号 | 名称 | 文号 | 本次结果 | 原文链接 | 说明 |",
             "| --- | --- | --- | --- | --- | --- |"]
    for r in ledger.get("results", []):
        lines.append("| {id} | {name} | {number} | {status} | {url} | {note} |".format(
            id=r.get("id", ""), name=(r.get("name") or "").replace("|", "／"),
            number=(r.get("number") or "").replace("|", "／"),
            status=r.get("status", ""), url=r.get("url", ""), note=(r.get("note") or "").replace("|", "／")))
    out.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--lib", required=True)
    ap.add_argument("--out-dir", required=True)
    ap.add_argument("--timeout", type=float, default=8.0)
    ap.add_argument("--workers", type=int, default=8)
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    lib_path = Path(args.lib)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    lib = json.loads(lib_path.read_text(encoding="utf-8"))
    entries = [e for e in (lib.get("条目") or []) if e.get("id") not in PENDING_IDS]
    if args.limit:
        entries = entries[:args.limit]

    results = []
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as ex:
        futs = {ex.submit(check_entry, e, args.timeout): e for e in entries}
        for f in as_completed(futs):
            try:
                results.append(f.result())
            except Exception as e:
                ent = futs[f]
                results.append({"id": ent.get("id"), "name": ent.get("名称"),
                                "url": ent.get("原文URL", ""), "number": ent.get("文号", ""),
                                "status": "未能联网", "http": None, "matched": [],
                                "note": f"核验异常：{type(e).__name__}: {str(e)[:100]}",
                                "checked_at": datetime.now().astimezone().isoformat(timespec="seconds")})
    order = {e.get("id"): i for i, e in enumerate(lib.get("条目") or [])}
    results.sort(key=lambda r: order.get(r.get("id"), 999))
    summary = {
        "total": len(results),
        "一致": sum(1 for r in results if r.get("status") == "一致"),
        "内容有更新或页面不完整": sum(1 for r in results if r.get("status") == "内容有更新或页面不完整"),
        "未能联网": sum(1 for r in results if r.get("status") == "未能联网"),
    }
    ledger = {
        "schema": "cardo/basis-verify/v1",
        "checked_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "method": "直连原文URL复核（未执行搜索引擎检索；PDF/非官方域按未复核处理）",
        "library": str(lib_path),
        "summary": summary,
        "results": results,
    }
    (out_dir / "依据核验记录.json").write_text(
        json.dumps(ledger, ensure_ascii=False, indent=2), encoding="utf-8")
    write_md(ledger, out_dir / "依据核对表.md")
    print(json.dumps(summary, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
