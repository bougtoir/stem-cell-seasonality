#!/usr/bin/env python3
"""
Automated reference verification for the Human Reproduction Opinion manuscript.

Queries each reference in the REFERENCES list against:
- Crossref API (journal articles)
- Open Library API (books)

Usage:
    python scripts/verify_references_humrep.py

Output:
    - output/humrep_reference_audit_report.md  (human-readable)
    - output/humrep_reference_audit_report.json (machine-readable)
"""

import json
import re
import time
import sys
from datetime import datetime, timezone
from pathlib import Path

try:
    import requests
except ImportError:
    print("ERROR: 'requests' package required. Install: pip install requests")
    sys.exit(1)

# ─── Configuration ────────────────────────────────────────────────────────────
CROSSREF_API = "https://api.crossref.org/works"
OPENLIBRARY_API = "https://openlibrary.org/search.json"
MAILTO = "bougtoir@gmail.com"
RATE_LIMIT_SECONDS = 1.0

SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR.parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


def load_references():
    """Import REFERENCES from create_humrep_opinion_docx.py."""
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "humrep", SCRIPT_DIR / "create_humrep_opinion_docx.py")
    mod = importlib.util.module_from_spec(spec)

    # Prevent the module from generating docx on import
    import types
    original_main = None
    spec.loader.exec_module(mod)

    return mod.REFERENCES


def query_crossref(ref):
    """Query Crossref for an article reference."""
    title = ref.get("title", "")
    first_author = ref["authors"].split(",")[0].strip()

    params = {
        "query.bibliographic": title,
        "query.author": first_author,
        "rows": 1,
        "mailto": MAILTO,
    }

    try:
        resp = requests.get(CROSSREF_API, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()

        items = data.get("message", {}).get("items", [])
        if not items:
            return None
        return items[0]
    except Exception as e:
        return {"error": str(e)}


def query_open_library(ref):
    """Query Open Library for a book reference."""
    title = ref.get("title", "")
    author = ref["authors"].split(",")[0].strip()

    params = {"title": title, "author": author, "limit": 1}
    try:
        resp = requests.get(OPENLIBRARY_API, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        docs = data.get("docs", [])
        if not docs:
            return None
        return docs[0]
    except Exception as e:
        return {"error": str(e)}


def title_overlap(t1, t2):
    """Compute word overlap ratio between two titles."""
    if not t1 or not t2:
        return 0.0
    words1 = set(re.findall(r'\w+', t1.lower()))
    words2 = set(re.findall(r'\w+', t2.lower()))
    if not words1 or not words2:
        return 0.0
    intersection = words1 & words2
    return len(intersection) / max(len(words1), len(words2))


def verify_article(ref):
    """Verify an article against Crossref."""
    cr = query_crossref(ref)
    if cr is None:
        return {"status": "UNVERIFIED", "reason": "No Crossref results", "issues": []}
    if "error" in cr:
        return {"status": "UNVERIFIED", "reason": f"API error: {cr['error']}", "issues": []}

    issues = []
    cr_title = cr.get("title", [""])[0] if cr.get("title") else ""
    cr_authors = cr.get("author", [])
    cr_first_author = cr_authors[0].get("family", "") if cr_authors else ""
    cr_journal = cr.get("container-title", [""])[0] if cr.get("container-title") else ""
    cr_year = ""
    if cr.get("published-print"):
        cr_year = str(cr["published-print"]["date-parts"][0][0])
    elif cr.get("published-online"):
        cr_year = str(cr["published-online"]["date-parts"][0][0])
    cr_volume = cr.get("volume", "")
    cr_pages = cr.get("page", "")
    cr_doi = cr.get("DOI", "")

    ms_title = ref.get("title", "")
    # Author format is "Surname Initials, ..." so first word is surname
    ms_first_author = ref["authors"].split(",")[0].split()[0].strip()
    ms_year = str(ref.get("year", ""))
    ms_volume = ref.get("volume", "")
    ms_pages = ref.get("pages", "")

    # Title overlap
    overlap = title_overlap(ms_title, cr_title)
    if overlap < 0.5:
        issues.append(f"[CRITICAL] Title mismatch (overlap={overlap:.0%}): "
                       f"MS='{ms_title[:50]}' vs CR='{cr_title[:50]}'")

    # First author
    if cr_first_author and ms_first_author:
        if cr_first_author.lower() != ms_first_author.lower():
            issues.append(f"[CRITICAL] First author: MS='{ms_first_author}' "
                           f"vs Crossref='{cr_first_author}'")

    # Year
    if cr_year and ms_year:
        year_diff = abs(int(cr_year) - int(ms_year))
        if year_diff > 1:
            issues.append(f"[WARNING] Year: MS={ms_year} vs Crossref={cr_year}")
        elif year_diff == 1:
            issues.append(f"[INFO] Year +/-1: MS={ms_year} vs Crossref={cr_year}")

    # Volume
    if cr_volume and ms_volume and cr_volume != ms_volume:
        issues.append(f"[WARNING] Volume: MS={ms_volume} vs Crossref={cr_volume}")

    # Pages
    if cr_pages and ms_pages:
        cr_pages_norm = cr_pages.replace("\u2013", "-").replace("–", "-")
        ms_pages_norm = ms_pages.replace("\u2013", "-").replace("–", "-")
        if cr_pages_norm != ms_pages_norm:
            issues.append(f"[WARNING] Pages: MS={ms_pages} vs Crossref={cr_pages}")

    # Journal name (INFO only)
    if cr_journal and ref.get("journal"):
        ms_journal = ref["journal"]
        if cr_journal.lower() != ms_journal.lower():
            issues.append(f"[INFO] Journal name: MS='{ms_journal}' vs Crossref='{cr_journal}'")

    # Determine status
    has_critical = any("[CRITICAL]" in i for i in issues)
    has_warning = any("[WARNING]" in i for i in issues)

    if has_critical:
        status = "CRITICAL_MISMATCH"
    elif has_warning:
        status = "WARNING"
    else:
        status = "MATCH"

    return {
        "status": status,
        "crossref_doi": cr_doi,
        "crossref_title": cr_title,
        "crossref_first_author": cr_first_author,
        "crossref_journal": cr_journal,
        "crossref_year": cr_year,
        "crossref_volume": cr_volume,
        "crossref_pages": cr_pages,
        "issues": issues,
    }


def verify_book(ref):
    """Verify a book against Open Library."""
    ol = query_open_library(ref)
    if ol is None:
        return {"status": "UNVERIFIED", "reason": "No Open Library results", "issues": []}
    if "error" in ol:
        return {"status": "UNVERIFIED", "reason": f"API error: {ol['error']}", "issues": []}

    issues = []
    ol_title = ol.get("title", "")
    ol_year = str(ol.get("first_publish_year", ""))

    overlap = title_overlap(ref.get("title", ""), ol_title)
    if overlap < 0.5:
        issues.append(f"[WARNING] Title overlap low ({overlap:.0%}): "
                       f"MS='{ref.get('title', '')[:40]}' vs OL='{ol_title[:40]}'")

    status = "WARNING" if issues else "MATCH"
    return {"status": status, "ol_title": ol_title, "ol_year": ol_year, "issues": issues}


def main():
    refs = load_references()
    print(f"Loaded {len(refs)} references from create_humrep_opinion_docx.py")

    results = []
    for i, ref in enumerate(refs):
        print(f"  Verifying [{i+1}/{len(refs)}] {ref['key']}...", end=" ", flush=True)

        if ref.get("type") == "book":
            result = verify_book(ref)
        else:
            result = verify_article(ref)

        result["ref_num"] = i + 1
        result["key"] = ref["key"]
        result["cite"] = ref["cite"]
        result["raw_authors"] = ref["authors"]
        result["raw_title"] = ref.get("title", "")
        result["raw_year"] = ref.get("year", "")
        results.append(result)

        print(result["status"])
        time.sleep(RATE_LIMIT_SECONDS)

    # Summary
    status_counts = {}
    for r in results:
        status_counts[r["status"]] = status_counts.get(r["status"], 0) + 1

    print(f"\n--- Summary ---")
    for s, c in sorted(status_counts.items()):
        print(f"  {s}: {c}")

    # JSON report
    report = {
        "generated": datetime.now(timezone.utc).isoformat(),
        "source": "scripts/create_humrep_opinion_docx.py",
        "total_references": len(refs),
        "summary": status_counts,
        "results": results,
    }
    json_path = OUTPUT_DIR / "humrep_reference_audit_report.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False)
    print(f"JSON report: {json_path}")

    # MD report
    md_lines = [
        "# Human Reproduction Reference Verification Audit Report\n",
        f"**Generated:** {report['generated']}",
        f"**Source:** `scripts/create_humrep_opinion_docx.py` REFERENCES list",
        f"**Total references:** {len(refs)}\n",
        "## Summary\n",
        "| Status | Count |",
        "|--------|-------|",
    ]
    for s in ["MATCH", "WARNING", "CRITICAL_MISMATCH", "UNVERIFIED"]:
        md_lines.append(f"| {s} | {status_counts.get(s, 0)} |")

    # Issues
    issues_found = [r for r in results if r["issues"]]
    if issues_found:
        md_lines.append("\n## Issues Requiring Attention\n")
        for r in issues_found:
            md_lines.append(f"### Ref {r['key']}: {r['status']}\n")
            md_lines.append(f"**Manuscript:** {r['raw_authors']}. {r['raw_title']}. {r['raw_year']}.\n")
            for issue in r["issues"]:
                md_lines.append(f"- {issue}")
            md_lines.append("")

    # Quick view
    md_lines.append("\n## All Results\n")
    md_lines.append("| Key | Status | Year | Notes |")
    md_lines.append("|-----|--------|------|-------|")
    for r in results:
        notes = "; ".join(r["issues"][:2]) if r["issues"] else ""
        if len(notes) > 60:
            notes = notes[:57] + "..."
        md_lines.append(f"| {r['key']} | {r['status']} | {r['raw_year']} | {notes} |")

    md_path = OUTPUT_DIR / "humrep_reference_audit_report.md"
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))
    print(f"MD report: {md_path}")


if __name__ == "__main__":
    main()
