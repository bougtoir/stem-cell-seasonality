#!/usr/bin/env python3
"""
Comprehensive verification of the Human Reproduction Opinion manuscript.

Checks:
1. All figures/tables are cited in manuscript body text
2. Figure/table numbering follows order of first appearance
3. All in-text citations have matching reference entries
4. All reference entries are cited at least once in text
5. Reference existence via Crossref API
"""

import re
import sys
import json
import time
from datetime import datetime, timezone
from pathlib import Path

try:
    import requests
except ImportError:
    print("ERROR: 'requests' package required")
    sys.exit(1)

SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR.parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

CROSSREF_API = "https://api.crossref.org/works"
MAILTO = "bougtoir@gmail.com"
RATE_LIMIT = 1.0


def load_module():
    """Import create_humrep_opinion_docx to access REFERENCES and text."""
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "humrep", SCRIPT_DIR / "create_humrep_opinion_docx.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def extract_body_text(mod):
    """Extract all body text from the generated docx manuscript."""
    from docx import Document
    docx_path = OUTPUT_DIR / "HumReprod_Opinion_InvisibleVariables.docx"
    if not docx_path.exists():
        # Generate manuscript first
        mod.create_manuscript()
    doc = Document(str(docx_path))
    return "\n".join(p.text for p in doc.paragraphs)


def extract_citations_from_source(source):
    """Extract all (Author, Year) style citations from source code strings."""
    # Match patterns like (Author et al., 2025) or (Author and Author, 2020)
    # or (Author, 2020) or multi-cite (Author et al., 2023; Author et al., 2020)
    # Also match bare mentions like "Author et al. (2025)"
    citations = set()

    # Pattern 1: parenthetical citations "(Author et al., Year)"
    paren_pattern = r'\(([^)]+)\)'
    for match in re.finditer(paren_pattern, source):
        content = match.group(1)
        # Split by semicolon for multi-citations
        for part in content.split(';'):
            part = part.strip()
            # Match "Author et al., Year" or "Author and Author, Year" or "Author, Year"
            cite_match = re.match(
                r'([A-Z][a-zA-Z\-]+(?:\s+(?:et\s+al\.|and\s+[A-Z][a-zA-Z\-]+))?),\s*(\d{4})',
                part)
            if cite_match:
                citations.add(f"{cite_match.group(1)}, {cite_match.group(2)}")

    # Pattern 2: narrative citations "Author et al. (Year)" or "Author (Year)"
    narrative_pattern = r'([A-Z][a-zA-Z\-]+(?:\s+et\s+al\.)?)\s*\((\d{4})\)'
    for match in re.finditer(narrative_pattern, source):
        author = match.group(1)
        year = match.group(2)
        if author not in ('Figure', 'Table', 'Phase', 'Panel', 'Section', 'Cycle',
                          'Solar', 'Supplementary'):
            citations.add(f"{author}, {year}")

    return citations


def extract_figure_mentions(source):
    """Extract all Figure N mentions from source and return order of first appearance."""
    # Match "Figure 1", "Figure 2", etc.
    fig_pattern = r'Figure\s+(\d+)'
    figures = []
    seen = set()
    for match in re.finditer(fig_pattern, source):
        fig_num = int(match.group(1))
        if fig_num not in seen:
            figures.append(fig_num)
            seen.add(fig_num)
    return figures


def extract_table_mentions(source):
    """Extract all Table N mentions from source."""
    table_pattern = r'Table\s+(\d+)'
    tables = []
    seen = set()
    for match in re.finditer(table_pattern, source):
        tbl_num = int(match.group(1))
        if tbl_num not in seen:
            tables.append(tbl_num)
            seen.add(tbl_num)
    return tables


def match_citation_to_ref(cite_str, references):
    """Try to match a citation string to a reference entry."""
    for ref in references:
        ref_cite = ref["cite"]
        if cite_str.strip() == ref_cite.strip():
            return ref
        # Normalize: remove "et al." variations
        cite_norm = cite_str.replace(" et al.", "").replace("et al.", "").strip()
        ref_norm = ref_cite.replace(" et al.", "").replace("et al.", "").strip()
        if cite_norm == ref_norm:
            return ref
    return None


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


def verify_reference_exists(ref):
    """Verify reference exists via Crossref API."""
    if ref.get("type") == "book":
        return {"status": "BOOK_SKIPPED", "reason": "Book — manual verification needed"}

    title = ref.get("title", "")
    first_author = ref["authors"].split(",")[0].split()[0].strip()

    params = {
        "query.bibliographic": title,
        "query.author": first_author,
        "rows": 3,
        "mailto": MAILTO,
    }

    try:
        resp = requests.get(CROSSREF_API, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        items = data.get("message", {}).get("items", [])
        if not items:
            return {"status": "NOT_FOUND", "reason": "No Crossref results"}

        # Check all returned items (not just first) for a title match
        best_match = None
        best_overlap = 0.0
        for item in items:
            item_title = item.get("title", [""])[0] if item.get("title") else ""
            ov = title_overlap(title, item_title)
            if ov > best_overlap:
                best_overlap = ov
                best_match = item

        cr = best_match if best_match else items[0]
        cr_title = cr.get("title", [""])[0] if cr.get("title") else ""
        cr_doi = cr.get("DOI", "")
        cr_authors = cr.get("author", [])
        cr_first = cr_authors[0].get("family", "") if cr_authors else ""

        overlap = title_overlap(title, cr_title)

        issues = []
        if overlap < 0.5:
            issues.append(f"Title overlap low ({overlap:.0%})")
        if cr_first and first_author and cr_first.lower() != first_author.lower():
            issues.append(f"Author mismatch: MS={first_author} vs CR={cr_first}")

        if issues:
            return {
                "status": "MISMATCH",
                "doi": cr_doi,
                "crossref_title": cr_title,
                "issues": issues,
            }

        return {
            "status": "VERIFIED",
            "doi": cr_doi,
            "crossref_title": cr_title,
        }

    except Exception as e:
        return {"status": "API_ERROR", "reason": str(e)}


def main():
    print("=" * 70)
    print("HUMAN REPRODUCTION MANUSCRIPT VERIFICATION")
    print("=" * 70)

    mod = load_module()
    refs = mod.REFERENCES
    source = extract_body_text(mod)

    all_issues = []
    all_results = {}

    # ──────────────────────────────────────────────
    # 1. Figure/Table citation check
    # ──────────────────────────────────────────────
    print("\n[1] FIGURE/TABLE CITATION CHECK")
    print("-" * 40)

    figures_in_text = extract_figure_mentions(source)
    tables_in_text = extract_table_mentions(source)

    # Check figure legends section
    fig_legend_nums = set()
    legend_pattern = r'Figure\s+(\d+)\.'
    # Look in figure legend section specifically
    for match in re.finditer(legend_pattern, source):
        fig_legend_nums.add(int(match.group(1)))

    print(f"Figures mentioned in body text: {figures_in_text}")
    print(f"Figures with legends: {sorted(fig_legend_nums)}")
    print(f"Tables mentioned in body text: {tables_in_text}")

    # Check for figures with legends but not cited in body
    body_section = source.split("Figure Legends")[0] if "Figure Legends" in source else source
    body_figs = extract_figure_mentions(body_section)

    for fig_num in sorted(fig_legend_nums):
        if fig_num not in body_figs:
            issue = f"Figure {fig_num} has a legend but is NOT cited in body text"
            all_issues.append(("CRITICAL", issue))
            print(f"  CRITICAL: {issue}")

    for fig_num in body_figs:
        if fig_num not in fig_legend_nums:
            issue = f"Figure {fig_num} cited in text but has NO legend"
            all_issues.append(("CRITICAL", issue))
            print(f"  CRITICAL: {issue}")

    if body_figs and list(body_figs) == sorted(body_figs):
        print(f"  OK: Figures cited in sequential order: {body_figs}")
    elif body_figs:
        issue = f"Figures NOT in sequential order: {body_figs}"
        all_issues.append(("WARNING", issue))
        print(f"  WARNING: {issue}")

    if not tables_in_text:
        print("  OK: No tables in manuscript (none expected)")

    all_results["figures_in_body"] = body_figs
    all_results["figures_with_legends"] = sorted(fig_legend_nums)
    all_results["tables_in_body"] = tables_in_text

    # ──────────────────────────────────────────────
    # 2. Figure/Table numbering order check
    # ──────────────────────────────────────────────
    print("\n[2] FIGURE/TABLE NUMBERING ORDER CHECK")
    print("-" * 40)

    if body_figs:
        expected = list(range(1, max(body_figs) + 1))
        if body_figs == expected:
            print(f"  OK: Figures numbered consecutively: {body_figs}")
        else:
            issue = f"Figure numbering gap: expected {expected}, got {body_figs}"
            all_issues.append(("WARNING", issue))
            print(f"  WARNING: {issue}")

        # Check first appearance order
        for i in range(len(body_figs) - 1):
            if body_figs[i] > body_figs[i + 1]:
                issue = (f"Figure {body_figs[i+1]} appears after "
                         f"Figure {body_figs[i]} in text")
                all_issues.append(("CRITICAL", issue))
                print(f"  CRITICAL: {issue}")
                break
        else:
            print(f"  OK: Figures appear in ascending order in text")

    # ──────────────────────────────────────────────
    # 3. In-text citation ↔ Reference list cross-check
    # ──────────────────────────────────────────────
    print("\n[3] IN-TEXT CITATION ↔ REFERENCE LIST CROSS-CHECK")
    print("-" * 40)

    # Extract citations from body text (before Figure Legends)
    body_source = source.split("Figure Legends")[0] if "Figure Legends" in source else source
    citations = extract_citations_from_source(body_source)

    print(f"Unique in-text citations found: {len(citations)}")
    print(f"References in list: {len(refs)}")

    # Check each citation has a matching reference
    unmatched_citations = []
    matched_ref_keys = set()
    for cite in sorted(citations):
        ref = match_citation_to_ref(cite, refs)
        if ref:
            matched_ref_keys.add(ref["key"])
        else:
            unmatched_citations.append(cite)

    if unmatched_citations:
        for cite in unmatched_citations:
            issue = f"In-text citation '{cite}' has no matching reference entry"
            all_issues.append(("CRITICAL", issue))
            print(f"  CRITICAL: {issue}")
    else:
        print("  OK: All in-text citations match reference entries")

    # Check each reference is cited in text
    uncited_refs = []
    for ref in refs:
        if ref["key"] not in matched_ref_keys:
            uncited_refs.append(ref)

    if uncited_refs:
        for ref in uncited_refs:
            issue = f"Reference '{ref['key']}' ({ref['cite']}) is NOT cited in text"
            all_issues.append(("CRITICAL", issue))
            print(f"  CRITICAL: {issue}")
    else:
        print("  OK: All references are cited in body text")

    all_results["citations_found"] = sorted(citations)
    all_results["unmatched_citations"] = unmatched_citations
    all_results["uncited_references"] = [r["key"] for r in uncited_refs]

    # ──────────────────────────────────────────────
    # 4. Reference existence verification (Crossref)
    # ──────────────────────────────────────────────
    print("\n[4] REFERENCE EXISTENCE VERIFICATION (Crossref)")
    print("-" * 40)

    ref_results = []
    for i, ref in enumerate(refs):
        print(f"  [{i+1}/{len(refs)}] {ref['key']}... ", end="", flush=True)
        result = verify_reference_exists(ref)
        result["key"] = ref["key"]
        result["cite"] = ref["cite"]
        result["title"] = ref.get("title", "")
        ref_results.append(result)
        print(result["status"])

        if result["status"] == "NOT_FOUND":
            issue = f"Reference '{ref['key']}' NOT FOUND on Crossref"
            all_issues.append(("CRITICAL", issue))
        elif result["status"] == "MISMATCH":
            severity = "WARNING"
            for iss in result.get("issues", []):
                if "Title overlap low" in iss:
                    severity = "CRITICAL"
            issue = (f"Reference '{ref['key']}' MISMATCH: "
                     f"{'; '.join(result.get('issues', []))}")
            all_issues.append((severity, issue))

        if ref.get("type") != "book":
            time.sleep(RATE_LIMIT)

    all_results["reference_verification"] = ref_results

    # ──────────────────────────────────────────────
    # Summary
    # ──────────────────────────────────────────────
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)

    critical_count = sum(1 for s, _ in all_issues if s == "CRITICAL")
    warning_count = sum(1 for s, _ in all_issues if s == "WARNING")

    verified = sum(1 for r in ref_results if r["status"] == "VERIFIED")
    mismatched = sum(1 for r in ref_results if r["status"] == "MISMATCH")
    not_found = sum(1 for r in ref_results if r["status"] == "NOT_FOUND")
    book_skip = sum(1 for r in ref_results if r["status"] == "BOOK_SKIPPED")
    api_err = sum(1 for r in ref_results if r["status"] == "API_ERROR")

    print(f"\nFigures: {len(body_figs)} cited, {len(fig_legend_nums)} legends")
    print(f"Tables: {len(tables_in_text)} (none expected)")
    print(f"Citations: {len(citations)} unique in-text, {len(refs)} in reference list")
    print(f"  Unmatched citations: {len(unmatched_citations)}")
    print(f"  Uncited references: {len(uncited_refs)}")
    print(f"Reference verification: {verified} VERIFIED, {mismatched} MISMATCH, "
          f"{not_found} NOT_FOUND, {book_skip} BOOK, {api_err} API_ERROR")
    print(f"\nTotal issues: {critical_count} CRITICAL, {warning_count} WARNING")

    if all_issues:
        print("\nAll issues:")
        for severity, issue in all_issues:
            print(f"  [{severity}] {issue}")

    # Save report
    report = {
        "generated": datetime.now(timezone.utc).isoformat(),
        "figures": {
            "in_body": body_figs,
            "with_legends": sorted(fig_legend_nums),
            "sequential": body_figs == sorted(body_figs),
        },
        "tables": {"in_body": tables_in_text},
        "citations": {
            "in_text_count": len(citations),
            "reference_count": len(refs),
            "unmatched": unmatched_citations,
            "uncited": [r["key"] for r in uncited_refs],
        },
        "reference_verification": ref_results,
        "issues": [{"severity": s, "message": m} for s, m in all_issues],
        "summary": {
            "critical": critical_count,
            "warning": warning_count,
        },
    }

    json_path = OUTPUT_DIR / "humrep_manuscript_verification.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False)
    print(f"\nJSON report: {json_path}")

    # MD report
    md_lines = [
        "# Human Reproduction Manuscript Verification Report\n",
        f"**Generated:** {report['generated']}\n",
        "## Figure/Table Check\n",
        f"- Figures cited in body: {body_figs}",
        f"- Figures with legends: {sorted(fig_legend_nums)}",
        f"- Sequential order: {'YES' if body_figs == sorted(body_figs) else 'NO'}",
        f"- Tables: {len(tables_in_text)} (none expected)\n",
        "## Citation Cross-Check\n",
        f"- Unique in-text citations: {len(citations)}",
        f"- References in list: {len(refs)}",
        f"- Unmatched citations: {len(unmatched_citations)}",
        f"- Uncited references: {len(uncited_refs)}\n",
    ]

    if unmatched_citations:
        md_lines.append("### Unmatched Citations\n")
        for c in unmatched_citations:
            md_lines.append(f"- `{c}`")
        md_lines.append("")

    if uncited_refs:
        md_lines.append("### Uncited References\n")
        for r in uncited_refs:
            md_lines.append(f"- `{r['key']}` ({r['cite']})")
        md_lines.append("")

    md_lines.append("## Reference Existence Verification\n")
    md_lines.append("| Key | Status | DOI | Notes |")
    md_lines.append("|-----|--------|-----|-------|")
    for r in ref_results:
        doi = r.get("doi", "")
        notes = "; ".join(r.get("issues", []))[:60] if r.get("issues") else r.get("reason", "")
        md_lines.append(f"| {r['key']} | {r['status']} | {doi} | {notes} |")

    if all_issues:
        md_lines.append("\n## Issues\n")
        for severity, issue in all_issues:
            md_lines.append(f"- **[{severity}]** {issue}")

    md_path = OUTPUT_DIR / "humrep_manuscript_verification.md"
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))
    print(f"MD report: {md_path}")

    return critical_count == 0


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
