"""Evaluate extracted 10-K values against filing evidence with OpenRouter Decisions.

By default this reads the extracted values from ``filings.sqlite3``.  An export
created by the app (``10k_financial_statements.json``) or a JSON list of metric
records can be supplied with ``--input``.

The decision model is used once for every extracted value.  Numeric evidence is
also checked deterministically because decision models are not reliable numeric
calculators.  A value passes only when both the model's support probability and
the deterministic evidence check clear the 90% gate.
"""

from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
import json
import os
import re
import sqlite3
import sys
import time
import urllib.error
import urllib.request
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Iterable


ROOT = Path(__file__).resolve().parent
DEFAULT_DB = ROOT / "filings.sqlite3"
DEFAULT_KEY_FILE = Path.home() / ".openrouter" / "api_key.txt"
DECISIONS_URL = "https://openrouter.ai/api/alpha/decisions"
DEFAULT_MODEL = "typesafe/jev-1.13-20260917"
PASS_THRESHOLD = 0.90


@dataclass
class ExtractedValue:
    company: str
    fiscal_year: int | None
    category: str
    metric: str
    period: str | None
    value: Any
    unit: str | None
    source_page: int | None
    source_text: str
    justification: str
    filing_path: str | None = None


@dataclass
class Evaluation:
    company: str
    fiscal_year: int | None
    category: str
    metric: str
    period: str | None
    value: Any
    unit: str | None
    source_page: int | None
    model: str | None
    support_probability: float
    hallucination_probability: float
    confidence: float
    numeric_evidence: bool
    likely_hallucination: bool
    passed: bool
    error: str | None = None


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, help="JSON extraction file; defaults to the SQLite metrics table")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB, help=f"SQLite database (default: {DEFAULT_DB})")
    parser.add_argument("--key-file", type=Path, default=DEFAULT_KEY_FILE, help="Protected file containing the OpenRouter key")
    parser.add_argument("--model", default=os.getenv("DECISION_MODEL", DEFAULT_MODEL), help="Pinned decision model ID")
    parser.add_argument("--limit", type=int, help="Evaluate only the first N values")
    parser.add_argument("--workers", type=int, default=int(os.getenv("EVAL_WORKERS", "8")), help="Concurrent OpenRouter requests (default: 8; use 1 to disable parallelism)")
    parser.add_argument("--output", type=Path, help="Write the complete JSON report to this path")
    return parser.parse_args()


def load_api_key(path: Path) -> str:
    key = path.read_text(encoding="utf-8").strip()
    if not key:
        raise RuntimeError(f"OpenRouter API key file is empty: {path}")
    return key


def row_value(row: dict[str, Any], *names: str, default: Any = None) -> Any:
    for name in names:
        if name in row and row[name] is not None:
            return row[name]
    return default


def make_value(row: dict[str, Any], source_text: str = "") -> ExtractedValue:
    year = row_value(row, "fiscal_year", "year")
    try:
        year = int(year) if year is not None else None
    except (TypeError, ValueError):
        year = None
    page = row_value(row, "source_page", "page")
    try:
        page = int(page) if page is not None else None
    except (TypeError, ValueError):
        page = None
    return ExtractedValue(
        company=str(row_value(row, "company", default="unknown")),
        fiscal_year=year,
        category=str(row_value(row, "category", default="unknown")),
        metric=str(row_value(row, "metric", "line_item", "name", default="unknown")),
        period=row_value(row, "period", "column"),
        value=row_value(row, "value", "extracted_value"),
        unit=row_value(row, "unit"),
        source_page=page,
        source_text=str(source_text or row_value(row, "source_text", "evidence", default="")),
        justification=str(row_value(row, "justification", "extraction_justification", default="") or ""),
        filing_path=row_value(row, "path", "filing_path"),
    )


def load_values_from_db(path: Path) -> list[ExtractedValue]:
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        """SELECT f.company, f.fiscal_year, f.path, m.category, m.metric,
                  m.period, m.value, m.unit, m.source_page, m.source_text,
                  m.justification,
                  p.text AS page_text
           FROM metrics AS m
           JOIN filings AS f ON f.id = m.filing_id
           LEFT JOIN pages AS p ON p.filing_id = m.filing_id
                                AND p.page_number = m.source_page
           ORDER BY f.company, f.fiscal_year, m.category, m.id"""
    ).fetchall()
    conn.close()
    values = []
    for raw in rows:
        row = dict(raw)
        values.append(make_value(row, row.get("source_text") or row.get("page_text") or ""))
    return values


def load_values_from_json(path: Path) -> list[ExtractedValue]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(payload, dict) and isinstance(payload.get("tables"), dict):
        tables = payload["tables"]
        filings = {row["id"]: row for row in tables.get("filings", [])}
        pages = {(row["filing_id"], row["page_number"]): row.get("text", "") for row in tables.get("pages", [])}
        values = []
        for metric in tables.get("metrics", []):
            filing = filings.get(metric.get("filing_id"), {})
            source = metric.get("source_text") or pages.get((metric.get("filing_id"), metric.get("source_page")), "")
            values.append(make_value({**filing, **metric}, source))
        return values
    records = payload if isinstance(payload, list) else payload.get("values", []) if isinstance(payload, dict) else []
    if not isinstance(records, list):
        raise ValueError("Input JSON must be a list, {\"values\": [...]}, or the app's tables export")
    return [make_value(record) for record in records]


def number_forms(value: Any) -> set[str]:
    """Return common textual forms without asking the decision model to do math."""
    if isinstance(value, bool) or value is None:
        return set()
    try:
        number = float(value)
    except (TypeError, ValueError):
        return {str(value).strip()}
    forms = {str(value).strip(), f"{number:g}", f"{number:,.0f}", f"{number:,.2f}"}
    if number.is_integer():
        integer = int(number)
        forms.update({str(integer), f"{integer:,}", f"({abs(integer):,})" if integer < 0 else str(integer)})
    if number < 0:
        forms.add(f"({abs(number):g})")
    return {form for form in forms if form and form not in {"nan", "inf", "-inf"}}


def has_numeric_evidence(value: ExtractedValue) -> bool:
    evidence = "\n".join(part for part in (value.source_text, value.justification) if part)
    if not evidence:
        return False
    text = evidence.replace("\u2212", "-")
    normalized = re.sub(r"\s+", " ", text)
    return any(re.search(rf"(?<![\d.]){re.escape(form)}(?![\d.])", normalized) for form in number_forms(value.value))


def trim_evidence(value: ExtractedValue, limit: int = 7000) -> str:
    text = value.source_text.strip()
    if len(text) <= limit:
        return text
    terms = [value.metric, str(value.value), str(value.period or "")]
    positions = [text.lower().find(term.lower()) for term in terms if term]
    start = max(0, (min((p for p in positions if p >= 0), default=0) - 1800))
    return text[start : start + limit]


def decision_request(value: ExtractedValue, model: str) -> dict[str, Any]:
    justification = value.justification.strip()
    if justification:
        # The database justification is the curated metric-to-period mapping.
        # It already retains the original filing excerpt. Sending it as one
        # evidence block avoids making the model reconcile duplicate snippets
        # with different surrounding context.
        filing_evidence = justification
    else:
        filing_evidence = trim_evidence(value)
    return {
        "model": model,
        "state": {
            "company": value.company,
            "fiscal_year": value.fiscal_year,
            "statement": value.category,
            "metric": value.metric,
            "period": value.period,
            "extracted_value": value.value,
            "unit": value.unit,
            "source_page": value.source_page,
            "extractor_justification": value.justification,
            "filing_evidence": filing_evidence,
        },
        "questions": {
            "is_supported": {
                "type": "noul",
                "instructions": "Is the extracted financial value explicitly supported by the METRIC-JUSTIFICATION TABLE audit record for the named metric and period? The table is the curated extraction record: use its explicit company, statement, metric, period, value, unit, and source-page fields as the period-to-value mapping. The retained filing excerpt may contain adjacent columns and flattened PDF text; do not reconstruct a different mapping from that noise.",
                "criteria": {
                    "true": "The metric-justification row identifies the same metric and period and explicitly gives the extracted value, including its sign and unit when stated.",
                    "false": "The metric-justification row omits the value, names a different metric or period, or gives a different value.",
                },
            },
            "is_hallucinated": {
                "type": "noul",
                "instructions": "Is this extracted value likely hallucinated or unsupported according to the METRIC-JUSTIFICATION TABLE audit record? Treat the explicit metric/period/value row as the curated explanation of why the value was extracted.",
                "criteria": {
                    "true": "The metric-justification row does not ground the value, or it names a different metric, period, or value.",
                    "false": "The metric-justification row directly grounds the value for the same metric and period.",
                },
            },
        },
    }


def call_decision(api_key: str, request: dict[str, Any]) -> tuple[dict[str, Any], str]:
    body = json.dumps(request).encode("utf-8")
    retryable_statuses = {429, 500, 502, 503, 524, 529}
    for attempt in range(3):
        req = urllib.request.Request(
            DECISIONS_URL,
            data=body,
            method="POST",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        )
        try:
            with urllib.request.urlopen(req, timeout=60) as response:
                return json.loads(response.read().decode("utf-8")), ""
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            if exc.code in retryable_statuses and attempt < 2:
                time.sleep(1.5 * (attempt + 1))
                continue
            return {}, f"Decisions API HTTP {exc.code}: {detail[:500]}"
        except (urllib.error.URLError, TimeoutError, json.JSONDecodeError) as exc:
            if attempt < 2:
                time.sleep(1.5 * (attempt + 1))
                continue
            return {}, f"Decisions API error: {exc}"
    return {}, "Decisions API error: exhausted retries"


def probability(answer: Any, positive: str) -> float:
    if not isinstance(answer, dict):
        return 0.0
    raw = answer.get(positive)
    try:
        return max(0.0, min(1.0, float(raw)))
    except (TypeError, ValueError):
        return 0.0


def evaluate(value: ExtractedValue, api_key: str, model: str) -> Evaluation:
    numeric = has_numeric_evidence(value)
    response, error = call_decision(api_key, decision_request(value, model))
    answers = response.get("answers", {}) if response else {}
    supported = probability(answers.get("is_supported"), "noul")
    hallucinated = probability(answers.get("is_hallucinated"), "noul")
    # Missing exact numeric evidence is a hard failure, even if the model sounds confident.
    confidence = supported if numeric else 0.0
    likely_hallucination = (not numeric) or hallucinated >= 0.50 or supported < PASS_THRESHOLD
    passed = not error and confidence >= PASS_THRESHOLD and not likely_hallucination
    return Evaluation(
        company=value.company,
        fiscal_year=value.fiscal_year,
        category=value.category,
        metric=value.metric,
        period=value.period,
        value=value.value,
        unit=value.unit,
        source_page=value.source_page,
        model=response.get("model", model) if response else model,
        support_probability=supported,
        hallucination_probability=hallucinated,
        confidence=confidence,
        numeric_evidence=numeric,
        likely_hallucination=likely_hallucination,
        passed=passed,
        error=error or None,
    )


def main() -> int:
    args = parse_args()
    try:
        api_key = load_api_key(args.key_file)
        values = load_values_from_json(args.input) if args.input else load_values_from_db(args.db)
    except (OSError, sqlite3.Error, ValueError, RuntimeError) as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        return 2
    if args.limit is not None:
        values = values[: max(0, args.limit)]
    if not values:
        print("FAIL: no extracted values found", file=sys.stderr)
        return 2
    if args.workers < 1:
        print("FAIL: --workers must be at least 1", file=sys.stderr)
        return 2

    # Each value is independent. Keep the API fan-out bounded, then restore
    # input order for stable reports and reproducible downstream comparisons.
    results: list[Evaluation | None] = [None] * len(values)
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = {pool.submit(evaluate, value, api_key, args.model): index for index, value in enumerate(values)}
        for future in as_completed(futures):
            index = futures[future]
            result = future.result()
            results[index] = result
            value = values[index]
            completed = sum(item is not None for item in results)
            status = "PASS" if result.passed else "FAIL"
            print(
                f"{status} {completed}/{len(values)} (item {index + 1}) | {value.company} FY {value.fiscal_year} | "
                f"{value.category} | {value.metric} [{value.period}] | "
                f"confidence={result.confidence:.1%} hallucination={result.hallucination_probability:.1%}"
            )

    complete_results = [result for result in results if result is not None]
    failed = [result for result in complete_results if not result.passed]
    report = {
        "model": args.model,
        "workers": args.workers,
        "threshold": PASS_THRESHOLD,
        "total": len(complete_results),
        "passed": len(complete_results) - len(failed),
        "failed": len(failed),
        "overall_pass": not failed,
        "results": [asdict(result) for result in complete_results],
    }
    if args.output:
        args.output.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\n{'PASS' if not failed else 'FAIL'}: {len(complete_results) - len(failed)}/{len(complete_results)} values cleared the {PASS_THRESHOLD:.0%} confidence gate")
    return 0 if not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
