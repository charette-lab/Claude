#!/usr/bin/env python3
"""subtags.py — derive the IPS granular sub-tags from existing research notes.

The risk framework (Investment Policy Statement) defines 11 exposure CATEGORIES,
each with a short list of granular SUB-TAG values, and runs its 20% factor-defense
rule on the SUB-TAGS (a software co, an airline and a luxury retailer are all "US
Consumer Discretionary" — that is the correlated wipeout the rule prevents). The
research backend (analyst.py) already scored the 11 binary categories AND wrote
rich `risk_notes` describing the specific exposure — but the sub-tag is left in
the prose. This module extracts the structured sub-tag from those notes, so the
factor rule can run at the granularity the framework intends.

It re-uses work already paid for: no web search, just a cheap classification of
text we already have. Names whose research records aren't in the cache cannot be
sub-tagged here (their notes live in the original cache).

    python3 subtags.py --out universe_run/sub_tags.json [--model claude-haiku-4-5] [--limit N]
"""
from __future__ import annotations
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from frameworks import RISK_TAGS_SHORT

CACHE_DIR = os.environ.get("NCP_CACHE", os.path.join(os.path.dirname(__file__), ".cache"))

# IPS category (in the engine's RISK_TAGS_SHORT order) -> allowed sub-tag values.
TAXONOMY = {
    "Demand":       ["US Consumer Discretionary", "Enterprise IT CapEx", "SMBs",
                     "Government/Defense", "Healthcare Providers"],
    "CustConc":     ["Single-Customer Reliant (>10% Rev)", "Top 5 Customers (>25% Rev)",
                     "Highly Fragmented Base"],
    "Macro":        ["Highly Rate-Sensitive", "Inflation Beneficiary", "Commodity-Dependent"],
    "FX":           ["High International Revenue (>50%)", "Emerging Market FX Exposure",
                     "Domestic US Revenue Only"],
    "Geo/SC":       ["China Manufacturing/Supply", "European End-Markets", "Domestic US Only",
                     "Taiwan Semiconductor Reliance"],
    "Funding":      ["Highly Leveraged", "Serial Diluter", "Cash-Flow Positive/Self-Funding"],
    "TechDep":      ["Apple/iOS Dependent", "Cloud Infrastructure Reliant", "AI Disruption Risk",
                     "Legacy On-Premise"],
    "Lifecycle/IP": ["Patent Cliff Approaching (<3 Years)", "Single-Blockbuster Dependent",
                     "IP/Copyright Heavy"],
    "Regulatory":   ["Antitrust Target", "FDA/Healthcare Policy", "Heavy Tariff Exposure",
                     "Environmental/ESG Targets"],
    "InputCost":    ["Highly Unionized Labor/Wage Sensitive", "Energy/Fuel Intensive",
                     "Raw Material Heavy"],
    "KeyPerson":    ["Founder-Led/Visionary Dependent", "Aging/Transitioning Leadership",
                     "Institutionalized Management"],
}

SYSTEM = """You classify a company's risk exposures into a fixed taxonomy, using ONLY the
evidence already gathered in its risk notes (plus its industry/country). For each
exposure CATEGORY that is flagged for the company, choose the ONE best-matching
sub-tag value from that category's allowed list — the value that best captures the
*correlation* the analyst described (what would make this stock drop with others).

The labels are written US-centrically; read them as exposure TYPES and map to the
nearest one for non-US companies (e.g. a European consumer-cyclical maps to "US
Consumer Discretionary" as the consumer-discretionary-demand bucket; a non-US
single-country revenue base maps to "Domestic US Revenue Only" as the single-home-
market bucket). For FX and Geographic categories, weight the company's actual
country/supply geography. If the notes genuinely support two distinct sub-tags in
one category, you may return up to two. Only classify flagged categories.

Return ONLY a JSON object mapping each flagged category to a list of 1-2 chosen
sub-tag values, taken verbatim from the allowed lists. No prose."""


def _flagged(rec):
    tags = rec.get("risk_tags") or []
    return [RISK_TAGS_SHORT[i] for i, v in enumerate(tags)
            if i < len(RISK_TAGS_SHORT) and v]


def _user_prompt(rec):
    cats = _flagged(rec)
    allowed = {c: TAXONOMY[c] for c in cats if c in TAXONOMY}
    return (f"Company: {rec.get('company')}  ({rec.get('ticker')})\n"
            f"Industry: {rec.get('industry') or rec.get('core_business') or '?'}\n"
            f"Country: {rec.get('country') or '?'}\n"
            f"Flagged categories and their allowed sub-tag values:\n"
            f"{json.dumps(allowed, indent=1)}\n\n"
            f"Risk notes (evidence):\n{rec.get('risk_notes') or ''}\n\n"
            f"Return JSON mapping each flagged category to its chosen sub-tag value(s).")


def classify(rec, client, model="claude-haiku-4-5"):
    """Return {category: [sub_tag values]} for one record's flagged categories."""
    if not (rec.get("risk_notes") or "").strip() or not _flagged(rec):
        return {}
    resp = client.messages.create(
        model=model, max_tokens=600, system=SYSTEM,
        messages=[{"role": "user", "content": _user_prompt(rec)}],
    )
    text = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
    import re
    m = re.search(r"\{.*\}", text, re.DOTALL)
    if not m:
        return {}
    try:
        raw = json.loads(m.group(0))
    except Exception:
        return {}
    # keep only allowed values, coerce to list
    out = {}
    for cat, vals in raw.items():
        if cat not in TAXONOMY:
            continue
        vals = vals if isinstance(vals, list) else [vals]
        keep = [v for v in vals if v in TAXONOMY[cat]]
        if keep:
            out[cat] = keep[:2]
    return out


def _load_records(cache_dir):
    recs = []
    for fn in os.listdir(cache_dir):
        if not fn.endswith(".json"):
            continue
        try:
            d = json.load(open(os.path.join(cache_dir, fn)))
        except Exception:
            continue
        if (d.get("risk_notes") or "").strip() and d.get("ticker"):
            recs.append(d)
    return recs


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=os.path.join(os.path.dirname(__file__), "universe_run", "sub_tags.json"))
    ap.add_argument("--model", default="claude-haiku-4-5")
    ap.add_argument("--workers", type=int, default=8)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--cache", default=CACHE_DIR)
    args = ap.parse_args()

    import anthropic
    client = anthropic.Anthropic()
    recs = _load_records(args.cache)
    if args.limit:
        recs = recs[:args.limit]
    print(f"classifying {len(recs)} records with notes (model={args.model})")

    from concurrent.futures import ThreadPoolExecutor, as_completed
    out = {}
    def work(r):
        try:
            return r["ticker"], classify(r, client, args.model)
        except Exception as e:
            return r["ticker"], {"_error": str(e)}
    done = 0
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futs = [ex.submit(work, r) for r in recs]
        for fut in as_completed(futs):
            t, st = fut.result()
            if st and "_error" not in st:
                out[t] = st
            done += 1
            if done % 25 == 0 or done == len(recs):
                print(f"  {done}/{len(recs)} | tagged {len(out)}", flush=True)
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    json.dump(out, open(args.out, "w"), indent=1)
    print(f"wrote sub-tags for {len(out)} names -> {args.out}")


if __name__ == "__main__":
    main()
