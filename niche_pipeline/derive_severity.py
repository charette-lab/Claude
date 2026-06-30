#!/usr/bin/env python3
"""derive_severity.py — derive the screener's 1-4 risk-severity scores from notes.

The LTM screener codes each of the 11 IPS risk categories as a 1-4 SEVERITY score
(1 = negligible, 4 = extreme), but only for ~993 names. The analyst research wrote
`risk_notes` for many more. This derives the same 1-4 severity per category from
those existing notes — a cheap text classification (no web search), so it extends
the screener's coverage for free.

Severity rates RISK, not mere exposure presence: defensive government demand is
real exposure but LOW severity (a 1), matching how the screener scores it.

    python3 derive_severity.py --out universe_run/severity_from_notes.json
"""
from __future__ import annotations
import argparse, json, os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from frameworks import RISK_TAGS_SHORT

CACHE_DIR = os.environ.get("NCP_CACHE", os.path.join(os.path.dirname(__file__), ".cache"))

CATEGORIES = [
    ("Demand", "Customer & End-Market Demand"),
    ("CustConc", "Literal Customer Concentration"),
    ("Macro", "Macro & Factor Sensitivity"),
    ("FX", "Currency & FX Exposure"),
    ("Geo/SC", "Geographic & Supply Chain Reliance"),
    ("Funding", "Capital Structure & Funding Risk"),
    ("TechDep", "Technological & Platform Dependency"),
    ("Lifecycle/IP", "Product Lifecycle & Intellectual Property Risk"),
    ("Regulatory", "Regulatory & Compliance Risk"),
    ("InputCost", "Input Cost Structure"),
    ("KeyPerson", "Key Person & Management Risk"),
]

SYSTEM = """You score a company's exposure to 11 standard risk categories on a 1-4 RISK
SEVERITY scale, using ONLY the supplied risk notes plus its industry/country.

  1 = negligible / no meaningful risk from this factor
  2 = mild
  3 = significant
  4 = extreme — a dominant, single-factor wipeout risk

Score the RISK (likelihood and impact of this factor driving a ~30% drawdown), NOT
mere exposure presence. Example: a company whose demand is defensive government/
utility spend HAS demand exposure but it is stable, so its Demand severity is LOW
(1-2). A highly cyclical US-consumer-discretionary name scores 3-4. Calibrate so
that most companies sit at 1-2 on most factors and 4 is reserved for a genuinely
dominant risk.

Score ALL 11 categories. Return ONLY a JSON object mapping each category key to an
integer 1-4. No prose."""

CAT_KEYS = [k for k, _ in CATEGORIES]


def _user(rec):
    cats = "\n".join(f"  {k}: {full}" for k, full in CATEGORIES)
    return (f"Company: {rec.get('company')} ({rec.get('ticker')})\n"
            f"Industry: {rec.get('industry') or rec.get('core_business') or '?'}\n"
            f"Country: {rec.get('country') or '?'}\n\n"
            f"Categories (key: name):\n{cats}\n\n"
            f"Risk notes (evidence):\n{rec.get('risk_notes') or ''}\n\n"
            f"Return JSON {{category_key: severity 1-4}} for all 11 keys.")


def classify_severity(rec, client, model="claude-haiku-4-5"):
    if not (rec.get("risk_notes") or "").strip():
        return None
    resp = client.messages.create(model=model, max_tokens=300, system=SYSTEM,
                                  messages=[{"role": "user", "content": _user(rec)}])
    text = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
    m = re.search(r"\{.*\}", text, re.DOTALL)
    if not m:
        return None
    try:
        raw = json.loads(m.group(0))
    except Exception:
        return None
    out = {}
    for k in CAT_KEYS:
        v = raw.get(k)
        if isinstance(v, (int, float)) and 1 <= v <= 4:
            out[k] = int(v)
    return out or None


def _records(cache_dir):
    recs = []
    for fn in os.listdir(cache_dir):
        if fn.endswith(".json"):
            try:
                d = json.load(open(os.path.join(cache_dir, fn)))
            except Exception:
                continue
            if (d.get("risk_notes") or "").strip() and d.get("ticker"):
                recs.append(d)
    return recs


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default=os.path.join(os.path.dirname(__file__), "universe_run", "severity_from_notes.json"))
    ap.add_argument("--model", default="claude-haiku-4-5")
    ap.add_argument("--workers", type=int, default=8)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--cache", default=CACHE_DIR)
    args = ap.parse_args()
    import anthropic
    from concurrent.futures import ThreadPoolExecutor, as_completed
    client = anthropic.Anthropic()
    recs = _records(args.cache)
    if args.limit:
        recs = recs[:args.limit]
    print(f"deriving severity for {len(recs)} records (model={args.model})")
    out, done = {}, 0
    def work(r):
        try:
            return r["ticker"], classify_severity(r, client, args.model)
        except Exception:
            return r["ticker"], None
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        for fut in as_completed([ex.submit(work, r) for r in recs]):
            t, s = fut.result(); done += 1
            if s:
                out[t] = s
            if done % 50 == 0 or done == len(recs):
                print(f"  {done}/{len(recs)} | scored {len(out)}", flush=True)
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    json.dump(out, open(args.out, "w"), indent=1)
    print(f"wrote severity for {len(out)} names -> {args.out}")


if __name__ == "__main__":
    main()
