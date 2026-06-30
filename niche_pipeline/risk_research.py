#!/usr/bin/env python3
"""risk_research.py — deep 11-category risk assessment per company (IPS schema).

Produces, for ONE company, the same structure as the authoritative risk table:
for each of the 11 standardized exposure categories a 1-4 severity (Low/Moderate/
High/Extreme), an evidence-based assessment, a trap-check (does it escape the
framework's correlation trap), plus one company-level achilles-heel (the ~30-50%
drawdown scenario), an overall confidence, and key missing information. Uses the
Anthropic web_search server tool for real evidence; results cached to .riskcache.

Severity rates RISK (likelihood x impact of driving a ~30% drawdown), not mere
exposure presence — defensive government/utility demand is Low even though
exposure exists.

    python3 risk_research.py risk_targets.json --workers 6
"""
from __future__ import annotations
import argparse, hashlib, json, os, re, time

MODEL = os.environ.get("NCP_RISK_MODEL", "claude-opus-4-8")
RCACHE = os.environ.get("NCP_RISK_CACHE", os.path.join(os.path.dirname(__file__), ".riskcache"))

# category -> (definition, the correlation "trap" it guards against)
CATEGORIES = [
    (1, "Customer & End-Market Demand", "who pays and what drives their spend; trap: several names that all fall if one end-market (e.g. US consumer) rolls over"),
    (2, "Literal Customer Concentration", "revenue from one/few entities; trap: names that all rely on the same big customer (e.g. Apple)"),
    (3, "Macro & Factor Sensitivity", "rates/inflation/commodity force that crushes margins or multiple; trap: names that all break on the same macro move"),
    (4, "Currency & FX Exposure", "earnings sensitivity to a strong USD / EM FX; trap: names that all miss on the same FX move"),
    (5, "Geographic & Supply Chain Reliance", "physical bottleneck (China mfg, Taiwan semi, single region); trap: names hit by the same geopolitical shock"),
    (6, "Capital Structure & Funding Risk", "reliance on capital markets (leverage, dilution); trap: names that all face crises if credit freezes"),
    (7, "Technological & Platform Dependency", "whose platform/keys they depend on (Apple, cloud, AI disruption, legacy); trap: names broken by one platform owner's policy change"),
    (8, "Product Lifecycle & Intellectual Property Risk", "moat with an expiration (patent cliff, single blockbuster, IP); trap: names with synchronized exclusivity loss"),
    (9, "Regulatory & Compliance Risk", "government action that breaks the model (antitrust, FDA, tariff, ESG); trap: names hit by the same regulatory cycle"),
    (10, "Input Cost Structure", "biggest uncontrollable expense (labor, energy/fuel, raw material); trap: names whose margins all collapse on the same input spike"),
    (11, "Key Person & Management Risk", "valuation premium tied to one leader (founder, aging leadership); trap: names that de-rate if a key founder exits"),
]

SCALE = "1=Low (negligible), 2=Moderate, 3=High, 4=Extreme (dominant single-factor wipeout risk)"

SYSTEM = f"""You are a buy-side risk analyst applying an 11-category exposure framework to ONE
company. Use web search for real, current evidence (segment disclosures, customer
concentration, geography, leverage, regulation, leadership). Be evidence-based.

For EACH of the 11 categories, rate RISK SEVERITY on {SCALE}. Severity is the
likelihood and impact of THAT factor driving a ~30% drawdown — NOT mere exposure
presence (defensive/inelastic demand is Low even if exposure exists; reserve 4 for
a genuinely dominant risk). For each category give: a concise assessment, and a
trap_check stating whether the company ESCAPES the correlation trap (the framework
guards against owning several names that drop together on the same factor).

Also give ONE company-level achilles_heel (the single most likely scenario that
causes a 30-50% drawdown), overall_confidence (0-100), and key missing_information.

The 11 categories (id: name — what it catches):
""" + "\n".join(f"  {i}: {n} — {d}" for i, n, d in CATEGORIES) + """

Return ONLY one JSON object, no prose:
{
 "ric": str, "company": str, "overall_confidence": int, "achilles_heel": str,
 "risks": [
   {"risk_id": int, "risk_name": str, "risk_level": "Low|Moderate|High|Extreme",
    "risk_level_numeric": int, "assessment": str, "trap_check": str,
    "missing_information": str, "sources": [str]}
   // exactly 11, risk_id 1..11 in order
 ]
}"""


def _cache_path(ric):
    os.makedirs(RCACHE, exist_ok=True)
    return os.path.join(RCACHE, hashlib.md5((ric or "").encode()).hexdigest()[:10] + ".json")


def research_risk(ric, company, hint=None, country=None, max_searches=8, retries=2):
    cp = _cache_path(ric)
    if os.path.exists(cp):
        return json.load(open(cp))
    import anthropic
    client = anthropic.Anthropic()
    user = f"Company: {company}\nTicker/RIC: {ric}\nIndustry: {hint or '?'}\nCountry: {country or '?'}"
    last = None
    for attempt in range(retries + 1):
        try:
            resp = client.messages.create(
                model=MODEL, max_tokens=6000, system=SYSTEM,
                tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": max_searches}],
                messages=[{"role": "user", "content": user}],
            )
            text = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
            m = re.search(r"\{.*\}", text, re.DOTALL)
            rec = json.loads(m.group(0))
            rec.setdefault("ric", ric); rec.setdefault("company", company)
            json.dump(rec, open(cp, "w"), indent=1)
            return rec
        except Exception as e:
            last = e; time.sleep(2 * (attempt + 1))
    print(f"  risk: {ric} failed after retries: {last}")
    return None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("targets", help="JSON list of {ric,company,industry,country}")
    ap.add_argument("--workers", type=int, default=6)
    args = ap.parse_args()
    from concurrent.futures import ThreadPoolExecutor, as_completed
    tgt = json.load(open(args.targets))
    todo = [t for t in tgt if not os.path.exists(_cache_path(t["ric"]))]
    print(f"risk targets: {len(tgt)} | cached: {len(tgt)-len(todo)} | todo: {len(todo)}", flush=True)
    ok = fail = 0
    def work(t):
        return t["ric"], research_risk(t["ric"], t["company"], t.get("industry"), t.get("country"))
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futs = [ex.submit(work, t) for t in todo]
        for i, fut in enumerate(as_completed(futs), 1):
            try:
                _, rec = fut.result(); ok += rec is not None; fail += rec is None
            except Exception as e:
                fail += 1; print(f"  err: {e}")
            if i % 10 == 0 or i == len(todo):
                print(f"  {i}/{len(todo)} | ok={ok} fail={fail}", flush=True)
    print(f"DONE ok={ok} fail={fail}")


if __name__ == "__main__":
    main()
