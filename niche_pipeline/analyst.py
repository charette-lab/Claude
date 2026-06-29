"""
analyst.py — the qualitative research backend.

The deterministic engines (frameworks.py, aip.py) need inputs that only exist
after real research: the ten 1-10 moat chapter scores, the identity of the
durable core, the ten core-level chapter scores, the shareholder register, and
the eleven binary risk tags. This module produces them for one company with a
single structured Claude call that uses server-side web search, then returns a
strict JSON record. Results are cached to disk so a re-run is free.

Backend: the Anthropic Messages API with the `web_search` server tool. Set
ANTHROPIC_API_KEY. If no key is configured the pipeline still runs every
deterministic module and simply leaves the researched fields blank (flagged
NEEDS_RESEARCH), so the program degrades gracefully instead of failing.
"""
from __future__ import annotations
import hashlib
import json
import os
import re
import time

from frameworks import (NCC_CHAPTERS, RISK_TAGS, RISK_TAG_DEFS)

MODEL = os.environ.get("NCP_MODEL", "claude-opus-4-8")
CACHE_DIR = os.environ.get("NCP_CACHE", os.path.join(os.path.dirname(__file__), ".cache"))


def _have_sdk_and_key():
    if not os.environ.get("ANTHROPIC_API_KEY"):
        return False
    try:
        import anthropic  # noqa: F401
        return True
    except Exception:
        return False


SYSTEM = """You are a buy-side analyst applying three proprietary frameworks to ONE company.
Use web search to gather real, current evidence (segment disclosures, market-share data,
the latest shareholder register / yuho, patent timelines, regulatory exposure). Be ruthless
and evidence-based; do not flatter the company. Return ONLY one JSON object, no prose.

SCORING RUBRICS

A) Unified Niche Compounder v3.2 — score each chapter 1-10 (10 = strongest moat),
   blending internal vs external evidence per the chapter's Domain (A 80/20, B 20/80, C 10/90):
   1 Criticality (cost of failure; vitamin->painkiller->life-support)
   2 Standard Premium (license to operate; generic->trusted->gatekeeper)
   3 Market Hegemony (big fish/small pond; fragmented->oligopoly->niche monopoly TAM<$1B share>50%)
   4 Ecosystem Trap (defensive adoption; standalone->convenient->mandatory network)
   5 Lifecycle Advantage (structural cost leadership; commodity->quality->structural)
   6 Substitution Threat (1=a new method is 10x better, 10=physics permits no alternative)
   7 Demand Horizon (1=fad, 10=Lindy/permanent basic need)
   8 Capital Allocation (empire-builder->steward->outsider)
   9 Reinvestment Scorecard (saturated->scaling->unlimited high-ROI runway)
   10 Reverse-Lindy (1=fragile/AI-gated, 10=antifragile at physical efficiency limit)
   Score the WHOLE company once, then identify the durable CORE business and score the core once.

B) Engaged Ownership v3.0 — for the identified core, also give a Detachability triple (1-10, high=cleaner exit):
   entanglement (Ch9: high=turnkey, owns value chain),
   capital_independence (Ch8 capital-drain: high=cash engine not a sink),
   core_independence (Ch11 core-coupling: high=own customers, selling rest doesn't hurt core).

C) Ownership block test (Japan & general) — find the real beneficial owners (ignore trust-bank
   nominee lines like Master Trust / Custody Bank). Give the largest single holder/cohesive bloc %
   (founder+family+holding-co, or a parent). Special resolutions need 2/3 of votes, so >1/3 can veto.

D) 11 binary risk tags — 1 = material exposure, 0 = none. Definitions:
%RISKDEFS%

Return EXACTLY this JSON shape:
{
 "ticker": str, "company": str,
 "core_business": str,                      // short identification of the durable core
 "chapters_company": [10 ints],             // Ch1..Ch10 whole company
 "chapters_core": [10 ints],                // Ch1..Ch10 core only
 "core_detach": {"entanglement": int, "capital_independence": int, "core_independence": int},
 "shareholders": [{"holder": str, "pct": float, "type": "real|nominee"}],
 "largest_bloc_pct": float,                 // largest REAL holder/bloc, % of votes
 "ownership_notes": str,
 "risk_tags": [11 ints],                    // order exactly as the definitions above
 "risk_notes": str,
 "sources": [str]
}"""


def _prompt_system():
    defs = "\n".join(f"   {i+1}. {t}: {RISK_TAG_DEFS[t]}" for i, t in enumerate(RISK_TAGS))
    return SYSTEM.replace("%RISKDEFS%", defs)


def _cache_path(ticker):
    os.makedirs(CACHE_DIR, exist_ok=True)
    h = hashlib.md5((ticker or "").encode()).hexdigest()[:10]
    return os.path.join(CACHE_DIR, f"{h}.json")


def _extract_json(text):
    m = re.search(r"\{.*\}", text, re.DOTALL)
    if not m:
        raise ValueError("no JSON in model response")
    return json.loads(m.group(0))


def research(ticker, company, hint=None, max_searches=8, retries=2):
    """Return the analyst JSON record for one company (cached)."""
    cp = _cache_path(ticker)
    if os.path.exists(cp):
        with open(cp) as f:
            return json.load(f)
    if not _have_sdk_and_key():
        return None
    import anthropic
    client = anthropic.Anthropic()
    user = f"Company: {company}\nTicker: {ticker}"
    if hint:
        user += f"\nContext hint (verify, do not trust blindly): {hint}"
    last = None
    for attempt in range(retries + 1):
        try:
            resp = client.messages.create(
                model=MODEL, max_tokens=4000, system=_prompt_system(),
                tools=[{"type": "web_search_20250305", "name": "web_search",
                        "max_uses": max_searches}],
                messages=[{"role": "user", "content": user}],
            )
            text = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
            rec = _extract_json(text)
            rec.setdefault("ticker", ticker)
            rec.setdefault("company", company)
            with open(cp, "w") as f:
                json.dump(rec, f, indent=2)
            return rec
        except Exception as e:  # transient API / parse error -> backoff
            last = e
            time.sleep(2 * (attempt + 1))
    print(f"  analyst: {ticker} failed after retries: {last}")
    return None
