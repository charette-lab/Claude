#!/usr/bin/env python3
"""Research a list of companies concurrently and cache each record.

Resumable: any name already cached (niche_pipeline/.cache) is skipped, so a
re-run only fills the gaps. Used to complete the qualitative research over the
gate-1 survivors / universe gap (see README).

    python3 research_names.py to_research.json --workers 8

`to_research.json` is a list of {"ticker","company","industry"} objects.
Requires ANTHROPIC_API_KEY (analyst.py uses the web_search server tool).
"""
import argparse, json, os, sys, time
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import analyst
from concurrent.futures import ThreadPoolExecutor, as_completed


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("names_json", help="list of {ticker,company,industry}")
    ap.add_argument("--workers", type=int, default=8)
    args = ap.parse_args()

    names = json.load(open(args.names_json))
    todo = [n for n in names if not os.path.exists(analyst._cache_path(n["ticker"]))]
    print(f"total={len(names)} cached={len(names)-len(todo)} todo={len(todo)} workers={args.workers}")
    if not todo:
        print("nothing to do — all cached"); return

    ok = fail = 0
    def work(n):
        return n["ticker"], analyst.research(n["ticker"], n["company"], hint=n.get("industry") or "")
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futs = [ex.submit(work, n) for n in todo]
        for i, fut in enumerate(as_completed(futs), 1):
            try:
                _, rec = fut.result()
                ok += 1 if rec is not None else 0
                fail += 1 if rec is None else 0
            except Exception as e:
                fail += 1; print(f"  error: {e}")
            if i % 10 == 0 or i == len(todo):
                print(f"  {i}/{len(todo)} | ok={ok} fail={fail}", flush=True)
    print(f"DONE ok={ok} fail={fail}")


if __name__ == "__main__":
    main()
