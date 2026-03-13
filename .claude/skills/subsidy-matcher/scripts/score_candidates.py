#!/usr/bin/env python3
"""補助金候補の 6 軸採点を補助する簡易スクリプト。"""

import argparse
import json

FIELDS = [
    "business_fit",
    "eligibility_fit",
    "investment_fit",
    "requirement_feasibility",
    "schedule_feasibility",
    "selection_likelihood",
]


def bounded(value):
    return max(0, min(10, int(value)))


def summarize(candidate):
    scores = {field: bounded(candidate.get(field, 0)) for field in FIELDS}
    total = sum(scores.values())
    blockers = candidate.get("blockers", [])
    feasible = not any(blocker.get("critical") for blocker in blockers)
    return {
        "name": candidate.get("name", ""),
        "scores": scores,
        "total": total,
        "feasible": feasible,
        "blockers": blockers,
        "comment": candidate.get("comment", ""),
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="候補JSON")
    parser.add_argument("--output", help="出力JSON")
    args = parser.parse_args()

    with open(args.input, "r", encoding="utf-8") as f:
        payload = json.load(f)

    results = [summarize(candidate) for candidate in payload.get("candidates", [])]
    results.sort(key=lambda item: (item["feasible"], item["total"]), reverse=True)
    output = {"results": results}

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(output, f, ensure_ascii=False, indent=2)
    else:
        print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
