#!/usr/bin/env python3
"""案件情報から検索クエリ候補を作る。"""

import argparse
import json


def compact(parts):
    return " ".join([part for part in parts if part]).strip()


def build_queries(case):
    industry = case.get("industry", "")
    challenge = case.get("challenge", "")
    location = case.get("location", "")
    investment = case.get("investment_type", "")
    year = str(case.get("year", "2026"))

    queries = []
    queries.append(compact(["中小企業", "補助金", industry, challenge, year, "公募"]))
    queries.append(compact([location, "補助金", industry, challenge, year]))
    queries.append(compact([investment, "補助金", industry, year]))
    queries.append(compact([industry, challenge, "補助金", "申請受付中", year]))
    return [query for query in queries if query]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="案件情報JSON")
    parser.add_argument("--output", help="出力JSON")
    args = parser.parse_args()

    with open(args.input, "r", encoding="utf-8") as f:
        case = json.load(f)

    payload = {"queries": build_queries(case)}

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    else:
        print(json.dumps(payload, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
