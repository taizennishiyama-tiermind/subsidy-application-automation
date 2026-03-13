#!/usr/bin/env python3
"""
財務指標の簡易計算。

入力:
  {
    "company_name": "...",
    "analysis_date": "2026-03-13",
    "periods": [{...}, {...}, {...}]
  }
"""

import argparse
import json


def safe_div(numerator, denominator):
    if denominator in (0, None):
        return None
    return numerator / denominator


def growth(current, previous):
    if previous is None:
        return None
    return safe_div(current - previous, previous)


def evaluate(name, value):
    if value is None:
        return "データ不足"
    if name == "equity_ratio":
        return "良好" if value >= 0.3 else "要改善"
    if name == "current_ratio":
        return "良好" if value >= 1.2 else "要改善"
    if name == "operating_profit_margin":
        return "良好" if value >= 0.05 else "普通"
    if name == "debt_repayment_years":
        return "良好" if value <= 10 else "要改善"
    return "普通"


def calculate(period, previous):
    cash_flow = period.get("operating_profit", 0) + period.get("depreciation", 0)
    ratios = {
        "year": period["year"],
        "operating_profit_margin": safe_div(period.get("operating_profit", 0), period.get("sales", 0)),
        "roa": safe_div(period.get("net_profit", 0), period.get("total_assets", 0)),
        "roe": safe_div(period.get("net_profit", 0), period.get("net_assets", 0)),
        "equity_ratio": safe_div(period.get("net_assets", 0), period.get("total_assets", 0)),
        "current_ratio": safe_div(period.get("current_assets", 0), period.get("current_liabilities", 0)),
        "fixed_ratio": safe_div(period.get("fixed_assets", 0), period.get("net_assets", 0)),
        "sales_growth_rate": growth(period.get("sales", 0), None if previous is None else previous.get("sales", 0)),
        "profit_growth_rate": growth(period.get("operating_profit", 0), None if previous is None else previous.get("operating_profit", 0)),
        "debt_repayment_years": safe_div(period.get("interest_bearing_debt", 0), cash_flow),
        "interest_coverage_ratio": safe_div(period.get("operating_profit", 0), period.get("interest_expense", 0)),
        "value_added": period.get("operating_profit", 0) + period.get("personnel_cost", 0) + period.get("depreciation", 0),
    }
    ratios["evaluations"] = {key: evaluate(key, value) for key, value in ratios.items() if key not in {"year", "evaluations"}}
    return ratios


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    with open(args.input, "r", encoding="utf-8") as f:
        source = json.load(f)

    periods = source.get("periods", [])
    results = []
    for index, period in enumerate(periods):
        previous = periods[index - 1] if index > 0 else None
        results.append(calculate(period, previous))

    payload = {
        "company_name": source.get("company_name", ""),
        "analysis_date": source.get("analysis_date", ""),
        "ratios": results,
    }

    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print(f"OK: wrote {args.output}")


if __name__ == "__main__":
    main()
