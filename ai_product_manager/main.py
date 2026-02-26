#!/usr/bin/env python3
"""
AI Product Manager — Main Orchestrator
Ties together the scoring engine, survey analyzer, and output generators.

Usage:
    python main.py                          # Run full pipeline with defaults
    python main.py --mode score             # Score only (JSON output)
    python main.py --mode report            # Generate markdown report
    python main.py --mode deck              # Generate PowerPoint deck
    python main.py --mode excel             # Generate scored Excel workbook
    python main.py --mode full              # All outputs
    python main.py --config custom.json     # Use custom scoring config
"""
import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

# Add parent to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from engines.scoring_engine import run_scoring_pipeline
from generators.report_generator import ReportGenerator
from generators.excel_generator import generate_scored_workbook


def find_project_files():
    """Locate survey CSV and capability Excel files."""
    project_dir = Path(__file__).parent.parent / "mnt" / ".projects"
    survey_path = None
    xlsx_path = None

    for p in project_dir.rglob("*.csv"):
        survey_path = p
    for p in project_dir.rglob("*.xlsx"):
        xlsx_path = p

    return survey_path, xlsx_path


def prepare_deck_data(capabilities, survey_insights, summary_stats):
    """Transform scoring results into the format expected by the JS deck generator."""
    demand_data = []
    for cap, data in sorted(
        survey_insights.get("capability_demand", {}).items(),
        key=lambda x: -x[1]["demand_count"]
    ):
        demand_data.append({
            "capability": cap,
            "count": data["demand_count"],
            "pct": data["demand_pct"],
        })

    # Category breakdown
    categories = {}
    for c in capabilities:
        cat = c["sheet"]
        if cat not in categories:
            categories[cat] = {"scores": [], "high": 0}
        categories[cat]["scores"].append(c["scores"]["final"])
        if c["priority"] == "HIGH":
            categories[cat]["high"] += 1

    cat_names = {
        "main": "Core Capabilities",
        "pricing": "Dynamic Pricing",
        "custom_logic": "Custom Logic",
        "experiences": "Experiences",
    }

    cat_list = []
    for cat, info in categories.items():
        avg = sum(info["scores"]) / len(info["scores"]) if info["scores"] else 0
        cat_list.append({
            "name": cat_names.get(cat, cat),
            "avg_score": round(avg, 2),
            "count": len(info["scores"]),
            "high_count": info["high"],
        })

    top5 = []
    for c in capabilities[:5]:
        top5.append({
            "rank": c["rank"],
            "name": c["name"],
            "tier": c["tier"],
            "priority": c["priority"],
            "score": c["scores"]["final"],
            "customer_impact": c["scores"]["customer_impact"],
            "business_impact": c["scores"]["business_impact"],
            "cost_to_implement": c["scores"]["cost_to_implement"],
        })

    all_caps = []
    for c in capabilities:
        all_caps.append({
            "rank": c["rank"],
            "name": c["name"],
            "tier": c["tier"],
            "priority": c["priority"],
            "score": c["scores"]["final"],
            "customer_impact": c["scores"]["customer_impact"],
            "business_impact": c["scores"]["business_impact"],
        })

    return {
        "stats": summary_stats,
        "survey": {
            "total_respondents": survey_insights.get("total_respondents", 0),
            "companies": survey_insights.get("companies", []),
        },
        "top5": top5,
        "all_capabilities": all_caps,
        "demand": demand_data,
        "categories": cat_list,
    }


def generate_deck(deck_data, output_path):
    """Run the Node.js deck generator."""
    data_path = Path(output_path).parent / "deck_data.json"
    with open(data_path, "w") as f:
        json.dump(deck_data, f, indent=2)

    generator_path = Path(__file__).parent / "generators" / "deck_generator.js"
    result = subprocess.run(
        ["node", str(generator_path), str(data_path), str(output_path)],
        capture_output=True, text=True, timeout=30,
        cwd=str(Path(__file__).parent.parent)
    )
    if result.returncode != 0:
        print(f"Deck generation error: {result.stderr}")
        return False
    return True


def main():
    parser = argparse.ArgumentParser(description="AI Product Manager — Capability Prioritization Pipeline")
    parser.add_argument("--mode", choices=["score", "report", "deck", "excel", "full"], default="full",
                       help="Output mode")
    parser.add_argument("--survey", help="Path to survey CSV")
    parser.add_argument("--xlsx", help="Path to capabilities Excel file")
    parser.add_argument("--config", help="Path to custom scoring config JSON")
    parser.add_argument("--output", help="Output directory", default=None)
    args = parser.parse_args()

    # Find data files
    survey_path = args.survey
    xlsx_path = args.xlsx
    if not survey_path or not xlsx_path:
        auto_survey, auto_xlsx = find_project_files()
        survey_path = survey_path or str(auto_survey) if auto_survey else None
        xlsx_path = xlsx_path or str(auto_xlsx) if auto_xlsx else None

    if not survey_path or not xlsx_path:
        print("Error: Could not find survey CSV and/or capabilities XLSX files.")
        print("Use --survey and --xlsx to specify paths.")
        sys.exit(1)

    # Set output directory
    output_dir = Path(args.output) if args.output else Path(__file__).parent / "outputs"
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"📊 AI Product Manager — Running Pipeline")
    print(f"   Survey: {survey_path}")
    print(f"   Capabilities: {xlsx_path}")
    print(f"   Mode: {args.mode}")
    print(f"   Output: {output_dir}")
    print()

    # Run scoring pipeline
    print("🔄 Running scoring engine...")
    results = run_scoring_pipeline(survey_path, xlsx_path, args.config)
    capabilities = results["capabilities"]
    insights = results["survey_insights"]
    stats = results["summary_stats"]
    print(f"   ✅ Scored {stats['total_capabilities']} capabilities")
    print(f"   📈 {stats['high_priority']} HIGH | {stats['medium_priority']} MEDIUM | {stats['low_priority']} LOW")
    print()

    if args.mode in ("score", "full"):
        score_path = output_dir / "scoring_results.json"
        with open(score_path, "w") as f:
            json.dump({
                "summary": stats,
                "capabilities": capabilities,
                "survey_insights": {k: v for k, v in insights.items()},
            }, f, indent=2, default=str)
        print(f"📝 Scoring results saved: {score_path}")

    if args.mode in ("report", "full"):
        report_path = output_dir / "capability_report.md"
        reporter = ReportGenerator(capabilities, insights, stats)
        reporter.save(str(report_path))
        print(f"📄 Report saved: {report_path}")

    if args.mode in ("excel", "full"):
        excel_path = output_dir / "capability_rankings.xlsx"
        generate_scored_workbook(capabilities, insights, stats, str(excel_path))
        print(f"📊 Excel workbook saved: {excel_path}")

    if args.mode in ("deck", "full"):
        deck_path = output_dir / "executive_brief.pptx"
        deck_data = prepare_deck_data(capabilities, insights, stats)
        success = generate_deck(deck_data, str(deck_path))
        if success:
            print(f"🎯 Executive deck saved: {deck_path}")
        else:
            print(f"⚠️  Deck generation had issues — check logs")

    print()
    print("✅ Pipeline complete!")
    return results


if __name__ == "__main__":
    main()
