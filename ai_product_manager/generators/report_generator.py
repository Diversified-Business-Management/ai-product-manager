"""
AI Product Manager - Report Generator
Generates markdown roadmap reports from scored capabilities and survey insights.
"""
from datetime import datetime


class ReportGenerator:
    def __init__(self, capabilities, survey_insights, summary_stats, config=None):
        self.capabilities = capabilities
        self.insights = survey_insights
        self.stats = summary_stats
        self.config = config or {}
        self.date = datetime.now().strftime("%B %d, %Y")

    def generate_full_report(self):
        sections = [
            self._header(),
            self._executive_summary(),
            self._survey_insights(),
            self._priority_matrix(),
            self._top_capabilities(),
            self._category_breakdown(),
            self._ai_recommendations(),
            self._appendix(),
        ]
        return "\n\n".join(sections)

    def _header(self):
        return f"""# AI Product Manager — Capability Prioritization Report

**Generated:** {self.date}
**Total Capabilities Scored:** {self.stats.get('total_capabilities', 0)}
**Survey Respondents:** {self.insights.get('total_respondents', 0)}

---"""

    def _executive_summary(self):
        high = self.stats.get("high_priority", 0)
        med = self.stats.get("medium_priority", 0)
        low = self.stats.get("low_priority", 0)
        avg = self.stats.get("avg_score", 0)
        top3 = self.capabilities[:3]
        top_names = ", ".join(c["name"] for c in top3)

        return f"""## Executive Summary

This report presents a weighted prioritization of {self.stats.get('total_capabilities', 0)} product capabilities, scored across three dimensions: Customer Impact (30%), Business Impact (50%), and Cost to Implement (20%). Scores are further adjusted using AI-driven signals from customer survey data.

**Priority Distribution:** {high} High, {med} Medium, {low} Low (avg score: {avg})

**Top 3 Capabilities:** {top_names}

The scoring model incorporates demand signals from {self.insights.get('total_respondents', 0)} enterprise customers including {', '.join(self.insights.get('companies', [])[:5])} and others."""

    def _survey_insights(self):
        lines = ["## Customer Survey Insights", ""]

        # Business models
        models = self.insights.get("business_models", {})
        if models:
            lines.append("### Business Models Represented")
            lines.append("")
            for model, count in sorted(models.items(), key=lambda x: -x[1]):
                lines.append(f"- **{model}**: {count} respondents")
            lines.append("")

        # Existing platforms
        platforms = self.insights.get("existing_platforms", {})
        if platforms:
            lines.append("### Existing Commerce Platforms")
            lines.append("")
            for plat, count in sorted(platforms.items(), key=lambda x: -x[1]):
                lines.append(f"- **{plat}**: {count} respondents")
            lines.append("")

        # Capability demand
        demand = self.insights.get("capability_demand", {})
        if demand:
            lines.append("### Capability Demand Heatmap")
            lines.append("")
            lines.append("| Capability | Demand Count | % of Respondents | Requesting Companies |")
            lines.append("|------------|-------------|-------------------|---------------------|")
            for cap, data in sorted(demand.items(), key=lambda x: -x[1]["demand_count"]):
                companies = ", ".join(data["requesting_companies"][:4])
                if len(data["requesting_companies"]) > 4:
                    companies += f" (+{len(data['requesting_companies']) - 4} more)"
                lines.append(f"| {cap} | {data['demand_count']} | {data['demand_pct']}% | {companies} |")

        return "\n".join(lines)

    def _priority_matrix(self):
        lines = ["## Priority Matrix", ""]
        lines.append("| Rank | Capability | Tier | Priority | Final Score | Customer | Business | Cost |")
        lines.append("|------|-----------|------|----------|-------------|----------|----------|------|")
        for c in self.capabilities:
            s = c["scores"]
            lines.append(
                f"| {c['rank']} | {c['name']} | {c['tier']} | **{c['priority']}** | "
                f"{s['final']:.2f} | {s['customer_impact']:.2f} | "
                f"{s['business_impact']:.2f} | {s['cost_to_implement']:.2f} |"
            )
        return "\n".join(lines)

    def _top_capabilities(self):
        lines = ["## Top 10 Capabilities — Deep Dive", ""]
        for c in self.capabilities[:10]:
            s = c["scores"]
            lines.append(f"### {c['rank']}. {c['name']}")
            lines.append(f"**Tier:** {c['tier']} | **Priority:** {c['priority']} | **Score:** {s['final']:.2f}")
            lines.append("")
            lines.append(f"- Customer Impact: {s['customer_impact']:.2f}")
            lines.append(f"- Business Impact: {s['business_impact']:.2f}")
            lines.append(f"- Cost to Implement: {s['cost_to_implement']:.2f}")
            if c.get("survey_demand", 0) > 0:
                lines.append(f"- Survey Demand: {c['survey_demand']} customers requesting this")
            if c.get("matched_survey_capability"):
                lines.append(f"- Matched Survey Category: {c['matched_survey_capability']}")
            if c.get("development_quarter"):
                lines.append(f"- Target Quarter: {c['development_quarter']}")
            if c.get("notes"):
                lines.append(f"- Notes: {c['notes']}")
            lines.append("")
        return "\n".join(lines)

    def _category_breakdown(self):
        categories = {}
        for c in self.capabilities:
            cat = c["sheet"]
            if cat not in categories:
                categories[cat] = []
            categories[cat].append(c)

        cat_names = {
            "main": "Core Capability Scoring",
            "pricing": "Dynamic Pricing",
            "custom_logic": "Custom Logic (Stack Ranked)",
            "experiences": "Experience Capabilities",
        }

        lines = ["## Category Breakdown", ""]
        for cat, caps in categories.items():
            display_name = cat_names.get(cat, cat)
            scores = [c["scores"]["final"] for c in caps]
            avg = sum(scores) / len(scores) if scores else 0
            high = len([c for c in caps if c["priority"] == "HIGH"])
            lines.append(f"### {display_name}")
            lines.append(f"**Capabilities:** {len(caps)} | **Avg Score:** {avg:.2f} | **High Priority:** {high}")
            lines.append("")
            for c in caps:
                lines.append(f"- [{c['priority']}] {c['name']} — {c['scores']['final']:.2f}")
            lines.append("")
        return "\n".join(lines)

    def _ai_recommendations(self):
        high = [c for c in self.capabilities if c["priority"] == "HIGH"]
        high_demand = sorted(
            [c for c in self.capabilities if c.get("survey_demand", 0) >= 3],
            key=lambda x: -x["survey_demand"]
        )

        lines = ["## AI-Powered Recommendations", ""]
        lines.append("### Immediate Action Items")
        lines.append("")
        if high:
            lines.append(f"1. **Focus on {high[0]['name']}** (Score: {high[0]['scores']['final']:.2f}) — highest overall priority combining business impact and customer demand.")
        if len(high) > 1:
            lines.append(f"2. **Accelerate {high[1]['name']}** (Score: {high[1]['scores']['final']:.2f}) — strong business case with high strategic value.")
        if len(high) > 2:
            lines.append(f"3. **Invest in {high[2]['name']}** (Score: {high[2]['scores']['final']:.2f}) — significant customer pain point with clear ROI.")
        lines.append("")

        if high_demand:
            lines.append("### Customer-Demand Driven Priorities")
            lines.append("")
            for c in high_demand[:5]:
                lines.append(f"- **{c['name']}**: {c['survey_demand']} customers requesting — {c.get('matched_survey_capability', 'general demand')}")
            lines.append("")

        lines.append("### Strategic Observations")
        lines.append("")
        lines.append("- Capabilities with high business impact but lower customer visibility may need better positioning in sales conversations.")
        lines.append("- Survey data shows strong demand for catalog and pricing capabilities — align roadmap messaging accordingly.")
        lines.append("- Consider bundling related high-priority items (e.g., Catalog + Merchandising) into a single release for maximum market impact.")

        return "\n".join(lines)

    def _appendix(self):
        return f"""## Appendix

### Scoring Methodology

Capabilities are scored on a 1-10 scale across three weighted dimensions:

- **Customer Impact (30%):** Breadth of customers impacted (50%), Severity of pain (25%), Level of pain mitigation (25%)
- **Business Impact (50%):** New ACV (30%), Sales & Retention (30%), Strategic Value (30%), Competitive USP (10%)
- **Cost to Implement (20%):** Engineering Cost (60%), Ability to Execute (25%), COGS Impact (15%)

### AI Adjustments Applied

- Survey demand boost: +0.5 (3+ customers), +1.0 (5+), +1.5 (7+)
- Competitive gap penalty: -0.3 when competitors lead
- Time decay: -0.1 per quarter for capabilities planned further out
- Tier multipliers: Core (1.0x), Commerce 1 (1.1x), Commerce 2 (0.9x), Ext Studio (1.05x)

---
*Report generated by AI Product Manager v1.0*"""

    def save(self, path):
        report = self.generate_full_report()
        with open(path, "w") as f:
            f.write(report)
        return path
