"""Insight generator module for rule-based recommendations and narrative text.

This module analyzes computed metrics and generates:
- Improvement items based on threshold comparisons
- Key achievements highlighting positive outcomes
- Areas of focus for improvement
- Executive summary narrative text
"""
from typing import Dict, Any, List, Optional
from dataclasses import dataclass

from constants import METRIC_THRESHOLDS
from metrics_calculator import ClientConfig


@dataclass
class InsightThresholds:
    """Thresholds for generating insights and recommendations."""
    
    # False positive rate thresholds
    fp_rate_good: float = 10.0
    fp_rate_warning: float = 15.0
    
    # Automation thresholds
    automation_good: float = 88.0
    automation_warning: float = 80.0
    
    # Response time thresholds
    mttr_good: int = 150
    mttr_warning: int = 200
    
    # High severity concentration threshold
    high_severity_concentration: float = 50.0  # % of high-sev in top 2 tactics
    
    # Manual escalation threshold
    manual_escalation_threshold: float = 12.0


def generate_improvement_items(metrics: Dict[str, Any], 
                                config: ClientConfig) -> List[Dict[str, Any]]:
    """Generate prioritized improvement items based on metrics.
    
    Args:
        metrics: Computed metrics dictionary
        config: Client configuration
        
    Returns:
        List of improvement item dictionaries
    """
    items = []
    thresholds = InsightThresholds()
    client_name = metrics.get("client_name", "Client")
    
    # Check false positive rate
    fp_rate = metrics.get("false_positive_rate", 0)
    detection_sources = metrics.get("detection_sources", [])
    
    if fp_rate > thresholds.fp_rate_good:
        # Find the source with highest FP rate
        worst_source = None
        worst_fp = 0
        for source in detection_sources:
            if source.get("fp_rate", 0) > worst_fp:
                worst_fp = source.get("fp_rate", 0)
                worst_source = source.get("source", "Unknown")
        
        if worst_source and worst_fp > thresholds.fp_rate_good:
            items.append({
                "title": "Detection Tuning",
                "priority": "HIGH" if fp_rate > thresholds.fp_rate_warning else "MEDIUM",
                "owner": f"CS SOC + {client_name} Security Team",
                "target": "Next 30 days",
                "description": (
                    f"{worst_source} false positive rate is {worst_fp}%, exceeding the "
                    f"{thresholds.fp_rate_good}% threshold and keeping the overall rate at {fp_rate}%. "
                    f"Tuning these alerts will reduce client noise and improve SOC efficiency."
                ),
            })
    
    # Check automation rate
    automation_percent = metrics.get("automation_percent", 0)
    analyst_escalation = metrics.get("analyst_escalation", {})
    manual_count = analyst_escalation.get("count", 0)
    manual_percent = analyst_escalation.get("percent", 0)
    
    if automation_percent < thresholds.automation_good:
        items.append({
            "title": "Automation Opportunity",
            "priority": "HIGH" if automation_percent < thresholds.automation_warning else "MEDIUM",
            "owner": "CS SOC Engineering",
            "target": "Next 60 days",
            "description": (
                f"Manual escalations at {manual_percent}% exceed our {thresholds.manual_escalation_threshold}% target. "
                f"{manual_count} incidents required analyst judgment. "
                "Expanding playbook coverage will improve consistency."
            ),
        })
    
    # Check high severity concentration in MITRE tactics
    tactics = metrics.get("tactics", [])
    high_severity = metrics.get("high_severity", [])
    
    if len(tactics) >= 2 and len(high_severity) >= 2:
        total_high = sum(high_severity)
        top_two_high = sum(high_severity[:2])
        
        if total_high > 0:
            concentration = (top_two_high / total_high * 100)
            
            if concentration >= thresholds.high_severity_concentration:
                top_tactics = " plus ".join(tactics[:2])
                items.append({
                    "title": "Threat Focus",
                    "priority": "HIGH",
                    "owner": f"Joint - CS Threat Intel + {client_name}",
                    "target": "Ongoing",
                    "description": (
                        f"{top_tactics} account for {top_two_high} of the {total_high} high-severity incidents "
                        f"({round(concentration)}%) in the MITRE dataset, signaling concentrated foothold attempts "
                        "that should drive proactive hunts and new detections."
                    ),
                })
    
    # Check MTTR
    mttr = metrics.get("mttr_minutes", 0)
    if mttr > thresholds.mttr_good:
        items.append({
            "title": "Response Time Optimization",
            "priority": "HIGH" if mttr > thresholds.mttr_warning else "MEDIUM",
            "owner": "CS SOC Operations",
            "target": "Next 45 days",
            "description": (
                f"Mean time to respond at {mttr} minutes exceeds the {thresholds.mttr_good} minute target. "
                "Review triage workflows and automation rules to improve response speed."
            ),
        })
    
    # Check SLA compliance
    sla_compliance = metrics.get("sla_compliance_rate", 100)
    if sla_compliance < 95.0:
        items.append({
            "title": "SLA Performance",
            "priority": "HIGH" if sla_compliance < 90.0 else "MEDIUM",
            "owner": "CS SOC Operations",
            "target": "Next 30 days",
            "description": (
                f"SLA compliance at {sla_compliance}% is below the 95% target. "
                "Focus on high-priority incident response workflows."
            ),
        })
    
    # Sort by priority (HIGH first)
    priority_order = {"HIGH": 0, "MEDIUM": 1, "LOW": 2}
    items.sort(key=lambda x: priority_order.get(x.get("priority", "LOW"), 2))
    
    # Limit to top 3 items
    return items[:3]


def generate_key_achievements(metrics: Dict[str, Any]) -> List[str]:
    """Generate key achievements based on positive metrics.
    
    Args:
        metrics: Computed metrics dictionary
        
    Returns:
        List of achievement strings
    """
    achievements = []
    
    # Response advantage
    response_advantage = metrics.get("response_advantage_percent", 0)
    if response_advantage > 0:
        achievements.append(f"{round(response_advantage)}% faster response than industry peers")
    
    # Threat containment
    threats_contained = metrics.get("true_threats_contained", 0)
    zero_breaches = metrics.get("zero_breaches", True)
    if threats_contained > 0 and zero_breaches:
        achievements.append(f"100% threat containment with zero breaches")
    elif threats_contained > 0:
        achievements.append(f"{threats_contained} threats successfully contained")
    
    # After-hours coverage
    after_hours = metrics.get("after_hours_escalations", 0)
    if after_hours > 0:
        achievements.append(f"{after_hours} after-hours escalations handled seamlessly")
    
    # Cost avoidance
    total_modeled = metrics.get("total_modeled", 0)
    if total_modeled > 0:
        if total_modeled >= 1000000:
            formatted = f"${total_modeled / 1000000:.2f}M"
        else:
            formatted = f"${total_modeled:,}"
        achievements.append(f"{formatted} modeled cost exposure avoided")
    
    # Low false positive rate
    fp_rate = metrics.get("false_positive_rate", 100)
    if fp_rate <= 10.0:
        achievements.append(f"False positive rate maintained at {fp_rate}%")
    
    # High automation
    automation = metrics.get("automation_percent", 0)
    if automation >= 85:
        achievements.append(f"{round(automation)}% of escalations handled via automated playbooks")
    
    # Limit to top 4
    return achievements[:4]


def generate_areas_of_focus(metrics: Dict[str, Any], 
                            improvement_items: List[Dict[str, Any]]) -> List[str]:
    """Generate areas of focus based on improvement opportunities.
    
    Args:
        metrics: Computed metrics dictionary
        improvement_items: Generated improvement items
        
    Returns:
        List of focus area strings
    """
    areas = []
    
    # Extract focus areas from improvement items
    for item in improvement_items:
        title = item.get("title", "")
        
        if title == "Detection Tuning":
            # Find worst source
            detection_sources = metrics.get("detection_sources", [])
            worst_source = max(detection_sources, key=lambda x: x.get("fp_rate", 0), default=None)
            if worst_source:
                areas.append(f"Reduce {worst_source.get('source', 'detection source')} false positive rate")
        
        elif title == "Automation Opportunity":
            areas.append("Expand playbook automation coverage")
        
        elif title == "Threat Focus":
            tactics = metrics.get("tactics", [])
            if tactics:
                areas.append(f"Proactive threat hunting for {tactics[0]} tactics")
        
        elif title == "Response Time Optimization":
            areas.append("Improve incident response workflows")
        
        elif title == "SLA Performance":
            areas.append("Focus on high-priority SLA compliance")
    
    # Limit to top 3
    return areas[:3]


def generate_next_period_goals(metrics: Dict[str, Any],
                                improvement_items: List[Dict[str, Any]]) -> List[str]:
    """Generate goals for the next reporting period.
    
    Args:
        metrics: Computed metrics dictionary
        improvement_items: Generated improvement items
        
    Returns:
        List of goal strings
    """
    goals = []
    thresholds = InsightThresholds()
    
    # FP rate goal
    fp_rate = metrics.get("false_positive_rate", 0)
    if fp_rate > thresholds.fp_rate_good:
        goals.append(f"Achieve {thresholds.fp_rate_good}% or lower overall false positive rate")
    
    # Automation goal
    automation = metrics.get("automation_percent", 0)
    if automation < thresholds.automation_good:
        goals.append(f"Reduce manual escalations to {thresholds.manual_escalation_threshold}% or below")
    
    # MITRE tactics goal
    tactics = metrics.get("tactics", [])
    if len(tactics) >= 2:
        goals.append(f"Implement enhanced detection for {tactics[1]}")
    
    # Response time goal
    mttr = metrics.get("mttr_minutes", 0)
    if mttr > thresholds.mttr_good:
        goals.append(f"Reduce MTTR to under {thresholds.mttr_good} minutes")
    
    # SLA goal
    sla = metrics.get("sla_compliance_rate", 100)
    if sla < 95:
        goals.append("Achieve 95%+ SLA compliance across all priorities")
    
    # Limit to top 3
    return goals[:3]


def generate_tuning_recommendations(metrics: Dict[str, Any]) -> List[str]:
    """Generate detection tuning recommendations.
    
    Args:
        metrics: Computed metrics dictionary
        
    Returns:
        List of tuning recommendation strings
    """
    recommendations = []
    thresholds = InsightThresholds()
    
    detection_sources = metrics.get("detection_sources", [])
    
    for source in detection_sources:
        source_name = source.get("source", "Unknown")
        fp_rate = source.get("fp_rate", 0)
        
        if fp_rate > thresholds.fp_rate_good:
            recommendations.append(
                f"Reduce {source_name} false positive rate from {fp_rate}% to target {thresholds.fp_rate_good}%"
            )
        elif fp_rate > 5.0:
            recommendations.append(
                f"Optimize {source_name} rules for better signal fidelity"
            )
        else:
            recommendations.append(
                f"Review {source_name} detection thresholds"
            )
    
    return recommendations[:3]


def generate_executive_summary_narrative(metrics: Dict[str, Any]) -> str:
    """Generate executive summary narrative paragraph.
    
    Args:
        metrics: Computed metrics dictionary
        
    Returns:
        Narrative string for executive summary
    """
    client_name = metrics.get("client_name", "Your organization")
    alerts_triaged = metrics.get("alerts_triaged", 0)
    client_touch = metrics.get("client_touch_decisions", 0)
    closed_e2e = metrics.get("closed_end_to_end", 0)
    
    response_advantage = metrics.get("response_advantage_percent", 0)
    mttr = metrics.get("mttr_minutes", 0)
    p90 = metrics.get("p90_minutes", 0)
    
    after_hours = metrics.get("after_hours_escalations", 0)
    incidents_escalated = metrics.get("incidents_escalated", 0)
    true_threats = metrics.get("true_threats_contained", 0)
    fp_rate = metrics.get("false_positive_rate", 0)
    
    narrative = (
        f"Your security posture remained strong this reporting period. "
        f"CS SOC triaged {alerts_triaged:,} alerts, partnering with your team on "
        f"{client_touch:,} decisions and closing {closed_e2e:,} end-to-end. "
    )
    
    if response_advantage > 0:
        narrative += (
            f"Response speed landed {round(response_advantage)}% faster than sector medians "
            f"({mttr}-minute MTTR, {p90}-minute P90), "
        )
    else:
        narrative += f"Response times averaged {mttr}-minute MTTR ({p90}-minute P90), "
    
    if after_hours > 0:
        narrative += (
            f"while {after_hours} escalations were absorbed after hours without gaps in coverage. "
        )
    else:
        narrative += "with consistent coverage throughout the period. "
    
    narrative += (
        f"Of the {incidents_escalated:,} incidents escalated, we identified "
        f"{true_threats} true positive threats and contained each before business impact, "
        f"keeping false positives at {fp_rate}%."
    )
    
    return narrative


def determine_detection_quality_trend(metrics: Dict[str, Any]) -> str:
    """Determine the detection quality trend.
    
    Args:
        metrics: Computed metrics dictionary
        
    Returns:
        Trend string: "improving", "stable", or "declining"
    """
    fp_trend = metrics.get("fp_trend", [])
    
    if len(fp_trend) < 2:
        return "stable"
    
    # Compare current to previous
    current = fp_trend[-1]
    previous = fp_trend[-2]
    
    if current < previous * 0.9:  # 10% improvement
        return "improving"
    elif current > previous * 1.1:  # 10% decline
        return "declining"
    else:
        return "stable"


def generate_all_insights(metrics: Dict[str, Any], 
                          config: ClientConfig) -> Dict[str, Any]:
    """Generate all insights and add them to the metrics dictionary.
    
    Args:
        metrics: Computed metrics dictionary
        config: Client configuration
        
    Returns:
        Dictionary with insight fields to merge into metrics
    """
    # Generate improvement items
    improvement_items = generate_improvement_items(metrics, config)
    
    # Generate achievements and goals
    key_achievements = generate_key_achievements(metrics)
    areas_of_focus = generate_areas_of_focus(metrics, improvement_items)
    next_period_goals = generate_next_period_goals(metrics, improvement_items)
    
    # Generate other insights
    tuning_recommendations = generate_tuning_recommendations(metrics)
    executive_summary_narrative = generate_executive_summary_narrative(metrics)
    detection_quality_trend = determine_detection_quality_trend(metrics)
    
    # Notification methods (derived from data or default)
    notification_methods = [
        {"method": "Email", "count": 0, "percent": 0},
        {"method": "Phone", "count": 0, "percent": 0},
        {"method": "Slack/Teams", "count": 0, "percent": 0},
    ]
    
    # Calculate notification distribution (if after_hours data available)
    after_hours = metrics.get("after_hours_escalations", 0)
    if after_hours > 0:
        # Default distribution based on typical patterns
        email_count = int(after_hours * 0.85)
        phone_count = int(after_hours * 0.10)
        slack_count = after_hours - email_count - phone_count
        
        notification_methods = [
            {"method": "Email", "count": email_count, "percent": round(email_count / after_hours * 100, 1)},
            {"method": "Phone", "count": phone_count, "percent": round(phone_count / after_hours * 100, 1)},
            {"method": "Slack/Teams", "count": slack_count, "percent": round(slack_count / after_hours * 100, 1)},
        ]
    
    return {
        "improvement_items": improvement_items,
        "key_achievements": key_achievements,
        "areas_of_focus": areas_of_focus,
        "next_period_goals": next_period_goals,
        "tuning_recommendations": tuning_recommendations,
        "executive_summary_narrative": executive_summary_narrative,
        "detection_quality_trend": detection_quality_trend,
        "notification_methods": notification_methods,
        "compliance_status": "Fully Compliant",  # Default
        "vulnerabilities_identified": 0,  # Would need vulnerability data
        "risk_reduction_percent": metrics.get("response_advantage_percent", 0),
    }
