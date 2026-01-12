"""Metrics calculator module for computing aggregated metrics from incident data.

This module takes parsed incident data and computes all the aggregated metrics
needed to populate the ReportData dataclass.
"""
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional, Tuple
from collections import Counter, defaultdict
import re

from data_parser import Incident


@dataclass
class ClientConfig:
    """Configuration for client-specific and benchmark data."""
    
    # Client tier (e.g., "Signature Tier", "Standard Tier")
    tier: str = "Standard Tier"
    
    # Industry benchmark values
    industry_mttr_minutes: int = 192
    industry_mttd_minutes: int = 66
    industry_incidents_per_day: float = 11.4
    
    # Report period (optional - can be derived from data)
    period_start: Optional[str] = None
    period_end: Optional[str] = None
    
    # Client name override (if None, derived from Organization column)
    client_name_override: Optional[str] = None
    
    # Cost modeling parameters
    analyst_hourly_rate: int = 85  # $/hour for analyst time
    coverage_hourly_rate: int = 220  # $/hour for 24x7 coverage
    breach_cost_estimate: int = 4200000  # Average breach cost
    
    # SLA targets by priority (in minutes)
    sla_targets: Dict[str, int] = field(default_factory=lambda: {
        "Critical": 30,
        "High": 60,
        "Medium": 180,
        "Low": 240,
    })
    
    # Business hours definition (for after-hours calculation)
    business_hours_start: int = 8  # 8 AM
    business_hours_end: int = 18  # 6 PM
    
    # Thresholds for recommendations
    fp_rate_threshold: float = 10.0  # FP rate threshold for tuning recommendations
    automation_target: float = 88.0  # Target automation percentage


def parse_priority_level(priority_str: Optional[str]) -> Optional[str]:
    """Extract priority level from priority string.
    
    Args:
        priority_str: String like "3 - HIGH" or "Critical"
        
    Returns:
        Normalized priority level: "Critical", "High", "Medium", "Low", or "Informational"
    """
    if not priority_str:
        return None
    
    priority_str = priority_str.upper()
    
    if "CRITICAL" in priority_str or "1 -" in priority_str:
        return "Critical"
    elif "HIGH" in priority_str or "2 -" in priority_str or "3 -" in priority_str:
        return "High"
    elif "MEDIUM" in priority_str or "4 -" in priority_str or "5 -" in priority_str:
        return "Medium"
    elif "LOW" in priority_str or "6 -" in priority_str or "7 -" in priority_str:
        return "Low"
    elif "INFO" in priority_str or "8 -" in priority_str:
        return "Informational"
    
    return None


def parse_vendor_severity(severity_str: Optional[str]) -> Optional[str]:
    """Normalize vendor severity string.
    
    Args:
        severity_str: Raw vendor severity string
        
    Returns:
        Normalized severity: "Critical", "High", "Medium", "Low", or "Informational"
    """
    if not severity_str:
        return None
    
    severity_str = severity_str.upper()
    
    if "CRITICAL" in severity_str:
        return "Critical"
    elif "HIGH" in severity_str:
        return "High"
    elif "MEDIUM" in severity_str or "MED" in severity_str:
        return "Medium"
    elif "LOW" in severity_str:
        return "Low"
    elif "INFO" in severity_str:
        return "Informational"
    
    return None


def is_after_hours(dt: Optional[datetime], config: ClientConfig) -> bool:
    """Determine if a datetime falls outside business hours.
    
    Args:
        dt: Datetime to check
        config: Client configuration with business hours
        
    Returns:
        True if after hours (evening/night or weekend)
    """
    if not dt:
        return False
    
    # Weekend check
    if dt.weekday() >= 5:  # Saturday = 5, Sunday = 6
        return True
    
    # After hours check (before 8 AM or after 6 PM)
    if dt.hour < config.business_hours_start or dt.hour >= config.business_hours_end:
        return True
    
    return False


def is_weekend(dt: Optional[datetime]) -> bool:
    """Check if datetime falls on a weekend."""
    if not dt:
        return False
    return dt.weekday() >= 5


def timedelta_to_minutes(td: Optional[timedelta]) -> int:
    """Convert timedelta to total minutes."""
    if not td:
        return 0
    return int(td.total_seconds() / 60)


def calculate_percentile(values: List[int], percentile: int) -> int:
    """Calculate the nth percentile of a list of values.
    
    Args:
        values: List of numeric values
        percentile: Percentile to calculate (0-100)
        
    Returns:
        The percentile value
    """
    if not values:
        return 0
    
    sorted_values = sorted(values)
    index = int(len(sorted_values) * percentile / 100)
    index = min(index, len(sorted_values) - 1)
    return sorted_values[index]


def calculate_period_metrics(incidents: List[Incident], config: ClientConfig) -> Dict[str, Any]:
    """Calculate metrics for a single period.
    
    Args:
        incidents: List of incidents for this period
        config: Client configuration
        
    Returns:
        Dictionary of computed metrics
    """
    if not incidents:
        return {
            "incidents_escalated": 0,
            "mttr_minutes": 0,
            "mttd_minutes": 0,
            "p90_minutes": 0,
            "false_positive_rate": 0.0,
        }
    
    # Count incidents
    total_incidents = len(incidents)
    
    # Calculate response times
    ttr_values = []
    ttd_values = []
    for inc in incidents:
        if inc.cs_soc_ttr:
            ttr_values.append(timedelta_to_minutes(inc.cs_soc_ttr))
        if inc.cs_soc_ttd:
            ttd_values.append(timedelta_to_minutes(inc.cs_soc_ttd))
    
    mttr = int(sum(ttr_values) / len(ttr_values)) if ttr_values else 0
    mttd = int(sum(ttd_values) / len(ttd_values)) if ttd_values else 0
    p90 = calculate_percentile(ttr_values, 90) if ttr_values else 0
    
    # Calculate false positive rate
    fp_count = sum(1 for inc in incidents 
                   if inc.cs_soc_verdict and "FALSE" in inc.cs_soc_verdict.upper())
    fp_rate = (fp_count / total_incidents * 100) if total_incidents > 0 else 0.0
    
    return {
        "incidents_escalated": total_incidents,
        "mttr_minutes": mttr,
        "mttd_minutes": mttd,
        "p90_minutes": p90,
        "false_positive_rate": round(fp_rate, 1),
    }


def calculate_volume_metrics(incidents: List[Incident], config: ClientConfig) -> Dict[str, Any]:
    """Calculate volume-related metrics from incidents.
    
    Args:
        incidents: List of incidents
        config: Client configuration
        
    Returns:
        Dictionary of volume metrics
    """
    total = len(incidents)
    
    # Closed status variations
    closed_count = sum(1 for inc in incidents 
                       if inc.current_status and "CLOSED" in inc.current_status.upper())
    
    # True positives (threats contained)
    true_positives = sum(1 for inc in incidents 
                         if inc.cs_soc_verdict and "TRUE POSITIVE" in inc.cs_soc_verdict.upper())
    
    # Escalation methods
    playbook_auto_count = sum(1 for inc in incidents 
                              if inc.initial_escalation_method and 
                              inc.initial_escalation_method.upper() != "CS SOC")
    analyst_count = sum(1 for inc in incidents 
                        if inc.initial_escalation_method and 
                        inc.initial_escalation_method.upper() == "CS SOC")
    
    playbook_percent = (playbook_auto_count / total * 100) if total > 0 else 0
    analyst_percent = (analyst_count / total * 100) if total > 0 else 0
    
    return {
        "alerts_triaged": total,
        "incidents_escalated": total,
        "closed_end_to_end": closed_count,
        "true_threats_contained": true_positives,
        "playbook_auto": {
            "count": playbook_auto_count,
            "percent": round(playbook_percent),
        },
        "analyst_escalation": {
            "count": analyst_count,
            "percent": round(analyst_percent),
        },
        "automation_percent": round(playbook_percent, 1),
    }


def calculate_response_metrics(incidents: List[Incident], config: ClientConfig) -> Dict[str, Any]:
    """Calculate response time metrics.
    
    Args:
        incidents: List of incidents
        config: Client configuration
        
    Returns:
        Dictionary of response metrics
    """
    # All TTR values
    all_ttr = []
    all_ttd = []
    
    # TTR by priority
    critical_high_ttr = []
    medium_low_ttr = []
    
    for inc in incidents:
        priority = parse_priority_level(inc.current_priority)
        
        if inc.cs_soc_ttr:
            minutes = timedelta_to_minutes(inc.cs_soc_ttr)
            all_ttr.append(minutes)
            
            if priority in ("Critical", "High"):
                critical_high_ttr.append(minutes)
            elif priority in ("Medium", "Low"):
                medium_low_ttr.append(minutes)
        
        if inc.cs_soc_ttd:
            all_ttd.append(timedelta_to_minutes(inc.cs_soc_ttd))
    
    mttr = int(sum(all_ttr) / len(all_ttr)) if all_ttr else 0
    mttd = int(sum(all_ttd) / len(all_ttd)) if all_ttd else 0
    p90 = calculate_percentile(all_ttr, 90) if all_ttr else 0
    
    crit_high_mttr = int(sum(critical_high_ttr) / len(critical_high_ttr)) if critical_high_ttr else 0
    med_low_mttr = int(sum(medium_low_ttr) / len(medium_low_ttr)) if medium_low_ttr else 0
    
    # Response advantage vs industry
    industry_mttr = config.industry_mttr_minutes
    advantage = ((industry_mttr - mttr) / industry_mttr * 100) if industry_mttr > 0 else 0
    
    # SLA compliance
    met_sla = 0
    total_with_ttr = 0
    for inc in incidents:
        if not inc.cs_soc_ttr:
            continue
        
        priority = parse_priority_level(inc.current_priority)
        if not priority:
            continue
        
        target = config.sla_targets.get(priority, 180)
        minutes = timedelta_to_minutes(inc.cs_soc_ttr)
        total_with_ttr += 1
        if minutes <= target:
            met_sla += 1
    
    sla_compliance = (met_sla / total_with_ttr * 100) if total_with_ttr > 0 else 0
    
    return {
        "mttr_minutes": mttr,
        "mttd_minutes": mttd,
        "p90_minutes": p90,
        "critical_high_mttr": crit_high_mttr,
        "medium_low_mttr": med_low_mttr,
        "response_advantage_percent": round(advantage, 1),
        "sla_compliance_rate": round(sla_compliance, 1),
        "avg_response_time": mttr,
        "fastest_response_time": min(all_ttr) if all_ttr else 0,
        "industry_median_minutes": industry_mttr,
    }


def calculate_detection_sources(incidents: List[Incident]) -> List[Dict[str, Any]]:
    """Calculate detection source breakdown.
    
    Args:
        incidents: List of incidents
        
    Returns:
        List of detection source dictionaries
    """
    # Group by product
    product_counts = Counter()
    product_fps = Counter()
    
    for inc in incidents:
        product = inc.product or "Unknown"
        product_counts[product] += 1
        
        if inc.cs_soc_verdict and "FALSE" in inc.cs_soc_verdict.upper():
            product_fps[product] += 1
    
    total = len(incidents)
    sources = []
    
    for product, count in product_counts.most_common():
        fp_count = product_fps.get(product, 0)
        fp_rate = (fp_count / count * 100) if count > 0 else 0
        percent = (count / total * 100) if total > 0 else 0
        
        sources.append({
            "source": product,
            "incidents": count,
            "percent": round(percent, 1),
            "fp_rate": round(fp_rate, 1),
        })
    
    return sources


def calculate_mitre_data(incidents: List[Incident]) -> Dict[str, Any]:
    """Calculate MITRE ATT&CK tactic data.
    
    Args:
        incidents: List of incidents
        
    Returns:
        Dictionary with tactics and severity breakdowns
    """
    # Count by tactic and priority
    tactic_priority_counts = defaultdict(lambda: defaultdict(int))
    
    for inc in incidents:
        tactic = inc.mitre_tactic_name
        if not tactic:
            tactic = "Unknown"
        
        priority = parse_priority_level(inc.current_priority) or "Unknown"
        tactic_priority_counts[tactic][priority] += 1
    
    # Get top tactics by total count
    tactic_totals = {
        tactic: sum(priorities.values())
        for tactic, priorities in tactic_priority_counts.items()
    }
    
    top_tactics = sorted(tactic_totals.keys(), key=lambda t: tactic_totals[t], reverse=True)[:5]
    
    # Build severity lists
    tactics = []
    high_severity = []
    medium_severity = []
    low_severity = []
    info_severity = []
    
    for tactic in top_tactics:
        tactics.append(tactic)
        high_severity.append(
            tactic_priority_counts[tactic].get("Critical", 0) + 
            tactic_priority_counts[tactic].get("High", 0)
        )
        medium_severity.append(tactic_priority_counts[tactic].get("Medium", 0))
        low_severity.append(tactic_priority_counts[tactic].get("Low", 0))
        info_severity.append(tactic_priority_counts[tactic].get("Informational", 0))
    
    return {
        "tactics": tactics,
        "high_severity": high_severity,
        "medium_severity": medium_severity,
        "low_severity": low_severity,
        "info_severity": info_severity,
    }


def calculate_severity_flows(incidents: List[Incident]) -> List[Dict[str, Any]]:
    """Calculate severity flow data for Sankey diagram.
    
    Maps vendor severity to CS priority.
    
    Args:
        incidents: List of incidents
        
    Returns:
        List of flow dictionaries
    """
    flow_counts = defaultdict(int)
    
    for inc in incidents:
        vendor_sev = parse_vendor_severity(inc.vendor_severity)
        cs_priority = parse_priority_level(inc.current_priority)
        
        if not vendor_sev or not cs_priority:
            continue
        
        key = (f"Vendor {vendor_sev}", f"CS {cs_priority}")
        flow_counts[key] += 1
    
    flows = []
    for (from_node, to_node), count in flow_counts.items():
        if count > 0:
            flows.append({
                "from": from_node,
                "to": to_node,
                "flow": count,
            })
    
    # Sort by flow count descending
    flows.sort(key=lambda x: x["flow"], reverse=True)
    
    return flows


def calculate_after_hours_metrics(incidents: List[Incident], config: ClientConfig) -> Dict[str, Any]:
    """Calculate after-hours coverage metrics.
    
    Args:
        incidents: List of incidents
        config: Client configuration
        
    Returns:
        Dictionary of after-hours metrics
    """
    total = len(incidents)
    
    after_hours_incidents = []
    business_hours_count = 0
    weekend_count = 0
    
    for inc in incidents:
        dt = inc.escalated_datetime_utc or inc.created_datetime_utc
        
        if is_weekend(dt):
            weekend_count += 1
            after_hours_incidents.append(inc)
        elif is_after_hours(dt, config):
            after_hours_incidents.append(inc)
        else:
            business_hours_count += 1
    
    after_hours_count = len(after_hours_incidents)
    weeknight_count = after_hours_count - weekend_count
    
    # After-hours by priority
    ah_by_priority = Counter()
    for inc in after_hours_incidents:
        priority = parse_priority_level(inc.current_priority) or "Unknown"
        ah_by_priority[priority] += 1
    
    return {
        "after_hours_escalations": after_hours_count,
        "after_hours_weeknight": weeknight_count,
        "after_hours_weekend": weekend_count,
        "after_hours_critical": ah_by_priority.get("Critical", 0),
        "after_hours_high": ah_by_priority.get("High", 0),
        "after_hours_medium": ah_by_priority.get("Medium", 0),
        "after_hours_low": ah_by_priority.get("Low", 0),
        "business_hours_percent": round((business_hours_count / total * 100) if total > 0 else 0, 1),
        "after_hours_percent": round((weeknight_count / total * 100) if total > 0 else 0, 1),
        "weekend_percent": round((weekend_count / total * 100) if total > 0 else 0, 1),
    }


def calculate_detection_quality(incidents: List[Incident]) -> Dict[str, Any]:
    """Calculate detection quality metrics.
    
    Args:
        incidents: List of incidents
        
    Returns:
        Dictionary of detection quality metrics
    """
    total = len(incidents)
    if total == 0:
        return {
            "true_threat_precision": 0.0,
            "signal_fidelity": 0.0,
            "client_validated": 0.0,
            "false_positive_rate": 0.0,
            "containment_rate": 0.0,
            "signal_to_noise_ratio": 0.0,
        }
    
    # Count by verdict
    true_positives = sum(1 for inc in incidents 
                         if inc.cs_soc_verdict and "TRUE POSITIVE" in inc.cs_soc_verdict.upper())
    false_positives = sum(1 for inc in incidents 
                          if inc.cs_soc_verdict and "FALSE" in inc.cs_soc_verdict.upper())
    
    # Contained threats (with response action)
    contained = sum(1 for inc in incidents 
                    if inc.response_action and inc.response_action_status and 
                    "SUCCESS" in inc.response_action_status.upper())
    
    # Client validated (closed by client)
    client_closed = sum(1 for inc in incidents 
                        if inc.closed_by and "CRITICALSTART" not in inc.closed_by.upper())
    
    true_threat_precision = (true_positives / total * 100) if total > 0 else 0
    fp_rate = (false_positives / total * 100) if total > 0 else 0
    signal_fidelity = 100 - fp_rate
    client_validated = (client_closed / total * 100) if total > 0 else 0
    containment_rate = (contained / true_positives * 100) if true_positives > 0 else 100
    
    # Signal to noise ratio (true positives / false positives)
    signal_to_noise = (true_positives / false_positives) if false_positives > 0 else float(true_positives)
    
    return {
        "true_threat_precision": round(true_threat_precision, 1),
        "signal_fidelity": round(signal_fidelity, 1),
        "client_validated": round(client_validated, 1),
        "false_positive_rate": round(fp_rate, 1),
        "containment_rate": round(containment_rate, 1),
        "signal_to_noise_ratio": round(signal_to_noise, 1),
    }


def calculate_collaboration_metrics(incidents: List[Incident]) -> Dict[str, Any]:
    """Calculate collaboration metrics.
    
    Args:
        incidents: List of incidents
        
    Returns:
        Dictionary of collaboration metrics
    """
    total = len(incidents)
    if total == 0:
        return {
            "avg_touches": 0.0,
            "client_participation": "0%",
            "client_led_closures": "0%",
        }
    
    # Count touches per incident (from "Touched By" field)
    touch_counts = []
    client_participated = 0
    client_closed = 0
    
    for inc in incidents:
        if inc.touched_by:
            # Count unique users in touched_by
            users = [u.strip() for u in inc.touched_by.split(",")]
            touch_counts.append(len(users))
            
            # Check if client participated
            for user in users:
                if "CRITICALSTART" not in user.upper():
                    client_participated += 1
                    break
        
        if inc.closed_by and "CRITICALSTART" not in inc.closed_by.upper():
            client_closed += 1
    
    avg_touches = sum(touch_counts) / len(touch_counts) if touch_counts else 0
    participation_rate = (client_participated / total * 100) if total > 0 else 0
    closure_rate = (client_closed / total * 100) if total > 0 else 0
    
    return {
        "avg_touches": round(avg_touches, 1),
        "client_participation": f"{round(participation_rate)}%",
        "client_led_closures": f"{round(closure_rate)}%",
    }


def calculate_response_by_priority(incidents: List[Incident], config: ClientConfig) -> List[Dict[str, Any]]:
    """Calculate response metrics by priority level.
    
    Args:
        incidents: List of incidents
        config: Client configuration
        
    Returns:
        List of response metrics by priority
    """
    priority_metrics = {}
    
    for inc in incidents:
        priority = parse_priority_level(inc.current_priority)
        if not priority or not inc.cs_soc_ttr:
            continue
        
        if priority not in priority_metrics:
            priority_metrics[priority] = {
                "count": 0,
                "ttr_values": [],
                "target": config.sla_targets.get(priority, 180),
            }
        
        priority_metrics[priority]["count"] += 1
        priority_metrics[priority]["ttr_values"].append(timedelta_to_minutes(inc.cs_soc_ttr))
    
    results = []
    priority_order = ["Critical", "High", "Medium", "Low"]
    
    for priority in priority_order:
        if priority not in priority_metrics:
            continue
        
        metrics = priority_metrics[priority]
        avg_response = int(sum(metrics["ttr_values"]) / len(metrics["ttr_values"])) if metrics["ttr_values"] else 0
        target = metrics["target"]
        
        results.append({
            "priority": priority,
            "count": metrics["count"],
            "avg_response": avg_response,
            "target": target,
            "met_sla": avg_response <= target,
        })
    
    return results


def calculate_cost_avoidance(incidents: List[Incident], config: ClientConfig, 
                             response_metrics: Dict[str, Any]) -> Dict[str, Any]:
    """Calculate cost avoidance metrics.
    
    Args:
        incidents: List of incidents
        config: Client configuration
        response_metrics: Pre-calculated response metrics
        
    Returns:
        Dictionary of cost avoidance metrics
    """
    total_incidents = len(incidents)
    
    # Analyst hours saved (based on avg 1.5 hours per incident)
    analyst_hours = int(total_incidents * 1.5)
    analyst_cost = analyst_hours * config.analyst_hourly_rate
    
    # Get period days for coverage calculation
    dates = []
    for inc in incidents:
        if inc.created_datetime_utc:
            dates.append(inc.created_datetime_utc)
    
    if dates:
        period_days = (max(dates) - min(dates)).days + 1
    else:
        period_days = 30
    
    # 24x7 coverage cost equivalent
    coverage_hours = period_days * 24
    coverage_cost = int(coverage_hours * config.coverage_hourly_rate / 10)  # Scaled down
    
    # Breach exposure avoided (true positives * breach cost factor)
    true_positives = sum(1 for inc in incidents 
                         if inc.cs_soc_verdict and "TRUE POSITIVE" in inc.cs_soc_verdict.upper())
    breach_exposure = true_positives * int(config.breach_cost_estimate * 0.15)  # 15% per threat
    
    total_modeled = analyst_cost + coverage_cost + breach_exposure
    
    return {
        "analyst_hours": analyst_hours,
        "analyst_cost_equivalent": analyst_cost,
        "coverage_cost_equivalent": coverage_cost,
        "breach_exposure_avoided": breach_exposure,
        "total_modeled": total_modeled,
        "coverage_hours": coverage_hours,
    }


def calculate_industry_comparison(response_metrics: Dict[str, Any], 
                                   incidents_per_day: float,
                                   config: ClientConfig) -> List[Dict[str, Any]]:
    """Calculate industry comparison metrics.
    
    Args:
        response_metrics: Pre-calculated response metrics
        incidents_per_day: Calculated incidents per day
        config: Client configuration
        
    Returns:
        List of industry comparison dictionaries
    """
    mttr = response_metrics.get("mttr_minutes", 0)
    mttd = response_metrics.get("mttd_minutes", 0)
    
    comparisons = []
    
    # MTTR comparison
    if mttr > 0:
        mttr_diff = ((config.industry_mttr_minutes - mttr) / config.industry_mttr_minutes * 100)
        comparisons.append({
            "metric": "MTTR",
            "yours": mttr,
            "industry": config.industry_mttr_minutes,
            "difference": f"{abs(round(mttr_diff))}% {'Better' if mttr_diff > 0 else 'Slower'}",
        })
    
    # MTTD comparison
    if mttd > 0:
        mttd_diff = ((config.industry_mttd_minutes - mttd) / config.industry_mttd_minutes * 100)
        comparisons.append({
            "metric": "MTTD",
            "yours": mttd,
            "industry": config.industry_mttd_minutes,
            "difference": f"{abs(round(mttd_diff))}% {'Better' if mttd_diff > 0 else 'Slower'}",
        })
    
    # Incidents/Day comparison
    if incidents_per_day > 0:
        ipd_diff = ((config.industry_incidents_per_day - incidents_per_day) / config.industry_incidents_per_day * 100)
        comparisons.append({
            "metric": "Incidents/Day",
            "yours": incidents_per_day,
            "industry": config.industry_incidents_per_day,
            "difference": f"{abs(round(ipd_diff))}% {'Better' if ipd_diff > 0 else 'Higher'}",
        })
    
    return comparisons


def calculate_trend_data(all_periods: List[List[Incident]], config: ClientConfig) -> Dict[str, Any]:
    """Calculate trend data across multiple periods.
    
    Args:
        all_periods: List of incident lists (one per period, chronological order)
        config: Client configuration
        
    Returns:
        Dictionary with trend arrays and period labels
    """
    mttr_trend = []
    mttd_trend = []
    fp_trend = []
    period_labels = []
    
    num_periods = len(all_periods)
    
    for i, incidents in enumerate(all_periods):
        metrics = calculate_period_metrics(incidents, config)
        
        mttr_trend.append(metrics["mttr_minutes"])
        mttd_trend.append(metrics["mttd_minutes"])
        fp_trend.append(metrics["false_positive_rate"])
        
        # Generate period label
        if i == num_periods - 1:
            period_labels.append("Current")
        else:
            periods_back = num_periods - 1 - i
            period_labels.append(f"Period -{periods_back}")
    
    return {
        "mttr_trend": mttr_trend,
        "mttd_trend": mttd_trend,
        "fp_trend": fp_trend,
        "period_labels": period_labels,
    }


def calculate_all_metrics(all_periods: List[List[Incident]], 
                          client_name: str,
                          config: ClientConfig) -> Dict[str, Any]:
    """Calculate all metrics from incident data.
    
    This is the main entry point for metric calculation.
    
    Args:
        all_periods: List of incident lists (one per period)
        client_name: Client name (derived from data)
        config: Client configuration
        
    Returns:
        Dictionary of all metrics matching ReportData fields
    """
    # Current period is the last one
    current_period = all_periods[-1] if all_periods else []
    
    # Calculate date range
    dates = []
    for inc in current_period:
        if inc.created_datetime_utc:
            dates.append(inc.created_datetime_utc)
    
    if dates:
        period_start = min(dates)
        period_end = max(dates)
        period_days = (period_end - period_start).days + 1
    else:
        period_start = datetime.now()
        period_end = datetime.now()
        period_days = 1
    
    # Calculate all metric groups
    volume = calculate_volume_metrics(current_period, config)
    response = calculate_response_metrics(current_period, config)
    detection_sources = calculate_detection_sources(current_period)
    mitre = calculate_mitre_data(current_period)
    severity_flows = calculate_severity_flows(current_period)
    after_hours = calculate_after_hours_metrics(current_period, config)
    detection_quality = calculate_detection_quality(current_period)
    collaboration = calculate_collaboration_metrics(current_period)
    response_by_priority = calculate_response_by_priority(current_period, config)
    trends = calculate_trend_data(all_periods, config)
    
    # Derived metrics
    incidents_per_day = round(len(current_period) / period_days, 1) if period_days > 0 else 0
    
    cost = calculate_cost_avoidance(current_period, config, response)
    industry_comparison = calculate_industry_comparison(response, incidents_per_day, config)
    
    # Combine all metrics
    metrics = {
        # Client Info
        "client_name": config.client_name_override or client_name,
        "tier": config.tier,
        "period_start": config.period_start or period_start.strftime("%B %d, %Y"),
        "period_end": config.period_end or period_end.strftime("%B %d, %Y"),
        "period_days": period_days,
        "report_date": datetime.now().strftime("%B %d, %Y"),
        
        # Volume metrics
        **volume,
        
        # Response metrics
        **response,
        
        # Detection sources
        "detection_sources": detection_sources,
        
        # MITRE data
        **mitre,
        
        # Severity flows
        "severity_flows": severity_flows,
        
        # After-hours
        **after_hours,
        
        # Detection quality
        **detection_quality,
        
        # Collaboration
        **collaboration,
        
        # Response by priority
        "response_by_priority": response_by_priority,
        
        # Industry comparison
        "industry_comparison": industry_comparison,
        
        # Trends
        **trends,
        
        # Cost avoidance
        **cost,
        
        # Derived
        "incidents_per_day": incidents_per_day,
        "client_touch_decisions": volume["incidents_escalated"] - volume["closed_end_to_end"],
        "threats_blocked": volume["true_threats_contained"],
        "zero_breaches": True,  # Assumed unless data indicates otherwise
    }
    
    return metrics
