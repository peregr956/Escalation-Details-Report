from dataclasses import dataclass, field
from typing import Dict, Any, Optional, List


@dataclass
class ReportData:
    """Data class to hold all report metrics."""
    
    # Client Info
    client_name: str = ""
    tier: str = ""
    period_start: str = ""
    period_end: str = ""
    period_days: int = 0
    report_date: str = ""
    
    # Hero Metrics
    alerts_triaged: int = 0
    client_touch_decisions: int = 0
    closed_end_to_end: int = 0
    true_threats_contained: int = 0
    response_advantage_percent: float = 0.0
    mttr_minutes: int = 0
    p90_minutes: int = 0
    industry_median_minutes: int = 0
    after_hours_escalations: int = 0
    coverage_hours: int = 0
    automation_percent: float = 0.0
    
    # Executive Summary
    incidents_escalated: int = 0
    incidents_per_day: float = 0.0
    false_positive_rate: float = 0.0
    
    # Cost Avoidance
    total_modeled: int = 0
    analyst_hours: int = 0
    analyst_cost_equivalent: int = 0
    coverage_cost_equivalent: int = 0
    breach_exposure_avoided: int = 0
    
    # Performance Metrics
    critical_high_mttr: int = 0
    medium_low_mttr: int = 0
    mttd_minutes: int = 0
    containment_rate: float = 0.0
    
    # Detection Quality
    true_threat_precision: float = 0.0
    signal_fidelity: float = 0.0
    client_validated: float = 0.0
    
    # Industry Comparison (list of dicts)
    industry_comparison: List[Dict[str, Any]] = field(default_factory=list)
    
    # Detection Sources (list of dicts)
    detection_sources: List[Dict[str, Any]] = field(default_factory=list)
    
    # Escalation Methods
    playbook_auto: Dict[str, Any] = field(default_factory=dict)
    analyst_escalation: Dict[str, Any] = field(default_factory=dict)
    
    # Trend Data (for charts)
    mttr_trend: List[int] = field(default_factory=list)
    mttd_trend: List[int] = field(default_factory=list)
    fp_trend: List[float] = field(default_factory=list)
    period_labels: List[str] = field(default_factory=list)
    
    # Operational Load
    business_hours_percent: float = 0.0
    after_hours_percent: float = 0.0
    weekend_percent: float = 0.0
    
    # MITRE Data
    tactics: List[str] = field(default_factory=list)
    high_severity: List[int] = field(default_factory=list)
    medium_severity: List[int] = field(default_factory=list)
    low_severity: List[int] = field(default_factory=list)
    info_severity: List[int] = field(default_factory=list)
    
    # Severity Flow Data (list of dicts)
    severity_flows: List[Dict[str, Any]] = field(default_factory=list)
    
    # Improvement Items (list of dicts)
    improvement_items: List[Dict[str, Any]] = field(default_factory=list)
    
    # Collaboration Metrics
    avg_touches: float = 0.0
    client_participation: str = ""
    client_led_closures: str = ""


def get_report_data() -> ReportData:
    """Returns a populated ReportData instance with all metrics from the HTML report."""
    
    return ReportData(
        # Client Info
        client_name="Lennar Corporation",
        tier="Signature Tier",
        period_start="August 1, 2025",
        period_end="August 31, 2025",
        period_days=31,
        report_date="November 5, 2025",
        
        # Hero Metrics
        alerts_triaged=2110,
        client_touch_decisions=1690,
        closed_end_to_end=420,
        true_threats_contained=11,
        response_advantage_percent=34,
        mttr_minutes=126,
        p90_minutes=87,
        industry_median_minutes=192,
        after_hours_escalations=158,
        coverage_hours=744,
        automation_percent=86,
        
        # Executive Summary
        incidents_escalated=267,
        incidents_per_day=8.9,
        false_positive_rate=9.0,
        
        # Cost Avoidance
        total_modeled=7550000,
        analyst_hours=452,
        analyst_cost_equivalent=38000,
        coverage_cost_equivalent=163000,
        breach_exposure_avoided=7340000,
        
        # Performance Metrics
        critical_high_mttr=67,
        medium_low_mttr=52,
        mttd_minutes=4,
        containment_rate=98,
        
        # Detection Quality
        true_threat_precision=31.4,
        signal_fidelity=91,
        client_validated=86.9,
        
        # Industry Comparison
        industry_comparison=[
            {
                "metric": "MTTR",
                "yours": 126,
                "industry": 192,
                "difference": "34% Better"
            },
            {
                "metric": "MTTD",
                "yours": 42,
                "industry": 66,
                "difference": "36% Better"
            },
            {
                "metric": "Incidents/Day",
                "yours": 8.9,
                "industry": 11.4,
                "difference": "22% Better"
            }
        ],
        
        # Detection Sources
        detection_sources=[
            {
                "source": "Palo Alto Cortex XDR",
                "incidents": 189,
                "percent": 70.8,
                "fp_rate": 11.2
            },
            {
                "source": "Microsoft Sentinel",
                "incidents": 52,
                "percent": 19.5,
                "fp_rate": 5.8
            },
            {
                "source": "CrowdStrike Falcon",
                "incidents": 26,
                "percent": 9.7,
                "fp_rate": 7.7
            }
        ],
        
        # Escalation Methods
        playbook_auto={
            "count": 229,
            "percent": 86
        },
        analyst_escalation={
            "count": 38,
            "percent": 14
        },
        
        # Trend Data
        mttr_trend=[168, 150, 126],
        mttd_trend=[54, 48, 42],
        fp_trend=[12.1, 10.8, 9.0],
        period_labels=["Period -2", "Period -1", "Current"],
        
        # Operational Load
        business_hours_percent=51,
        after_hours_percent=41,
        weekend_percent=8,
        
        # MITRE Data
        tactics=["Persistence", "Defense Evasion", "Execution", "Discovery", "Initial Access"],
        high_severity=[12, 8, 5, 3, 2],
        medium_severity=[38, 31, 22, 18, 12],
        low_severity=[22, 18, 15, 24, 16],
        info_severity=[5, 3, 2, 8, 3],
        
        # Severity Flow Data
        severity_flows=[
            {"from": "Vendor Critical", "to": "CS Critical", "flow": 7},
            {"from": "Vendor Critical", "to": "CS High", "flow": 1},
            {"from": "Vendor High", "to": "CS Critical", "flow": 2},
            {"from": "Vendor High", "to": "CS High", "flow": 7},
            {"from": "Vendor High", "to": "CS Medium", "flow": 4},
            {"from": "Vendor High", "to": "CS Low", "flow": 2},
            {"from": "Vendor Medium", "to": "CS High", "flow": 9},
            {"from": "Vendor Medium", "to": "CS Medium", "flow": 72},
            {"from": "Vendor Medium", "to": "CS Low", "flow": 32},
            {"from": "Vendor Medium", "to": "CS Informational", "flow": 8},
            {"from": "Vendor Low", "to": "CS Medium", "flow": 18},
            {"from": "Vendor Low", "to": "CS Low", "flow": 67},
            {"from": "Vendor Low", "to": "CS Informational", "flow": 18},
            {"from": "Vendor Informational", "to": "CS Low", "flow": 5},
            {"from": "Vendor Informational", "to": "CS Informational", "flow": 15}
        ],
        
        # Improvement Items
        improvement_items=[
            {
                "title": "Detection Tuning",
                "priority": "HIGH",
                "owner": "CS SOC + Lennar Security Team",
                "target": "Next 30 days",
                "description": "Palo Alto Cortex XDR false positive rate is 11.2%, exceeding the 10.0% threshold and keeping the overall rate at 9.0%. Tuning these alerts will reduce client noise and improve SOC efficiency, directly lowering the 1,690 client-touch decisions surfaced in the hero."
            },
            {
                "title": "Automation Opportunity",
                "priority": "MEDIUM",
                "owner": "CS SOC Engineering",
                "target": "Next 60 days",
                "description": "Manual escalations at 14% exceed our 12% target. 38 incidents required analyst judgment. Expanding playbook coverage will improve consistency."
            },
            {
                "title": "Threat Focus",
                "priority": "HIGH",
                "owner": "Joint - CS Threat Intel + Lennar",
                "target": "Ongoing",
                "description": "Persistence plus Defense Evasion account for 20 of the 30 high-severity incidents (67%) in the MITRE dataset, signaling concentrated foothold attempts that should drive proactive hunts and new detections."
            }
        ],
        
        # Collaboration Metrics
        avg_touches=2.3,
        client_participation="72%",
        client_led_closures="21%"
    )
