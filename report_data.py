"""Report data module for the Escalation Details Report.

This module contains the ReportData dataclass and functions for loading
report data either from static sample data or dynamically from Excel files.
"""
from dataclasses import dataclass, field, fields
from pathlib import Path
from typing import Dict, Any, Optional, List, Union


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
    
    # After-Hours Details (for detailed breakdown slide)
    after_hours_weeknight: int = 0
    after_hours_weekend: int = 0
    after_hours_critical: int = 0
    after_hours_high: int = 0
    after_hours_medium: int = 0
    after_hours_low: int = 0
    notification_methods: List[Dict[str, Any]] = field(default_factory=list)
    
    # Response Efficiency Details (for response efficiency slide)
    response_by_priority: List[Dict[str, Any]] = field(default_factory=list)
    sla_compliance_rate: float = 0.0
    avg_response_time: int = 0
    fastest_response_time: int = 0
    
    # Detection Quality Context (for detailed detection quality slide)
    detection_quality_trend: str = ""
    tuning_recommendations: List[str] = field(default_factory=list)
    signal_to_noise_ratio: float = 0.0
    
    # Security Outcomes Summary (for comprehensive outcomes slide)
    zero_breaches: bool = True
    threats_blocked: int = 0
    vulnerabilities_identified: int = 0
    compliance_status: str = ""
    risk_reduction_percent: float = 0.0
    
    # Executive Insights (for key takeaways and messaging)
    key_achievements: List[str] = field(default_factory=list)
    areas_of_focus: List[str] = field(default_factory=list)
    next_period_goals: List[str] = field(default_factory=list)
    executive_summary_narrative: str = ""
    
    # New fields added per Jan 2026 stakeholder feedback
    # Customer Success Manager info (for dynamic contact slide)
    csm_name: str = ""
    csm_email: str = ""
    
    # Business hours definition (for Operational Coverage slide clarity)
    business_hours_definition: str = ""  # e.g., "9AM-5PM EST, Mon-Fri"
    
    # Cost calculation methodology (for Value Delivered slide transparency)
    cost_calculation_methodology: str = ""
    cost_calculation_source: str = ""
    
    # Industry benchmarks for comparative visualizations
    mttr_industry_benchmark: float = 0.0
    mttd_industry_benchmark: float = 0.0


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
        client_led_closures="21%",
        
        # After-Hours Details
        after_hours_weeknight=129,
        after_hours_weekend=29,
        after_hours_critical=3,
        after_hours_high=18,
        after_hours_medium=89,
        after_hours_low=48,
        notification_methods=[
            {"method": "Email", "count": 142, "percent": 89.9},
            {"method": "Phone", "count": 12, "percent": 7.6},
            {"method": "Slack/Teams", "count": 4, "percent": 2.5}
        ],
        
        # Response Efficiency Details
        response_by_priority=[
            {"priority": "Critical", "count": 9, "avg_response": 23, "target": 30, "met_sla": True},
            {"priority": "High", "count": 21, "avg_response": 67, "target": 60, "met_sla": False},
            {"priority": "Medium", "count": 165, "avg_response": 126, "target": 180, "met_sla": True},
            {"priority": "Low", "count": 72, "avg_response": 52, "target": 240, "met_sla": True}
        ],
        sla_compliance_rate=94.8,
        avg_response_time=126,
        fastest_response_time=8,
        
        # Detection Quality Context
        detection_quality_trend="improving",
        tuning_recommendations=[
            "Reduce Palo Alto Cortex XDR false positive rate from 11.2% to target 10.0%",
            "Optimize Microsoft Sentinel rules for better signal fidelity",
            "Review CrowdStrike Falcon detection thresholds"
        ],
        signal_to_noise_ratio=9.1,
        
        # Security Outcomes Summary
        zero_breaches=True,
        threats_blocked=11,
        vulnerabilities_identified=23,
        compliance_status="Fully Compliant",
        risk_reduction_percent=34.0,
        
        # Executive Insights
        key_achievements=[
            "34% faster response than industry peers",
            "100% threat containment with zero breaches",
            "158 after-hours escalations handled seamlessly",
            "$7.55M modeled cost exposure avoided"
        ],
        areas_of_focus=[
            "Reduce Palo Alto Cortex XDR false positive rate",
            "Expand playbook automation coverage",
            "Proactive threat hunting for Persistence tactics"
        ],
        next_period_goals=[
            "Achieve 10% or lower overall false positive rate",
            "Reduce manual escalations to 12% or below",
            "Implement enhanced detection for Defense Evasion"
        ],
        executive_summary_narrative="Your security posture remained strong this reporting period. CS SOC triaged 2,110 alerts, partnering with your team on 1,690 decisions and closing 420 end-to-end. Response speed landed 34% faster than sector medians (126-minute MTTR, 87-minute P90), while 158 escalations were absorbed after hours without gaps in coverage. Of the 267 alerts escalated, we identified 11 true positives and contained each before business impact, keeping false positives at 9.0%.",
        
        # New fields per Jan 2026 stakeholder feedback
        csm_name="Sarah Chen",
        csm_email="sarah.chen@criticalstart.com",
        business_hours_definition="9AM-5PM EST, Mon-Fri",
        cost_calculation_methodology="Ponemon Cost of a Data Breach 2025",
        cost_calculation_source="IBM Security / Ponemon Institute",
        mttr_industry_benchmark=192.0,
        mttd_industry_benchmark=66.0
    )


def load_report_data(
    excel_paths: Union[List[Path], List[str]],
    config_path: Optional[Union[Path, str]] = None,
    client_name_override: Optional[str] = None
) -> ReportData:
    """Load and compute report data from Excel files.
    
    This function orchestrates the data loading pipeline:
    1. Parse Excel files into Incident records
    2. Calculate aggregated metrics
    3. Generate insights and recommendations
    4. Return populated ReportData instance
    
    Args:
        excel_paths: List of paths to Excel files (1-3 files, chronological order).
                     The last file is treated as the current period.
        config_path: Optional path to client configuration YAML file.
        client_name_override: Optional override for client name.
        
    Returns:
        ReportData instance with all metrics computed from the input data.
        
    Raises:
        FileNotFoundError: If Excel or config files don't exist.
        ValueError: If required columns are missing or data is invalid.
        
    Example:
        >>> data = load_report_data(
        ...     ["aug.xlsx", "sep.xlsx", "oct.xlsx"],
        ...     config_path="client.yaml"
        ... )
        >>> print(data.client_name, data.mttr_minutes)
    """
    # Import here to avoid circular imports and allow optional dependencies
    from data_parser import load_multiple_periods
    from metrics_calculator import calculate_all_metrics
    from config_loader import load_config
    from insight_generator import generate_all_insights
    
    # Convert paths to Path objects
    excel_paths = [Path(p) for p in excel_paths]
    if config_path:
        config_path = Path(config_path)
    
    # Validate input files exist
    for path in excel_paths:
        if not path.exists():
            raise FileNotFoundError(f"Excel file not found: {path}")
    
    # Load configuration
    config = load_config(config_path)
    
    # Apply client name override if provided
    if client_name_override:
        config.client_name_override = client_name_override
    
    # Parse Excel files
    all_periods, derived_client_name = load_multiple_periods(excel_paths)
    
    # Calculate all metrics
    metrics = calculate_all_metrics(all_periods, derived_client_name, config)
    
    # Generate insights
    insights = generate_all_insights(metrics, config)
    
    # Merge insights into metrics
    metrics.update(insights)
    
    # Get valid field names from ReportData dataclass
    valid_fields = {f.name for f in fields(ReportData)}
    
    # Create and return ReportData instance
    return ReportData(**{
        k: v for k, v in metrics.items()
        if k in valid_fields
    })


def validate_report_data(data: ReportData) -> List[str]:
    """Validate a ReportData instance for completeness.
    
    Args:
        data: ReportData instance to validate.
        
    Returns:
        List of warning messages (empty if valid).
    """
    warnings = []
    
    # Check required fields
    if not data.client_name:
        warnings.append("Client name is empty")
    
    if data.incidents_escalated == 0:
        warnings.append("No incidents found in data")
    
    if data.mttr_minutes == 0:
        warnings.append("MTTR is zero - no response time data")
    
    # Check for reasonable values
    if data.false_positive_rate > 50:
        warnings.append(f"False positive rate ({data.false_positive_rate}%) seems unusually high")
    
    if data.response_advantage_percent < -50:
        warnings.append(f"Response advantage ({data.response_advantage_percent}%) indicates much slower than industry")
    
    # Check list fields
    if not data.detection_sources:
        warnings.append("No detection sources found")
    
    if not data.tactics:
        warnings.append("No MITRE tactics data found")
    
    return warnings
