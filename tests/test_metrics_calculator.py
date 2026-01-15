"""Unit tests for metrics_calculator.py

This module provides comprehensive test coverage for the metrics calculation
functions that compute aggregated metrics from incident data.
"""
import pytest
from datetime import datetime, timedelta
from typing import List

from data_parser import Incident
from metrics_calculator import (
    ClientConfig,
    parse_priority_level,
    parse_vendor_severity,
    is_after_hours,
    is_weekend,
    timedelta_to_minutes,
    calculate_percentile,
    calculate_period_metrics,
    calculate_volume_metrics,
    calculate_response_metrics,
    calculate_detection_sources,
    calculate_mitre_data,
    calculate_severity_flows,
    calculate_after_hours_metrics,
    calculate_detection_quality,
    calculate_collaboration_metrics,
    calculate_response_by_priority,
    calculate_cost_avoidance,
    calculate_industry_comparison,
    calculate_trend_data,
    calculate_monthly_trend_data,
)


# =============================================================================
# Fixtures
# =============================================================================

@pytest.fixture
def default_config():
    """Return a default ClientConfig for testing."""
    return ClientConfig()


@pytest.fixture
def sample_incident():
    """Return a basic sample incident for testing."""
    return Incident(
        incident_id=1,
        organization="Test Client",
        product="Test EDR",
        current_status="Closed",
        current_priority="3 - HIGH",
        cs_soc_verdict="True Positive",
        cs_soc_ttr=timedelta(hours=1, minutes=30),  # 90 minutes
        cs_soc_ttd=timedelta(minutes=45),
        created_datetime_utc=datetime(2025, 6, 15, 10, 0),
        escalated_datetime_utc=datetime(2025, 6, 15, 10, 30),
        mitre_tactic_name="Persistence",
        vendor_severity="High",
        initial_escalation_method="Playbook",
    )


@pytest.fixture
def sample_incidents(sample_incident):
    """Return a list of sample incidents for testing."""
    incidents = []
    
    # Base incident (True Positive, High priority)
    incidents.append(sample_incident)
    
    # False positive incident
    incidents.append(Incident(
        incident_id=2,
        organization="Test Client",
        product="Test EDR",
        current_status="Closed",
        current_priority="5 - MEDIUM",
        cs_soc_verdict="False Positive",
        cs_soc_ttr=timedelta(hours=2),  # 120 minutes
        cs_soc_ttd=timedelta(minutes=30),
        created_datetime_utc=datetime(2025, 6, 15, 14, 0),
        escalated_datetime_utc=datetime(2025, 6, 15, 14, 15),
        mitre_tactic_name="Defense Evasion",
        vendor_severity="Medium",
        initial_escalation_method="CS SOC",
    ))
    
    # Critical priority incident (after hours)
    incidents.append(Incident(
        incident_id=3,
        organization="Test Client",
        product="Firewall",
        current_status="Closed",
        current_priority="1 - CRITICAL",
        cs_soc_verdict="True Positive",
        cs_soc_ttr=timedelta(minutes=25),
        cs_soc_ttd=timedelta(minutes=10),
        created_datetime_utc=datetime(2025, 6, 14, 22, 0),  # 10 PM - after hours
        escalated_datetime_utc=datetime(2025, 6, 14, 22, 15),
        mitre_tactic_name="Persistence",
        vendor_severity="Critical",
        initial_escalation_method="Playbook",
        response_action="Block IP",
        response_action_status="Success",
    ))
    
    # Weekend incident
    incidents.append(Incident(
        incident_id=4,
        organization="Test Client",
        product="Test EDR",
        current_status="Closed",
        current_priority="6 - LOW",
        cs_soc_verdict="Benign",
        cs_soc_ttr=timedelta(hours=3),  # 180 minutes
        cs_soc_ttd=timedelta(minutes=60),
        created_datetime_utc=datetime(2025, 6, 14, 10, 0),  # Saturday
        escalated_datetime_utc=datetime(2025, 6, 14, 10, 30),
        mitre_tactic_name="Discovery",
        vendor_severity="Low",
        initial_escalation_method="Playbook",
    ))
    
    return incidents


# =============================================================================
# Tests for parse_priority_level
# =============================================================================

class TestParsePriorityLevel:
    """Tests for the parse_priority_level function."""
    
    def test_critical_variations(self):
        """Test Critical priority parsing."""
        assert parse_priority_level("Critical") == "Critical"
        assert parse_priority_level("CRITICAL") == "Critical"
        assert parse_priority_level("1 - CRITICAL") == "Critical"
    
    def test_high_variations(self):
        """Test High priority parsing."""
        assert parse_priority_level("High") == "High"
        assert parse_priority_level("HIGH") == "High"
        assert parse_priority_level("2 - HIGH") == "High"
        assert parse_priority_level("3 - HIGH") == "High"
    
    def test_medium_variations(self):
        """Test Medium priority parsing."""
        assert parse_priority_level("Medium") == "Medium"
        assert parse_priority_level("MEDIUM") == "Medium"
        assert parse_priority_level("4 - MEDIUM") == "Medium"
        assert parse_priority_level("5 - MEDIUM") == "Medium"
    
    def test_low_variations(self):
        """Test Low priority parsing."""
        assert parse_priority_level("Low") == "Low"
        assert parse_priority_level("LOW") == "Low"
        assert parse_priority_level("6 - LOW") == "Low"
        assert parse_priority_level("7 - LOW") == "Low"
    
    def test_informational_variations(self):
        """Test Informational priority parsing."""
        assert parse_priority_level("Info") == "Informational"
        assert parse_priority_level("INFORMATIONAL") == "Informational"
        assert parse_priority_level("8 - INFORMATIONAL") == "Informational"
    
    def test_none_and_empty(self):
        """Test None and empty string handling."""
        assert parse_priority_level(None) is None
        assert parse_priority_level("") is None
    
    def test_unknown_priority(self):
        """Test unknown priority strings."""
        assert parse_priority_level("Unknown") is None
        assert parse_priority_level("Random") is None


# =============================================================================
# Tests for parse_vendor_severity
# =============================================================================

class TestParseVendorSeverity:
    """Tests for the parse_vendor_severity function."""
    
    def test_severity_parsing(self):
        """Test all severity levels."""
        assert parse_vendor_severity("Critical") == "Critical"
        assert parse_vendor_severity("High") == "High"
        assert parse_vendor_severity("Medium") == "Medium"
        assert parse_vendor_severity("Med") == "Medium"
        assert parse_vendor_severity("Low") == "Low"
        assert parse_vendor_severity("Informational") == "Informational"
        assert parse_vendor_severity("Info") == "Informational"
    
    def test_none_and_empty(self):
        """Test None and empty string handling."""
        assert parse_vendor_severity(None) is None
        assert parse_vendor_severity("") is None


# =============================================================================
# Tests for is_after_hours
# =============================================================================

class TestIsAfterHours:
    """Tests for the is_after_hours function."""
    
    def test_business_hours(self, default_config):
        """Test that business hours are correctly identified."""
        # 10 AM on a Tuesday - business hours
        dt = datetime(2025, 6, 17, 10, 0)  # Tuesday
        assert is_after_hours(dt, default_config) is False
    
    def test_early_morning(self, default_config):
        """Test early morning (before 8 AM) is after hours."""
        dt = datetime(2025, 6, 17, 6, 0)  # 6 AM Tuesday
        assert is_after_hours(dt, default_config) is True
    
    def test_evening(self, default_config):
        """Test evening (after 6 PM) is after hours."""
        dt = datetime(2025, 6, 17, 20, 0)  # 8 PM Tuesday
        assert is_after_hours(dt, default_config) is True
    
    def test_weekend_saturday(self, default_config):
        """Test Saturday is after hours."""
        dt = datetime(2025, 6, 14, 12, 0)  # Noon Saturday
        assert is_after_hours(dt, default_config) is True
    
    def test_weekend_sunday(self, default_config):
        """Test Sunday is after hours."""
        dt = datetime(2025, 6, 15, 12, 0)  # Noon Sunday
        # Note: June 15, 2025 is actually a Sunday
        assert is_after_hours(dt, default_config) is True
    
    def test_none_datetime(self, default_config):
        """Test None datetime returns False."""
        assert is_after_hours(None, default_config) is False
    
    def test_custom_business_hours(self):
        """Test with custom business hours configuration."""
        config = ClientConfig(business_hours_start=9, business_hours_end=17)
        
        # 8 AM should be after hours with 9-5 schedule
        dt = datetime(2025, 6, 17, 8, 0)  # Tuesday
        assert is_after_hours(dt, config) is True
        
        # 9 AM should be business hours
        dt = datetime(2025, 6, 17, 9, 0)
        assert is_after_hours(dt, config) is False


# =============================================================================
# Tests for is_weekend
# =============================================================================

class TestIsWeekend:
    """Tests for the is_weekend function."""
    
    def test_weekday(self):
        """Test weekdays are not weekends."""
        # Monday through Friday
        assert is_weekend(datetime(2025, 6, 16, 12, 0)) is False  # Monday
        assert is_weekend(datetime(2025, 6, 17, 12, 0)) is False  # Tuesday
        assert is_weekend(datetime(2025, 6, 18, 12, 0)) is False  # Wednesday
        assert is_weekend(datetime(2025, 6, 19, 12, 0)) is False  # Thursday
        assert is_weekend(datetime(2025, 6, 20, 12, 0)) is False  # Friday
    
    def test_weekend(self):
        """Test Saturday and Sunday are weekends."""
        assert is_weekend(datetime(2025, 6, 14, 12, 0)) is True  # Saturday
        assert is_weekend(datetime(2025, 6, 15, 12, 0)) is True  # Sunday
    
    def test_none(self):
        """Test None returns False."""
        assert is_weekend(None) is False


# =============================================================================
# Tests for timedelta_to_minutes
# =============================================================================

class TestTimedeltaToMinutes:
    """Tests for the timedelta_to_minutes function."""
    
    def test_basic_conversion(self):
        """Test basic timedelta to minutes conversion."""
        assert timedelta_to_minutes(timedelta(hours=1)) == 60
        assert timedelta_to_minutes(timedelta(minutes=30)) == 30
        assert timedelta_to_minutes(timedelta(hours=2, minutes=30)) == 150
    
    def test_none(self):
        """Test None returns 0."""
        assert timedelta_to_minutes(None) == 0
    
    def test_zero(self):
        """Test zero timedelta returns 0."""
        assert timedelta_to_minutes(timedelta(0)) == 0


# =============================================================================
# Tests for calculate_percentile
# =============================================================================

class TestCalculatePercentile:
    """Tests for the calculate_percentile function."""
    
    def test_p90(self):
        """Test 90th percentile calculation."""
        values = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100]
        assert calculate_percentile(values, 90) == 100
    
    def test_p50(self):
        """Test 50th percentile (median) calculation."""
        values = [10, 20, 30, 40, 50]
        assert calculate_percentile(values, 50) == 30
    
    def test_empty_list(self):
        """Test empty list returns 0."""
        assert calculate_percentile([], 90) == 0
    
    def test_single_value(self):
        """Test single value list."""
        assert calculate_percentile([42], 90) == 42


# =============================================================================
# Tests for calculate_period_metrics
# =============================================================================

class TestCalculatePeriodMetrics:
    """Tests for the calculate_period_metrics function."""
    
    def test_with_incidents(self, sample_incidents, default_config):
        """Test metric calculation with sample incidents."""
        metrics = calculate_period_metrics(sample_incidents, default_config)
        
        assert metrics["incidents_escalated"] == 4
        assert metrics["mttr_minutes"] > 0
        assert metrics["mttd_minutes"] > 0
        assert metrics["p90_minutes"] >= metrics["mttr_minutes"]
        assert 0 <= metrics["false_positive_rate"] <= 100
    
    def test_empty_incidents(self, default_config):
        """Test metric calculation with empty incident list."""
        metrics = calculate_period_metrics([], default_config)
        
        assert metrics["incidents_escalated"] == 0
        assert metrics["mttr_minutes"] == 0
        assert metrics["mttd_minutes"] == 0
        assert metrics["p90_minutes"] == 0
        assert metrics["false_positive_rate"] == 0.0
    
    def test_false_positive_rate_calculation(self, default_config):
        """Test false positive rate is calculated correctly."""
        # 1 FP out of 4 = 25%
        incidents = [
            Incident(organization="Test", current_status="Closed", 
                    current_priority="High", cs_soc_verdict="False Positive",
                    cs_soc_ttr=timedelta(hours=1)),
            Incident(organization="Test", current_status="Closed",
                    current_priority="High", cs_soc_verdict="True Positive",
                    cs_soc_ttr=timedelta(hours=1)),
            Incident(organization="Test", current_status="Closed",
                    current_priority="High", cs_soc_verdict="True Positive",
                    cs_soc_ttr=timedelta(hours=1)),
            Incident(organization="Test", current_status="Closed",
                    current_priority="High", cs_soc_verdict="True Positive",
                    cs_soc_ttr=timedelta(hours=1)),
        ]
        
        metrics = calculate_period_metrics(incidents, default_config)
        assert metrics["false_positive_rate"] == 25.0


# =============================================================================
# Tests for calculate_volume_metrics
# =============================================================================

class TestCalculateVolumeMetrics:
    """Tests for the calculate_volume_metrics function."""
    
    def test_volume_metrics(self, sample_incidents, default_config):
        """Test volume metric calculation."""
        metrics = calculate_volume_metrics(sample_incidents, default_config)
        
        assert metrics["alerts_triaged"] == 4
        assert metrics["incidents_escalated"] == 4
        assert "closed_end_to_end" in metrics
        assert "true_threats_contained" in metrics
        assert "playbook_auto" in metrics
        assert "analyst_escalation" in metrics
    
    def test_automation_percentage(self, sample_incidents, default_config):
        """Test automation percentage is calculated."""
        metrics = calculate_volume_metrics(sample_incidents, default_config)
        
        # Verify playbook vs analyst breakdown sums to 100%
        total_percent = (metrics["playbook_auto"]["percent"] + 
                        metrics["analyst_escalation"]["percent"])
        assert total_percent == 100


# =============================================================================
# Tests for calculate_response_metrics
# =============================================================================

class TestCalculateResponseMetrics:
    """Tests for the calculate_response_metrics function."""
    
    def test_response_metrics(self, sample_incidents, default_config):
        """Test response metric calculation."""
        metrics = calculate_response_metrics(sample_incidents, default_config)
        
        assert "mttr_minutes" in metrics
        assert "mttd_minutes" in metrics
        assert "p90_minutes" in metrics
        assert "critical_high_mttr" in metrics
        assert "medium_low_mttr" in metrics
        assert "response_advantage_percent" in metrics
        assert "sla_compliance_rate" in metrics
    
    def test_response_advantage_calculation(self, default_config):
        """Test response advantage is calculated relative to industry."""
        incidents = [
            Incident(
                organization="Test",
                current_status="Closed",
                current_priority="High",
                cs_soc_ttr=timedelta(minutes=96),  # 50% of 192 industry median
            )
        ]
        
        metrics = calculate_response_metrics(incidents, default_config)
        
        # Should be 50% better than industry (192 - 96) / 192 = 0.5
        assert metrics["response_advantage_percent"] == 50.0


# =============================================================================
# Tests for calculate_detection_sources
# =============================================================================

class TestCalculateDetectionSources:
    """Tests for the calculate_detection_sources function."""
    
    def test_detection_sources(self, sample_incidents):
        """Test detection source breakdown."""
        sources = calculate_detection_sources(sample_incidents)
        
        assert isinstance(sources, list)
        assert len(sources) > 0
        
        # Each source should have required fields
        for source in sources:
            assert "source" in source
            assert "incidents" in source
            assert "percent" in source
            assert "fp_rate" in source
    
    def test_source_percentages_sum_to_100(self, sample_incidents):
        """Test that source percentages sum to ~100%."""
        sources = calculate_detection_sources(sample_incidents)
        
        total_percent = sum(s["percent"] for s in sources)
        assert 99.0 <= total_percent <= 101.0  # Allow rounding tolerance


# =============================================================================
# Tests for calculate_mitre_data
# =============================================================================

class TestCalculateMitreData:
    """Tests for the calculate_mitre_data function."""
    
    def test_mitre_data(self, sample_incidents):
        """Test MITRE ATT&CK data calculation."""
        mitre = calculate_mitre_data(sample_incidents)
        
        assert "tactics" in mitre
        assert "high_severity" in mitre
        assert "medium_severity" in mitre
        assert "low_severity" in mitre
        assert "info_severity" in mitre
        
        # All lists should be same length
        assert len(mitre["tactics"]) == len(mitre["high_severity"])
        assert len(mitre["tactics"]) == len(mitre["medium_severity"])
        assert len(mitre["tactics"]) == len(mitre["low_severity"])


# =============================================================================
# Tests for calculate_severity_flows
# =============================================================================

class TestCalculateSeverityFlows:
    """Tests for the calculate_severity_flows function."""
    
    def test_severity_flows(self, sample_incidents):
        """Test severity flow calculation for Sankey diagram."""
        flows = calculate_severity_flows(sample_incidents)
        
        assert isinstance(flows, list)
        
        for flow in flows:
            assert "from" in flow
            assert "to" in flow
            assert "flow" in flow
            assert flow["from"].startswith("Vendor ")
            assert flow["to"].startswith("CS ")


# =============================================================================
# Tests for calculate_after_hours_metrics
# =============================================================================

class TestCalculateAfterHoursMetrics:
    """Tests for the calculate_after_hours_metrics function."""
    
    def test_after_hours_metrics(self, sample_incidents, default_config):
        """Test after-hours metric calculation."""
        metrics = calculate_after_hours_metrics(sample_incidents, default_config)
        
        assert "after_hours_escalations" in metrics
        assert "after_hours_weeknight" in metrics
        assert "after_hours_weekend" in metrics
        assert "business_hours_percent" in metrics
        assert "after_hours_percent" in metrics
        assert "weekend_percent" in metrics
    
    def test_percentages_sum_to_100(self, sample_incidents, default_config):
        """Test that hour percentages sum to ~100%."""
        metrics = calculate_after_hours_metrics(sample_incidents, default_config)
        
        total = (metrics["business_hours_percent"] + 
                metrics["after_hours_percent"] + 
                metrics["weekend_percent"])
        assert 99.0 <= total <= 101.0


# =============================================================================
# Tests for calculate_detection_quality
# =============================================================================

class TestCalculateDetectionQuality:
    """Tests for the calculate_detection_quality function."""
    
    def test_detection_quality(self, sample_incidents):
        """Test detection quality metric calculation."""
        quality = calculate_detection_quality(sample_incidents)
        
        assert "true_threat_precision" in quality
        assert "signal_fidelity" in quality
        assert "false_positive_rate" in quality
        assert "containment_rate" in quality
    
    def test_empty_incidents(self):
        """Test with empty incident list."""
        quality = calculate_detection_quality([])
        
        assert quality["true_threat_precision"] == 0.0
        assert quality["signal_fidelity"] == 0.0
        assert quality["false_positive_rate"] == 0.0


# =============================================================================
# Tests for calculate_collaboration_metrics
# =============================================================================

class TestCalculateCollaborationMetrics:
    """Tests for the calculate_collaboration_metrics function."""
    
    def test_collaboration_metrics(self):
        """Test collaboration metric calculation."""
        incidents = [
            Incident(
                organization="Test",
                current_status="Closed",
                current_priority="High",
                touched_by="analyst@criticalstart.com, client@company.com",
                closed_by="client@company.com",
            ),
            Incident(
                organization="Test",
                current_status="Closed",
                current_priority="Medium",
                touched_by="analyst@criticalstart.com",
                closed_by="analyst@criticalstart.com",
            ),
        ]
        
        metrics = calculate_collaboration_metrics(incidents)
        
        assert "avg_touches" in metrics
        assert "client_participation" in metrics
        assert "client_led_closures" in metrics
    
    def test_empty_incidents(self):
        """Test with empty incident list."""
        metrics = calculate_collaboration_metrics([])
        
        assert metrics["avg_touches"] == 0.0
        assert metrics["client_participation"] == "0%"
        assert metrics["client_led_closures"] == "0%"


# =============================================================================
# Tests for calculate_response_by_priority
# =============================================================================

class TestCalculateResponseByPriority:
    """Tests for the calculate_response_by_priority function."""
    
    def test_response_by_priority(self, sample_incidents, default_config):
        """Test response metrics by priority level."""
        by_priority = calculate_response_by_priority(sample_incidents, default_config)
        
        assert isinstance(by_priority, list)
        
        for entry in by_priority:
            assert "priority" in entry
            assert "count" in entry
            assert "avg_response" in entry
            assert "target" in entry
            assert "met_sla" in entry


# =============================================================================
# Tests for calculate_cost_avoidance
# =============================================================================

class TestCalculateCostAvoidance:
    """Tests for the calculate_cost_avoidance function."""
    
    def test_cost_avoidance(self, sample_incidents, default_config):
        """Test cost avoidance calculation."""
        response = calculate_response_metrics(sample_incidents, default_config)
        cost = calculate_cost_avoidance(sample_incidents, default_config, response)
        
        assert "analyst_hours" in cost
        assert "analyst_cost_equivalent" in cost
        assert "coverage_cost_equivalent" in cost
        assert "breach_exposure_avoided" in cost
        assert "total_modeled" in cost
        
        # Total should equal sum of components
        expected_total = (cost["analyst_cost_equivalent"] + 
                         cost["coverage_cost_equivalent"] + 
                         cost["breach_exposure_avoided"])
        assert cost["total_modeled"] == expected_total


# =============================================================================
# Tests for calculate_industry_comparison
# =============================================================================

class TestCalculateIndustryComparison:
    """Tests for the calculate_industry_comparison function."""
    
    def test_industry_comparison(self, default_config):
        """Test industry comparison calculation."""
        response = {"mttr_minutes": 100, "mttd_minutes": 33}
        incidents_per_day = 5.0
        
        comparisons = calculate_industry_comparison(response, incidents_per_day, default_config)
        
        assert isinstance(comparisons, list)
        assert len(comparisons) == 3  # MTTR, MTTD, Incidents/Day
        
        for comp in comparisons:
            assert "metric" in comp
            assert "yours" in comp
            assert "industry" in comp
            assert "difference" in comp


# =============================================================================
# Tests for calculate_trend_data
# =============================================================================

class TestCalculateTrendData:
    """Tests for the calculate_trend_data function."""
    
    def test_trend_data(self, sample_incidents, default_config):
        """Test trend data calculation across periods."""
        # Create 3 periods
        period1 = sample_incidents[:2]
        period2 = sample_incidents[2:3]
        period3 = sample_incidents
        
        trends = calculate_trend_data([period1, period2, period3], default_config)
        
        assert len(trends["mttr_trend"]) == 3
        assert len(trends["mttd_trend"]) == 3
        assert len(trends["fp_trend"]) == 3
        assert len(trends["period_labels"]) == 3
        assert trends["period_labels"][-1] == "Current"


# =============================================================================
# Tests for calculate_monthly_trend_data
# =============================================================================

class TestCalculateMonthlyTrendData:
    """Tests for the calculate_monthly_trend_data function."""
    
    def test_monthly_trend_data(self, default_config):
        """Test monthly trend calculation from single period spanning months."""
        incidents = [
            # January incident
            Incident(
                organization="Test",
                current_status="Closed",
                current_priority="High",
                cs_soc_ttr=timedelta(hours=2),
                cs_soc_ttd=timedelta(minutes=30),
                escalated_datetime_utc=datetime(2025, 1, 15, 10, 0),
            ),
            # February incident
            Incident(
                organization="Test",
                current_status="Closed",
                current_priority="High",
                cs_soc_ttr=timedelta(hours=1, minutes=30),
                cs_soc_ttd=timedelta(minutes=25),
                escalated_datetime_utc=datetime(2025, 2, 15, 10, 0),
            ),
            # March incident
            Incident(
                organization="Test",
                current_status="Closed",
                current_priority="High",
                cs_soc_ttr=timedelta(hours=1),
                cs_soc_ttd=timedelta(minutes=20),
                escalated_datetime_utc=datetime(2025, 3, 15, 10, 0),
            ),
        ]
        
        trends = calculate_monthly_trend_data(incidents, default_config)
        
        assert len(trends["mttr_trend"]) == 3
        assert len(trends["period_labels"]) == 3
        assert "Jan" in trends["period_labels"][0]
        assert "Feb" in trends["period_labels"][1]
        assert "Mar" in trends["period_labels"][2]


# =============================================================================
# Tests for ClientConfig
# =============================================================================

class TestClientConfig:
    """Tests for the ClientConfig dataclass."""
    
    def test_default_values(self):
        """Test default configuration values."""
        config = ClientConfig()
        
        assert config.tier == "Standard Tier"
        assert config.industry_mttr_minutes == 192
        assert config.industry_mttd_minutes == 66
        assert config.analyst_hourly_rate == 85
        assert config.breach_cost_estimate == 4200000
        assert config.business_hours_start == 8
        assert config.business_hours_end == 18
    
    def test_custom_values(self):
        """Test custom configuration values."""
        config = ClientConfig(
            tier="Signature Tier",
            industry_mttr_minutes=150,
            analyst_hourly_rate=100,
        )
        
        assert config.tier == "Signature Tier"
        assert config.industry_mttr_minutes == 150
        assert config.analyst_hourly_rate == 100
    
    def test_sla_targets(self):
        """Test SLA targets are properly initialized."""
        config = ClientConfig()
        
        assert config.sla_targets["Critical"] == 30
        assert config.sla_targets["High"] == 60
        assert config.sla_targets["Medium"] == 180
        assert config.sla_targets["Low"] == 240


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
