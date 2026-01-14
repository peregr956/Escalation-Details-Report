"""Configuration loader for client config YAML files.

This module handles loading and parsing client configuration YAML files
into the ClientConfig dataclass used by the metrics calculator.
"""
from pathlib import Path
from typing import Optional, Dict, Any

from metrics_calculator import ClientConfig


def load_config(config_path: Optional[Path] = None) -> ClientConfig:
    """Load client configuration from a YAML file.
    
    Args:
        config_path: Path to the YAML config file. If None, returns defaults.
        
    Returns:
        ClientConfig instance with loaded or default values
        
    Raises:
        FileNotFoundError: If config file doesn't exist
        ValueError: If config file is invalid
    """
    if config_path is None:
        return ClientConfig()
    
    config_path = Path(config_path)
    if not config_path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path}")
    
    try:
        import yaml
    except ImportError:
        raise ImportError("PyYAML is required. Install with: pip install pyyaml")
    
    with open(config_path, 'r') as f:
        raw_config = yaml.safe_load(f)
    
    if not raw_config:
        return ClientConfig()
    
    return parse_config(raw_config)


def parse_config(raw_config: Dict[str, Any]) -> ClientConfig:
    """Parse raw config dictionary into ClientConfig.
    
    Args:
        raw_config: Dictionary loaded from YAML
        
    Returns:
        ClientConfig instance
    """
    config = ClientConfig()
    
    # Client tier
    if "tier" in raw_config:
        config.tier = str(raw_config["tier"])
    
    # Client name override
    if "client_name_override" in raw_config and raw_config["client_name_override"]:
        config.client_name_override = str(raw_config["client_name_override"])
    
    # Industry benchmarks availability
    if "industry_benchmarks_available" in raw_config:
        config.industry_benchmarks_available = bool(raw_config["industry_benchmarks_available"])
    
    # Industry benchmarks
    benchmarks = raw_config.get("industry_benchmarks", {})
    if "mttr_minutes" in benchmarks:
        config.industry_mttr_minutes = int(benchmarks["mttr_minutes"])
    if "mttd_minutes" in benchmarks:
        config.industry_mttd_minutes = int(benchmarks["mttd_minutes"])
    if "incidents_per_day" in benchmarks:
        config.industry_incidents_per_day = float(benchmarks["incidents_per_day"])
    
    # Report period
    period = raw_config.get("report_period", {})
    if period:
        if "start" in period:
            config.period_start = str(period["start"])
        if "end" in period:
            config.period_end = str(period["end"])
    
    # Cost modeling
    cost = raw_config.get("cost_modeling", {})
    if "analyst_hourly_rate" in cost:
        config.analyst_hourly_rate = int(cost["analyst_hourly_rate"])
    if "coverage_hourly_rate" in cost:
        config.coverage_hourly_rate = int(cost["coverage_hourly_rate"])
    if "breach_cost_estimate" in cost:
        config.breach_cost_estimate = int(cost["breach_cost_estimate"])
    
    # SLA targets
    sla = raw_config.get("sla_targets", {})
    if sla:
        config.sla_targets = {str(k): int(v) for k, v in sla.items()}
    
    # After-hours availability
    if "after_hours_available" in raw_config:
        config.after_hours_available = bool(raw_config["after_hours_available"])
    
    # Business hours
    hours = raw_config.get("business_hours", {})
    if "start" in hours:
        config.business_hours_start = int(hours["start"])
    if "end" in hours:
        config.business_hours_end = int(hours["end"])
    
    # Thresholds
    thresholds = raw_config.get("thresholds", {})
    if "fp_rate_good" in thresholds:
        config.fp_rate_threshold = float(thresholds["fp_rate_good"])
    if "automation_target" in thresholds:
        config.automation_target = float(thresholds["automation_target"])
    
    return config


def create_default_config(output_path: Path) -> None:
    """Create a default configuration file.
    
    Args:
        output_path: Path where to write the config file
    """
    default_content = '''# Client Configuration for Escalation Details Report
# Copy this file and customize for each client.

tier: "Standard Tier"
client_name_override: null

industry_benchmarks:
  mttr_minutes: 192
  mttd_minutes: 66
  incidents_per_day: 11.4

# report_period:
#   start: "2025-08-01"
#   end: "2025-08-31"

cost_modeling:
  analyst_hourly_rate: 85
  coverage_hourly_rate: 220
  breach_cost_estimate: 4200000

sla_targets:
  Critical: 30
  High: 60
  Medium: 180
  Low: 240

business_hours:
  start: 8
  end: 18

thresholds:
  fp_rate_good: 10.0
  fp_rate_warning: 15.0
  automation_target: 88.0
  manual_escalation_max: 12.0
  mttr_good: 150
  mttr_warning: 200
'''
    
    output_path = Path(output_path)
    with open(output_path, 'w') as f:
        f.write(default_content)
