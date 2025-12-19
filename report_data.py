from dataclasses import dataclass
from typing import Dict, Any, Optional


@dataclass
class ReportData:
    """Data class to hold all report metrics."""
    
    # Executive Summary data
    executive_summary: Dict[str, Any] = None
    
    # Value Delivered data
    value_delivered: Dict[str, Any] = None
    
    # Protection Achieved data
    protection_achieved: Dict[str, Any] = None
    
    # Threat Landscape data (optional)
    threat_landscape: Optional[Dict[str, Any]] = None
    
    # Insights and Opportunities data
    insights_opportunities: Dict[str, Any] = None
    
    # Forward Direction data
    forward_direction: Dict[str, Any] = None
    
    def __post_init__(self):
        """Initialize placeholder dictionaries if not provided."""
        if self.executive_summary is None:
            self.executive_summary = {}
        if self.value_delivered is None:
            self.value_delivered = {}
        if self.protection_achieved is None:
            self.protection_achieved = {}
        if self.insights_opportunities is None:
            self.insights_opportunities = {}
        if self.forward_direction is None:
            self.forward_direction = {}

