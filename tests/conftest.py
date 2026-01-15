"""Pytest configuration and fixtures for Escalation Details Report tests.

This module provides shared fixtures and configuration for all test modules.
"""
import sys
from pathlib import Path

# Add parent directory to path so tests can import project modules
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))
