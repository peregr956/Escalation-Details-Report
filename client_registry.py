"""Client registry module for managing multiple client configurations.

This module provides functions to load, list, and retrieve client
configurations from a central registry file (clients/registry.yaml).
Designed to scale to hundreds of clients with varying configurations.
"""
from pathlib import Path
from typing import Dict, Any, List, Optional
import logging

logger = logging.getLogger(__name__)

# Default registry path relative to this module
REGISTRY_PATH = Path(__file__).parent / "clients" / "registry.yaml"


def load_registry(registry_path: Optional[Path] = None) -> Dict[str, Any]:
    """Load the client registry manifest.

    Args:
        registry_path: Optional path to registry file (defaults to clients/registry.yaml)

    Returns:
        Registry dictionary with 'version' and 'clients' keys

    Raises:
        FileNotFoundError: If registry file doesn't exist
        ValueError: If registry format is invalid
    """
    try:
        import yaml
    except ImportError:
        raise ImportError("PyYAML is required. Install with: pip install pyyaml")

    path = registry_path or REGISTRY_PATH

    if not path.exists():
        raise FileNotFoundError(f"Client registry not found: {path}")

    with open(path, 'r') as f:
        registry = yaml.safe_load(f)

    if not isinstance(registry, dict):
        raise ValueError(f"Invalid registry format: expected dict, got {type(registry)}")

    if 'clients' not in registry:
        raise ValueError("Registry missing 'clients' key")

    return registry


def get_client(client_id: str, registry_path: Optional[Path] = None) -> Dict[str, Any]:
    """Get client configuration by ID.

    Args:
        client_id: Client identifier (e.g., 'acme', 'burlington')
        registry_path: Optional path to registry file

    Returns:
        Client configuration dictionary containing:
        - name: Display name
        - tier: Service tier
        - config: Path to config YAML file
        - data_format: Data format profile (optional)
        - csm_name: Customer success manager name (optional)
        - csm_email: CSM email address (optional)

    Raises:
        KeyError: If client_id not found
        FileNotFoundError: If registry doesn't exist
    """
    registry = load_registry(registry_path)
    clients = registry.get('clients', {})

    if client_id not in clients:
        available = sorted(clients.keys())[:10]
        available_str = ', '.join(available)
        if len(clients) > 10:
            available_str += f", ... ({len(clients)} total)"
        raise KeyError(f"Client '{client_id}' not found. Available: {available_str}")

    return clients[client_id]


def list_clients(
    filter_tier: Optional[str] = None,
    filter_data_format: Optional[str] = None,
    registry_path: Optional[Path] = None
) -> List[str]:
    """List available client IDs with optional filtering.

    Args:
        filter_tier: Only return clients with this service tier
        filter_data_format: Only return clients with this data format
        registry_path: Optional path to registry file

    Returns:
        List of client IDs matching the filters

    Example:
        >>> list_clients(filter_tier="Signature Tier")
        ['acme', 'megacorp', 'techco']
    """
    registry = load_registry(registry_path)
    clients = registry.get('clients', {})

    result = []
    for client_id, config in clients.items():
        # Apply tier filter
        if filter_tier and config.get('tier') != filter_tier:
            continue

        # Apply data format filter
        if filter_data_format and config.get('data_format', 'standard') != filter_data_format:
            continue

        result.append(client_id)

    return sorted(result)


def get_client_config_path(client_id: str, registry_path: Optional[Path] = None) -> Path:
    """Get the full path to a client's configuration file.

    Args:
        client_id: Client identifier
        registry_path: Optional path to registry file

    Returns:
        Path to the client's YAML configuration file

    Raises:
        KeyError: If client not found
        FileNotFoundError: If config file doesn't exist
    """
    client = get_client(client_id, registry_path)
    config_path = client.get('config')

    if not config_path:
        raise ValueError(f"Client '{client_id}' has no config path defined")

    # Resolve path relative to registry directory
    registry_dir = (registry_path or REGISTRY_PATH).parent.parent
    full_path = registry_dir / config_path

    if not full_path.exists():
        raise FileNotFoundError(f"Config file not found for client '{client_id}': {full_path}")

    return full_path


def add_client(
    client_id: str,
    name: str,
    tier: str = "Standard Tier",
    config_path: Optional[str] = None,
    data_format: str = "standard",
    registry_path: Optional[Path] = None,
    **extra_fields
) -> None:
    """Add or update a client in the registry.

    Args:
        client_id: Unique client identifier
        name: Display name for the client
        tier: Service tier (e.g., "Signature Tier", "Standard Tier")
        config_path: Path to client config YAML (relative to project root)
        data_format: Data format profile ('standard', 'burlington')
        registry_path: Optional path to registry file
        **extra_fields: Additional fields (csm_name, csm_email, etc.)

    Note:
        Writes changes to the registry file immediately.
    """
    try:
        import yaml
    except ImportError:
        raise ImportError("PyYAML is required. Install with: pip install pyyaml")

    path = registry_path or REGISTRY_PATH

    # Load existing registry or create new
    if path.exists():
        registry = load_registry(path)
    else:
        registry = {'version': 1, 'clients': {}}

    # Build client entry
    client_entry = {
        'name': name,
        'tier': tier,
        'data_format': data_format,
    }

    if config_path:
        client_entry['config'] = config_path

    # Add extra fields
    client_entry.update(extra_fields)

    # Add/update client
    registry['clients'][client_id] = client_entry

    # Write back to file
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w') as f:
        yaml.dump(registry, f, default_flow_style=False, sort_keys=False)

    logger.info(f"Added/updated client '{client_id}' in registry")


def remove_client(client_id: str, registry_path: Optional[Path] = None) -> bool:
    """Remove a client from the registry.

    Args:
        client_id: Client identifier to remove
        registry_path: Optional path to registry file

    Returns:
        True if client was removed, False if not found
    """
    try:
        import yaml
    except ImportError:
        raise ImportError("PyYAML is required. Install with: pip install pyyaml")

    path = registry_path or REGISTRY_PATH
    registry = load_registry(path)

    if client_id not in registry.get('clients', {}):
        return False

    del registry['clients'][client_id]

    with open(path, 'w') as f:
        yaml.dump(registry, f, default_flow_style=False, sort_keys=False)

    logger.info(f"Removed client '{client_id}' from registry")
    return True
