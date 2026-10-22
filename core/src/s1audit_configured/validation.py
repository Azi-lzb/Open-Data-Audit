from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Any

from .workbook_config import load_configuration


def validate_configuration_workbook(path: str | Path) -> dict[str, Any]:
    """Load the XLSX through the same strict validation path used by the engine."""
    configuration = load_configuration(path)
    rules_by_business = Counter(rule.business for rule in configuration.rules)
    enabled_by_business = Counter(rule.business for rule in configuration.rules if rule.enabled)
    return {
        "valid": True,
        "path": str(Path(path).resolve()),
        "version": configuration.version,
        "version_status": configuration.version_status,
        "rules": len(configuration.rules),
        "enabled_rules": sum(rule.enabled for rule in configuration.rules),
        "disabled_rules": sum(not rule.enabled for rule in configuration.rules),
        "metrics": len(configuration.metrics),
        "external_metrics": len(configuration.external_metrics),
        "source_contracts": len(configuration.source_contracts),
        "applicability_scopes": len(configuration.applicability_scopes),
        "businesses": [
            {
                "name": business,
                "rules": rules_by_business[business],
                "enabled_rules": enabled_by_business[business],
            }
            for business in sorted(rules_by_business)
        ],
    }
