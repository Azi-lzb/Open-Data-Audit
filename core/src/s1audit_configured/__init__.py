"""Independent XLSX-configured audit engine using V_CUR/V_PRE/V_EXT DSL."""

from .engine import AuditEngine
from .model import (
    AuditRunResult,
    ApplicabilityScope,
    Configuration,
    EvalValue,
    Evidence,
    ExternalMetricSpec,
    MetricSpec,
    RuleDefinition,
    RuleResult,
    RuntimeContext,
    SourceBinding,
    SourceContract,
)
from .workbook_config import load_configuration
from .errors import ConfigurationError, DSLParseError, EvaluationError, RetrievalError

__all__ = [
    "AuditEngine",
    "AuditRunResult",
    "ApplicabilityScope",
    "Configuration",
    "ConfigurationError",
    "DSLParseError",
    "EvalValue",
    "Evidence",
    "EvaluationError",
    "ExternalMetricSpec",
    "MetricSpec",
    "RuleDefinition",
    "RuleResult",
    "RuntimeContext",
    "RetrievalError",
    "SourceBinding",
    "SourceContract",
    "load_configuration",
]
