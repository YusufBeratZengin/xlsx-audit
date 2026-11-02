"""Rule registry.

Order matters only for readability: the report is sorted by severity and
location before it is rendered.
"""

from __future__ import annotations

from typing import Dict, List, Sequence, Tuple, Type

from .base import Context, Rule
from .broken_refs import BrokenRefsRule
from .circular_refs import CircularRefsRule
from .external_links import ExternalLinksRule
from .hidden_sheets import HiddenSheetsRule
from .magic_numbers import MagicNumbersRule
from .range_checker import RangeCheckerRule

__all__ = [
    "ALL_RULES",
    "RULE_CLASSES",
    "Context",
    "Rule",
    "build_rules",
    "rule_names",
]

RULE_CLASSES: Tuple[Type[Rule], ...] = (
    BrokenRefsRule,
    MagicNumbersRule,
    ExternalLinksRule,
    HiddenSheetsRule,
    CircularRefsRule,
    RangeCheckerRule,
)


def rule_names() -> List[str]:
    return [cls.name for cls in RULE_CLASSES]


def build_rules(selected: Sequence[str] | None = None) -> List[Rule]:
    """Instantiate every rule, or only the ones named in ``selected``."""
    instances: Dict[str, Type[Rule]] = {cls.name: cls for cls in RULE_CLASSES}
    if selected is None:
        return [cls() for cls in RULE_CLASSES]
    unknown = [name for name in selected if name not in instances]
    if unknown:
        raise ValueError(
            f"unknown rule(s): {', '.join(sorted(unknown))}. "
            f"Available: {', '.join(sorted(instances))}"
        )
    return [instances[name]() for name in selected]


ALL_RULES = RULE_CLASSES
