"""
engine/__init__.py — 외부에서 사용하는 심볼을 한 곳에서 노출한다.

사용 예:
    from engine import audit, AuditResult, calc_material_cost, ...
"""

from engine.auditor import (
    audit,
    audit_scenario,
    AuditResult,
)

from engine.material import (
    calc_material_cost,
    load_material_master,
    get_market_prices,
)

from engine.difficulty import (
    DIFFICULTY_TABLE,
    get_coefficient,
    get_label,
    get_all_options,
    parse_level_from_option,
    auto_difficulty,
)

from engine.analyzer import (
    analyze_step_file,
    is_step_available,
)

__all__ = [
    # auditor
    "audit",
    "audit_scenario",
    "AuditResult",
    # material
    "calc_material_cost",
    "load_material_master",
    "get_market_prices",
    # difficulty
    "DIFFICULTY_TABLE",
    "get_coefficient",
    "get_label",
    "get_all_options",
    "parse_level_from_option",
    "auto_difficulty",
    # analyzer
    "analyze_step_file",
    "is_step_available",
]
