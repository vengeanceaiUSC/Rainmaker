"""Phase 1 financial pipeline: SEC XBRL → Pydantic 3-statement → DCF → Excel inject."""

__all__ = [
    "edgar",
    "map_xbrl",
    "models",
    "three_statement",
    "dcf",
    "export_csv",
    "named_range_map",
    "prepare_template",
    "inject_template",
    "wire_dcf",
]
