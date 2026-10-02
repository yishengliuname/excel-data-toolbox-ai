"""Single source of truth for deterministic operation contracts.

The planner, capability manifest and execution gate share these definitions.
No domain logic or external provider calls belong in this module.
"""
from __future__ import annotations
from collections.abc import Mapping
from types import MappingProxyType

INPUT_COUNTS: Mapping[str, tuple[int, int]] = MappingProxyType(
    {
        "clean": (1, 1),
        "normalize_fields": (1, 1),
        "deduplicate": (1, 1),
        "partition": (1, 1),
        "derive_columns": (1, 1),
        "select_rename_sort": (1, 1),
        "concat": (2, 20),
        "join": (2, 2),
        "lookup": (2, 2),
        "summary": (1, 1),
        "split": (1, 1),
        "mask": (1, 1),
        "validate": (1, 1),
        "reconcile": (2, 2),
        "fuzzy_cluster": (1, 1),
        "fuzzy_lookup": (2, 2),
        "quality": (1, 1),
        "describe": (1, 1),
        "correlation": (1, 1),
        "outliers": (1, 1),
        "trend": (1, 1),
        "contribution": (1, 1),
        "pivot": (1, 1),
        "compare": (2, 2),
        "rfm": (1, 1),
        "recipe": (1, 1),
        "finance": (1, 1),
        "sales_management_report": (1, 1),
        "quarterly_sales_report": (2, 12),
        "inventory_management_report": (5, 12),
        "hr_management_report": (4, 20),
        # The generic compiler is schema-driven rather than case-count driven.
        # Large customer workbooks commonly contain dozens of facts, masters
        # and notes, so keep the same safe upper bound as enterprise diagnosis.
        "adaptive_analysis_report": (1, 100),
        "customer_order_analysis": (2, 100),
        "selection_recommendation_report": (1, 20),
        # Business diagnosis is capability-gated by the domain recognisers,
        # not by an arbitrary sheet count.  A validated store-period P&L may
        # be one fact sheet plus one summary, while large projects may contain
        # dozens of independent fact/master tables.
        "enterprise_diagnosis_report": (1, 100),
    }
)

PARAM_KEYS: Mapping[str, tuple[frozenset[str], frozenset[str]]] = MappingProxyType(
    {
        "normalize_fields": (frozenset({"fields"}), frozenset({"fields"})),
        "deduplicate": (frozenset({"subset"}), frozenset({"subset", "keep", "allow_conflicts"})),
        "partition": (frozenset({"rules"}), frozenset({"rules", "reason_column"})),
        "derive_columns": (frozenset({"formulas"}), frozenset({"formulas", "errors"})),
        "clean": (
            frozenset(),
            frozenset(
                {
                    "trim_whitespace",
                    "normalize_blank_strings",
                    "drop_empty_rows",
                    "drop_empty_columns",
                    "drop_duplicates",
                    "duplicate_subset",
                    "keep_duplicate",
                    "infer_types",
                    "type_inference_threshold",
                    "missing_strategy",
                    "missing_subset",
                    "drop_missing_how",
                    "fill_values",
                    "fill_numeric_with",
                    "fill_text_with",
                    "fill_boolean_with",
                    "reset_index",
                }
            ),
        ),
        "select_rename_sort": (
            frozenset(),
            frozenset({"columns", "rename", "sort_by", "ascending", "na_position", "reset_index"}),
        ),
        "concat": (
            frozenset(),
            frozenset({"join", "ignore_index", "source_column"}),
        ),
        "join": (
            frozenset(),
            frozenset({"on", "left_on", "right_on", "how", "suffixes", "validate", "allow_many_to_many", "max_output_rows"}),
        ),
        "lookup": (
            frozenset({"source_key"}),
            frozenset(
                {
                    "source_key",
                    "lookup_key",
                    "value_columns",
                    "keep_lookup_duplicate",
                    "add_match_column",
                    "match_column",
                }
            ),
        ),
        "summary": (
            frozenset({"by", "aggregations"}),
            frozenset({"by", "aggregations", "dropna", "sort"}),
        ),
        "split": (
            frozenset(),
            frozenset({"by", "rows_per_table", "drop_group_columns"}),
        ),
        "mask": (
            frozenset({"columns"}),
            frozenset({"columns", "strategy", "salt", "mask_char", "keep_start", "keep_end"}),
        ),
        "validate": (
            frozenset({"rules"}),
            frozenset({"rules", "include_values", "max_value_chars", "fail_on_error"}),
        ),
        "reconcile": (
            frozenset({"left_amount", "right_amount"}),
            frozenset(
                {
                    "left_amount",
                    "right_amount",
                    "left_date",
                    "right_date",
                    "left_key_columns",
                    "right_key_columns",
                    "left_secondary_columns",
                    "right_secondary_columns",
                    "amount_tolerance",
                    "date_tolerance_days",
                    "enable_split_candidates",
                    "max_candidates_per_row",
                    "max_candidate_pairs",
                    "max_split_combinations",
                }
            ),
        ),
        "fuzzy_cluster": (
            frozenset({"column"}),
            frozenset({"column", "threshold", "max_unique"}),
        ),
        "fuzzy_lookup": (
            frozenset({"source_key", "lookup_key", "value_columns"}),
            frozenset(
                {
                    "source_key",
                    "lookup_key",
                    "value_columns",
                    "threshold",
                    "ambiguous_gap",
                }
            ),
        ),
        "quality": (frozenset(), frozenset({"key_columns"})),
        "describe": (
            frozenset(),
            frozenset({"columns", "include_text", "percentiles"}),
        ),
        "correlation": (
            frozenset(),
            frozenset({"columns", "method", "min_periods"}),
        ),
        "outliers": (
            frozenset(),
            frozenset({"columns", "method", "iqr_multiplier", "z_threshold"}),
        ),
        "trend": (
            frozenset({"date_column", "value_columns"}),
            frozenset(
                {
                    "date_column",
                    "value_columns",
                    "frequency",
                    "aggregation",
                    "group_by",
                    "period_column",
                }
            ),
        ),
        "contribution": (
            frozenset({"category_columns", "value_column"}),
            frozenset(
                {
                    "category_columns",
                    "value_column",
                    "aggregation",
                    "pareto_threshold",
                    "top_n",
                    "include_other",
                }
            ),
        ),
        "pivot": (
            frozenset({"index", "columns"}),
            frozenset({"index", "columns", "values", "aggregation", "fill_value", "margins", "margins_name"}),
        ),
        "compare": (
            frozenset({"key_columns"}),
            frozenset({"key_columns", "compare_columns", "suffixes", "include_unchanged"}),
        ),
        "rfm": (
            frozenset({"customer_column", "date_column", "amount_column"}),
            frozenset(
                {
                    "customer_column",
                    "date_column",
                    "amount_column",
                    "transaction_column",
                    "reference_date",
                    "quantiles",
                }
            ),
        ),
        "recipe": (
            frozenset({"name", "steps"}),
            frozenset({"name", "description", "steps", "schema_version"}),
        ),
        "finance": (
            frozenset({"task", "columns"}),
            frozenset({"task", "columns", "as_of_date", "buckets", "perspective", "tolerance"}),
        ),
        "sales_management_report": (
            frozenset(
                {
                    "date_column",
                    "product_column",
                    "region_column",
                    "salesperson_column",
                    "sales_column",
                    "cost_column",
                    "satisfaction_column",
                }
            ),
            frozenset(
                {
                    "date_column",
                    "product_column",
                    "region_column",
                    "salesperson_column",
                    "sales_column",
                    "cost_column",
                    "satisfaction_column",
                    "quantity_column",
                    "satisfaction_threshold",
                }
            ),
        ),
        "quarterly_sales_report": (
            frozenset({"source_names"}),
            frozenset({"source_names", "satisfaction_threshold"}),
        ),
        "inventory_management_report": (
            frozenset({"source_names"}),
            frozenset({"source_names", "recent_days", "overstock_multiplier"}),
        ),
        "hr_management_report": (
            frozenset({"source_names"}),
            frozenset({"source_names", "expected_workdays", "excellent_score", "attention_score"}),
        ),
        "adaptive_analysis_report": (
            frozenset({"source_names"}),
            frozenset({"source_names", "user_request", "top_n", "outlier_multiplier"}),
        ),
        "customer_order_analysis": (
            frozenset({"source_names"}),
            frozenset({"source_names", "user_request", "low_margin_threshold"}),
        ),
        "selection_recommendation_report": (
            frozenset({"source_names", "top_n"}),
            frozenset({"source_names", "user_request", "top_n", "include_charts"}),
        ),
        "enterprise_diagnosis_report": (
            frozenset({"source_names"}),
            frozenset({"source_names", "user_request", "low_margin_threshold"}),
        ),
    }
)


ALLOWED_OPERATIONS = frozenset(INPUT_COUNTS)
REPORT_OPERATIONS = frozenset(name for name in ALLOWED_OPERATIONS if name.endswith("_report")) | {"customer_order_analysis"}

if ALLOWED_OPERATIONS != frozenset(PARAM_KEYS):
    raise RuntimeError("Operation contracts are incomplete")
