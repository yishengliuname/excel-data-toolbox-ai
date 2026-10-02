"""Reusable, explicit transformations; no eval, fuzzy guesses or row deletion.

Rejected records are separate artifacts. Domain decisions belong in the plan,
not in these operators. Missing or invalid arithmetic is never treated as zero.
"""
from __future__ import annotations

from collections.abc import Mapping
import math
import unicodedata

import numpy as np
import pandas as pd

from .recipes import RecipeStep, _condition_mask

OPERATIONS = frozenset({"normalize_fields", "deduplicate", "partition", "derive_columns"})


def _items(value, label):
    if not isinstance(value, (list, tuple)) or not value or len(value) > 100:
        raise ValueError(f"{label} must contain 1..100 items")
    return value


def _name(value):
    if not isinstance(value, str) or not value.strip() or len(value) > 256:
        raise ValueError("Column/reason must be nonempty text of at most 256 characters")


def _operand(value):
    if not isinstance(value, Mapping) or len(value) != 1:
        raise ValueError("Operand must be exactly {column: name} or {literal: number}")
    if "column" in value:
        _name(value["column"])
    elif "literal" in value:
        number = value["literal"]
        if isinstance(number, bool) or not isinstance(number, (int, float)) or not math.isfinite(number):
            raise ValueError("Arithmetic literals must be finite numbers")
    else:
        raise ValueError("Unknown operand")


def validate_params(operation: str, params: Mapping) -> None:
    if operation == "normalize_fields":
        fields = params["fields"]
        if not isinstance(fields, Mapping) or not fields or len(fields) > 100:
            raise ValueError("fields must be a nonempty mapping")
        for column, spec in fields.items():
            _name(column)
            if not isinstance(spec, Mapping) or set(spec) - {"strip", "case", "unicode", "mapping"}:
                raise ValueError("Normalization allows strip, case, unicode and exact mapping only")
            if "strip" in spec and not isinstance(spec["strip"], bool):
                raise ValueError("strip must be boolean")
            if spec.get("case", "preserve") not in {"preserve", "upper", "lower"}:
                raise ValueError("case must be preserve/upper/lower")
            if spec.get("unicode", "NFKC") not in {"NFKC", "NFC", "none"}:
                raise ValueError("unicode must be NFKC/NFC/none")
            mapping = spec.get("mapping", {})
            if not isinstance(mapping, Mapping) or len(mapping) > 1000:
                raise ValueError("mapping must be an exact text mapping")
            for key, value in mapping.items():
                _name(key)
                _name(value)
    elif operation == "deduplicate":
        for column in _items(params["subset"], "subset"):
            _name(column)
        keep = params.get("keep", "first")
        if keep is not False and keep not in {"first", "last"}:
            raise ValueError("keep must be first/last/false")
        if not isinstance(params.get("allow_conflicts", False), bool):
            raise ValueError("allow_conflicts must be boolean")
    elif operation == "partition":
        _name(params.get("reason_column", "排除原因"))
        for rule in _items(params["rules"], "rules"):
            if not isinstance(rule, Mapping) or set(rule) - {"conditions", "combine", "reason"} or "reason" not in rule:
                raise ValueError("Partition rules require conditions and reason")
            _name(rule["reason"])
            RecipeStep("filter", {key: value for key, value in rule.items() if key != "reason"})
    elif operation == "derive_columns":
        if params.get("errors", "raise") not in {"raise", "coerce"}:
            raise ValueError("errors must be raise/coerce")
        for formula in _items(params["formulas"], "formulas"):
            if not isinstance(formula, Mapping) or set(formula) - {"output", "operator", "left", "right", "round"}:
                raise ValueError("Unknown arithmetic fields")
            if not {"output", "operator", "left", "right"} <= set(formula):
                raise ValueError("Formula requires output/operator/left/right")
            _name(formula["output"])
            if formula["operator"] not in {"add", "subtract", "multiply", "divide"}:
                raise ValueError("Only add/subtract/multiply/divide are supported")
            _operand(formula["left"])
            _operand(formula["right"])
            digits = formula.get("round")
            if digits is not None and (isinstance(digits, bool) or not isinstance(digits, int) or not 0 <= digits <= 12):
                raise ValueError("round must be 0..12")


def execute(operation: str, frame: pd.DataFrame, params: Mapping):
    validate_params(operation, params)
    result = frame.copy(deep=True).reset_index(drop=True)
    if not result.columns.is_unique:
        raise ValueError("Duplicate field names must be resolved before processing")
    if operation == "normalize_fields":
        changed = {}
        for column, spec in params["fields"].items():
            series = result[column]
            if series.dropna().map(lambda value: not isinstance(value, str)).any():
                raise ValueError(f"{column}: text normalization cannot reinterpret numeric identifiers")
            def normalize(value):
                if pd.isna(value):
                    return value
                mode = spec.get("unicode", "NFKC")
                text = unicodedata.normalize(mode, value) if mode != "none" else value
                if spec.get("strip", True):
                    text = text.strip()
                case = spec.get("case", "preserve")
                text = text.upper() if case == "upper" else text.lower() if case == "lower" else text
                return spec.get("mapping", {}).get(text, text)
            result[column] = series.map(normalize)
            changed[column] = int((~result[column].eq(series).fillna(False) & ~series.isna()).sum())
        return {"primary": result}, {"changed_cells": changed}
    if operation == "deduplicate":
        subset = list(params["subset"])
        if result[subset].isna().any(axis=None) or result[subset].map(
            lambda x: isinstance(x, str) and not x.strip()
        ).any(axis=None):
            raise ValueError("Missing duplicate keys require review; cannot silently remove them")
        unique_records = result.drop_duplicates()
        if unique_records.duplicated(subset=subset, keep=False).any() and not params.get("allow_conflicts", False):
            raise ValueError("Duplicate keys have conflicting values: review before selecting a survivor")
        discarded = result.duplicated(subset=subset, keep=params.get("keep", "first"))
        return {"primary": result.loc[~discarded].reset_index(drop=True),
                "rejected": result.loc[discarded].reset_index(drop=True)}, {
                    "input_rows": len(result), "rejected_rows": int(discarded.sum()), "keys": subset}
    if operation == "partition":
        reason_column = params.get("reason_column", "排除原因")
        if reason_column in result.columns:
            raise ValueError("Reason column already exists")
        reasons = pd.Series("", index=result.index, dtype=object)
        for rule in params["rules"]:
            combine = rule.get("combine", "and")
            matched = pd.Series(combine == "and", index=result.index)
            for condition in rule["conditions"]:
                current = _condition_mask(result[condition["column"]], condition["operator"], condition.get("value"))
                matched = matched & current if combine == "and" else matched | current
            reasons.loc[matched] = reasons.loc[matched].map(lambda x: (x + "；" if x else "") + rule["reason"])
        rejected = reasons.ne("")
        excluded = result.loc[rejected].copy()
        excluded[reason_column] = reasons.loc[rejected]
        return {"primary": result.loc[~rejected].reset_index(drop=True),
                "rejected": excluded.reset_index(drop=True)}, {"rejected_rows": int(rejected.sum()), "input_rows": len(result)}
    if operation == "derive_columns":
        invalid = pd.Series(False, index=result.index)
        for formula in params["formulas"]:
            output = formula["output"]
            if output in result.columns:
                raise ValueError(f"Refusing to overwrite existing column: {output}")
            def operand(spec):
                if "literal" in spec:
                    return pd.Series(float(spec["literal"]), index=result.index)
                raw = result[spec["column"]]
                if pd.api.types.is_datetime64_any_dtype(raw.dtype) or raw.map(lambda x: isinstance(x, (bool, pd.Timestamp))).any():
                    raise ValueError("Dates and booleans are not arithmetic amounts")
                return pd.to_numeric(raw, errors="coerce")
            left, right = operand(formula["left"]), operand(formula["right"])
            operator = formula["operator"]
            with np.errstate(divide="ignore", invalid="ignore", over="ignore"):
                value = {"add": lambda: left + right, "subtract": lambda: left - right,
                         "multiply": lambda: left * right, "divide": lambda: left / right}[operator]()
            bad = ~np.isfinite(value.astype(float))
            if bad.any() and params.get("errors", "raise") == "raise":
                raise ValueError(f"{output}: {int(bad.sum())} missing/invalid operands or division by zero")
            invalid |= bad
            result[output] = value.mask(bad)
            if "round" in formula:
                result[output] = result[output].round(formula["round"])
        return {"primary": result, "review": result.loc[invalid].reset_index(drop=True)}, {
            "review_rows": int(invalid.sum()), "derived_columns": [x["output"] for x in params["formulas"]]}
    raise ValueError(f"Unsupported structured operation: {operation}")
