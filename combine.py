from __future__ import annotations
import logging
import os
import re
import sys
import time
import types
from datetime import datetime
from pathlib import Path
import openpyxl
import pandas as pd
from playwright.sync_api import sync_playwright
import argparse
from copy import copy
from dataclasses import dataclass, field
import re
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
import pandas as pd
from openpyxl import load_workbook


TARGET_COLUMNS = [
    "Product",
    "Module",
    "Feature Name",
    "Profile Options Value",
    "ESS Jobs Value",
    "Action Required",
]

CANONICAL_TEMPLATE_PATH = Path("SCM_REDWOOD_FEATURES.xlsx")
DIRECT_COLUMNS = ["Product", "Module", "Feature Name", "Action Required"]
DEFAULT_MISSING_VALUE = ""
REFERENCE_ACTION_VALUES = {
    "OPT IN",
    "OPTIN",
    "Auto",
    "Profile Options",
    "ESS Jobs",
    "OPT IN + Profile Options",
    "OPTIN + Profile Options",
    "OPT IN + ESS Jobs",
    "OPTIN + ESS Jobs",
    "Profile Options + ESS Jobs",
    "OPT IN + ESS Jobs + Profile Options",
    "OPTIN + ESS Jobs + Profile Options",
}

HEADER_HINTS: Dict[str, List[str]] = {
    "Product": ["product", "pillar", "offering"],
    "Module": ["module", "functional area", "area"],
    "Feature Name": ["feature name", "feature", "redwood feature", "capability"],
    "Action Required": ["action required", "action", "status", "opt in", "opt-in"],
}

PROFILE_TASK_HINTS = ["task name", "task", "setup task", "offering task"]
PROFILE_OPTION_HINTS = [
    "profile option",
    "profile options",
    "profile option code",
    "profile option name",
    "option name",
]
PROFILE_VALUE_HINTS = [
    "profile option value",
    "profile value",
    "option value",
    "value",
]

ESS_PROCESS_HINTS = ["process name", "process", "ess job", "job name", "job"]
ESS_PARAMETER_HINTS = ["parameter label", "parameter name", "parameter"]
ESS_VALUE_HINTS = ["parameter value", "ess value", "value"]
ADDITIONAL_INFO_HINTS = [
    "additional info",
    "additional information",
    "additional details",
    "implementation details",
    "setup details",
    "setup steps",
    "comments",
    "notes",
]

ESS_PROCESS_LABELS = [
    "ess job name",
    "ess job",
    "scheduled process name",
    "scheduled process",
    "process name",
    "process",
    "job name",
    "job",
]
ESS_PARAMETER_LABELS = [
    "parameter label",
    "parameter name",
    "paramter label",
    "paramter name",
    "parametr label",
    "parametr name",
    "parameter",
    "paramter",
    "parametr",
]
ESS_PARAMETER_VALUE_LABELS = [
    "parameter value",
    "paramter value",
    "parametr value",
    "value",
]
ESS_STOP_WORDS = {
    "run",
    "execute",
    "submit",
    "schedule",
    "scheduled",
    "process",
    "ess",
    "job",
    "with",
    "using",
    "the",
    "following",
    "below",
    "as",
    "is",
    "are",
    "should",
    "must",
    "needs",
    "need",
    "to",
    "be",
    "set",
    "enter",
    "provide",
    "provided",
}


@dataclass(frozen=True)
class ReferencePatterns:
    ess_processes: Tuple[str, ...] = ()
    ess_parameters: Tuple[str, ...] = ()
    action_values: Tuple[str, ...] = tuple(sorted(REFERENCE_ACTION_VALUES))
    rows_by_full_key: Dict[Tuple[str, str, str], Dict[str, str]] = field(default_factory=dict)
    rows_by_product_feature_key: Dict[Tuple[str, str], Dict[str, str]] = field(default_factory=dict)
    rows_by_feature_key: Dict[str, Dict[str, str]] = field(default_factory=dict)


def normalize_header(value: object) -> str:
    """Return a lowercase, punctuation-insensitive header key."""
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()


def clean_cell(value: object, missing_value: str = DEFAULT_MISSING_VALUE) -> str:
    if pd.isna(value):
        return missing_value

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "null"}:
        return missing_value
    return text


def clean_series(
    series: pd.Series,
    missing_value: str = DEFAULT_MISSING_VALUE,
) -> pd.Series:
    return series.map(lambda value: clean_cell(value, missing_value))


def empty_series(index: pd.Index, missing_value: str = DEFAULT_MISSING_VALUE) -> pd.Series:
    return pd.Series([missing_value] * len(index), index=index, dtype="object")


def header_score(target: str, candidate: str) -> float:
    target_norm = normalize_header(target)
    candidate_norm = normalize_header(candidate)

    if target_norm == candidate_norm:
        return 1.0
    if target_norm and target_norm in candidate_norm:
        return 0.92
    if candidate_norm and candidate_norm in target_norm:
        return 0.88

    return SequenceMatcher(None, target_norm, candidate_norm).ratio()


def find_column(
    columns: Iterable[object],
    target: str,
    hints: Optional[Sequence[str]] = None,
    threshold: float = 0.72,
) -> Optional[str]:
    """Find the closest matching column name for a target and optional hints."""
    column_names = [str(column) for column in columns]
    candidates = [target, *(hints or [])]

    best: Tuple[float, Optional[str]] = (0.0, None)
    for column in column_names:
        for candidate in candidates:
            score = header_score(candidate, column)
            if score > best[0]:
                best = (score, column)

    if best[0] >= threshold:
        return best[1]
    return None


def find_exact_output_column(columns: Iterable[object], target: str) -> Optional[str]:
    target_norm = normalize_header(target)
    for column in columns:
        if normalize_header(column) == target_norm:
            return str(column)
    return None


def first_present_series(
    frame: pd.DataFrame,
    columns: Sequence[Optional[str]],
    missing_value: str = DEFAULT_MISSING_VALUE,
) -> pd.Series:
    for column in columns:
        if column and column in frame.columns:
            values = clean_series(frame[column], missing_value)
            if values.ne(missing_value).any():
                return values

    return empty_series(frame.index, missing_value)


def has_real_value(value: object, missing_value: str = DEFAULT_MISSING_VALUE) -> bool:
    return clean_cell(value, missing_value) != missing_value


def compact_text(value: object) -> str:
    text = clean_cell(value, missing_value="")
    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\s*\n+\s*", " | ", text)
    return text.strip(" ;,|")


def clean_extracted_piece(value: str) -> str:
    text = compact_text(value)
    text = re.sub(r"^[\s:=-]+", "", text)
    text = re.sub(r"[\s.;,]+$", "", text)
    text = re.sub(r"\s*>\s*", " > ", text)
    words = text.split()

    while words and words[0].strip(":=-").lower() in ESS_STOP_WORDS:
        words.pop(0)

    return " ".join(words).strip(" ;,")


def clean_pattern_piece(value: str) -> str:
    text = compact_text(value)
    text = re.sub(r"^[\s:=-]+", "", text)
    text = re.sub(r"[\s.;,]+$", "", text)
    text = re.sub(r"\s*>\s*", " > ", text)
    return re.sub(r"\s+", " ", text).strip(" ;,")


ESS_ACRONYM_REPLACEMENTS = {
    "ess": "ESS",
    "oscs": "OSCS",
}

ESS_FIXED_PARAMETER_LABELS = {
    "ess job to create index definition and perform initial ingest to oscs": "Index Name To Reingest",
    "ess jobs to create index definition and perform initial ingest to oscs": "Index Name To Reingest",
}


def preserve_ess_acronyms(value: str) -> str:
    text = str(value or "")
    for word, replacement in ESS_ACRONYM_REPLACEMENTS.items():
        text = re.sub(rf"\b{word}\b", replacement, text, flags=re.IGNORECASE)
    text = re.sub(r"\bESS\s+jobs?\b", "ESS job", text, flags=re.IGNORECASE)
    text = re.sub(
        r"(^|,\s*|\|\s*)jobs\s+to\b",
        lambda match: f"{match.group(1)}ESS job To",
        text,
        flags=re.IGNORECASE,
    )
    return text


def _ess_process_key(value: str) -> str:
    text = preserve_ess_acronyms(clean_pattern_piece(value))
    text = re.sub(r"\bESS\s+jobs?\b", "ESS job", text, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", text).strip().casefold()


def normalize_ess_parameter_label_for_process(process: str, parameter: str) -> str:
    fixed_label = ESS_FIXED_PARAMETER_LABELS.get(_ess_process_key(process))
    if fixed_label:
        return fixed_label
    return clean_pattern_piece(parameter)


def format_ess_entry(process: str, parameter: str, value: str) -> str:
    process = preserve_ess_acronyms(clean_pattern_piece(process))
    parameter = normalize_ess_parameter_label_for_process(process, parameter)
    value = clean_pattern_piece(value)
    return preserve_ess_acronyms(f"{process} > {parameter} > {value}")


def _looks_like_rest_api_guidance(*parts: str) -> bool:
    text = normalize_header(" ".join(str(part or "") for part in parts))
    rest_terms = [
        "rest api guides",
        "rest service definition",
        "oracle help center",
        "apis schema",
        "quick start section",
    ]
    return any(term in text for term in rest_terms)


def is_valid_ess_entry(process: str, parameter: str, value: str) -> bool:
    process = clean_pattern_piece(process)
    parameter = clean_pattern_piece(parameter)
    value = clean_pattern_piece(value)
    if not (process and parameter and value):
        return False

    if _looks_like_rest_api_guidance(process, parameter, value):
        return False

    process_norm = normalize_header(process)
    parameter_norm = normalize_header(parameter)
    value_norm = normalize_header(value)

    if process_norm.startswith(("review the ", "role with ", "buyers who ", "users with ")):
        return False
    if "service area of interest" in parameter_norm:
        return False
    if "quick start" in value_norm:
        return False

    return True


def label_pattern(labels: Sequence[str]) -> str:
    return "|".join(re.escape(label) for label in sorted(labels, key=len, reverse=True))


def extract_labeled_value(
    text: str,
    labels: Sequence[str],
    stop_labels: Sequence[str],
) -> Optional[str]:
    label_expr = label_pattern(labels)
    stop_expr = label_pattern(stop_labels)
    pattern = re.compile(
        rf"\b(?:{label_expr})\b\s*(?:is|as|=|:|-)?\s*"
        rf"(.+?)(?=(?:\s*(?:,|;|\||\.|\n|\band\b|\bwith\b)\s*)?"
        rf"\b(?:{stop_expr})\b\s*(?:is|as|=|:|-)?|$)",
        flags=re.IGNORECASE,
    )
    match = pattern.search(text)
    if not match:
        return None

    value = clean_extracted_piece(match.group(1))
    return value or None


def normalize_existing_ess_pattern(text: str) -> Optional[str]:
    compacted = compact_text(text)
    if ">" not in compacted:
        return None

    entries: List[str] = []
    for raw_entry in split_existing_pattern_entries(compacted):
        if ">" not in raw_entry:
            continue

        pieces = [clean_pattern_piece(piece) for piece in raw_entry.split(">")]
        pieces = [piece for piece in pieces if piece]
        original_piece_count = len(pieces)
        while len(pieces) >= 3:
            if is_valid_ess_entry(pieces[0], pieces[1], pieces[2]):
                entries.append(format_ess_entry(pieces[0], pieces[1], pieces[2]))
            pieces = pieces[3:]
        if original_piece_count == 1:
            entries.append(pieces[0])

    return join_ess_entries(entries) if entries else None


def split_existing_pattern_entries(text: str) -> List[str]:
    return [
        entry
        for entry in re.split(r"\s*(?:\||\n|,(?=[^,>]+>))\s*", text)
        if clean_pattern_piece(entry)
    ]


def split_parameter_values(value: str) -> List[str]:
    values = [
        clean_pattern_piece(part)
        for part in re.split(r"\s*(?:,|\||\n|;|\band\b)\s*", value)
    ]
    return [part for part in values if part]


def join_pattern_entries(entries: Iterable[str]) -> str:
    output: List[str] = []
    seen: set[str] = set()
    for entry in entries:
        pieces = [clean_pattern_piece(piece) for piece in str(entry).split(">")]
        pieces = [piece for piece in pieces if piece]
        cleaned = " > ".join(pieces) if len(pieces) >= 2 else clean_pattern_piece(entry)
        key = cleaned.lower()
        if cleaned and key not in seen:
            output.append(cleaned)
            seen.add(key)
    return ", ".join(output)


def join_ess_entries(entries: Iterable[str]) -> str:
    return join_pattern_entries(entries)


PROFILE_CODE_RE = re.compile(r"^[A-Z][A-Z0-9]*(?:_[A-Z0-9]+)+$", re.IGNORECASE)
PROFILE_CODE_EXCLUDED_RE = re.compile(r"^(?:F\d+|WAVE\d+|[A-Z]+(?:\d+[A-Z]?)+)$", re.IGNORECASE)
PROFILE_CODE_EXCLUDED_SUFFIX_RE = re.compile(r"_(?:JOB|DUTY|PWA|PRIV)$", re.IGNORECASE)


def is_valid_profile_option_code(code: str) -> bool:
    normalized = clean_pattern_piece(code).strip().upper()
    if not normalized or normalized == "_":
        return False
    if not PROFILE_CODE_RE.fullmatch(normalized):
        return False
    if PROFILE_CODE_EXCLUDED_RE.fullmatch(normalized):
        return False
    if PROFILE_CODE_EXCLUDED_SUFFIX_RE.search(normalized):
        return False
    return True


def normalize_profile_option_code(code: str) -> str:
    return clean_pattern_piece(code).strip().upper()


def format_profile_option_entry(task: str, code: str, value: str) -> str:
    task = clean_pattern_piece(task) or "Manage Administrator Profile Values"
    code = normalize_profile_option_code(code)
    value = clean_pattern_piece(value) or "Yes"
    return f"{task} > {code} = {value}"


def normalize_profile_options_cell(
    value: object,
    missing_value: str = DEFAULT_MISSING_VALUE,
) -> str:
    text = clean_cell(value, missing_value)
    if text == missing_value:
        return missing_value

    if ">" not in text:
        return _format_direct_profile_value_for_combine(text) or missing_value

    entries = []
    for raw_entry in split_existing_pattern_entries(text):
        normalized_entry = raw_entry.strip().strip('"').strip("'")
        match = re.match(
            r"^(?P<task>[^>]+?)\s*>\s*(?P<code>[A-Z][A-Z0-9_]*[A-Z0-9])\s*=\s*(?P<value>[^,\n\r]+)$",
            normalized_entry,
            flags=re.IGNORECASE,
        )
        if not match:
            continue

        code = match.group("code")
        if not is_valid_profile_option_code(code):
            continue
        entries.append(
            format_profile_option_entry(
                match.group("task"),
                code,
                match.group("value"),
            )
        )

    return join_pattern_entries(entries) or missing_value


def _format_direct_profile_value_for_combine(value: str) -> str:
    if pd.isna(value):
        return ""
    text = _norm(str(value))
    if not text:
        return ""
    if ">" in text:
        normalized = normalize_profile_options_cell(text, "")
        return "" if normalized == DEFAULT_MISSING_VALUE else normalized

    candidate_codes = re.findall(r"\b[A-Z][A-Z0-9]*(?:_[A-Z0-9]+)+\b", text, flags=re.I)

    entries = []
    seen = set()
    for token in candidate_codes:
        if not is_valid_profile_option_code(token):
            continue
        code = normalize_profile_option_code(token)
        if code in seen:
            continue
        seen.add(code)
        entries.append(format_profile_option_entry("Manage Administrator Profile Values", code, "Y"))
    return ", ".join(entries)


def extract_index_name_pairs(text: str) -> List[Tuple[str, str]]:
    pairs: List[Tuple[str, str]] = []
    for line in re.split(r"[\n|]+", str(text or "")):
        if ":" not in line:
            continue

        label, raw_values = line.split(":", 1)
        label = clean_pattern_piece(label)
        if not label or normalize_header(label) in {"index name", "this index name"}:
            continue

        for value in split_parameter_values(raw_values):
            pairs.append((label, value))

    if pairs:
        return pairs

    pair_pattern = re.compile(
        r"([A-Za-z][A-Za-z0-9 /&()._-]{1,120}?)\s*:\s*"
        r"([A-Za-z0-9][A-Za-z0-9_.-]*(?:\s*(?:,|;|\band\b)\s*"
        r"[A-Za-z0-9][A-Za-z0-9_.-]*)*)",
        flags=re.IGNORECASE,
    )

    for match in pair_pattern.finditer(text):
        label = clean_pattern_piece(match.group(1))
        if normalize_header(label) in {"index name", "this index name"}:
            continue
        for value in split_parameter_values(match.group(2)):
            pairs.append((label, value))

    return pairs


def extract_scheduled_process_index_pattern(text: str) -> Optional[str]:
    process_pattern = re.compile(
        r"(?:\brun\s+(?:the\s+)?)?"
        r"((?:the\s+)?ess\s+jobs?\s+to\s+.+?)\s+scheduled\s+process\b",
        flags=re.IGNORECASE,
    )
    process_matches = list(process_pattern.finditer(text))
    if not process_matches:
        return None

    entries: List[str] = []
    for index, process_match in enumerate(process_matches):
        process = preserve_ess_acronyms(clean_pattern_piece(process_match.group(1)))
        process = re.sub(r"^the\s+", "", process, flags=re.IGNORECASE)

        end = process_matches[index + 1].start() if index + 1 < len(process_matches) else len(text)
        remainder = text[process_match.end() : end]
        index_name_match = re.search(
            r"\bwith\s+(?:this\s+)?index\s+name\s*:?",
            remainder,
            flags=re.IGNORECASE,
        )
        if index_name_match:
            remainder = remainder[index_name_match.end() :]

        pairs = extract_index_name_pairs(remainder)
        if not pairs:
            entries.append(process)
            continue

        entries.extend(format_ess_entry(process, label, value) for label, value in pairs)

    return join_ess_entries(entries) if entries else None


def parse_reference_ess_value(value: object) -> Tuple[List[str], List[str]]:
    text = compact_text(value)
    if not text:
        return [], []

    processes: List[str] = []
    parameters: List[str] = []
    for raw_entry in re.split(r"\s*\|\s*", text):
        parts = [preserve_ess_acronyms(clean_extracted_piece(part)) for part in raw_entry.split(">")]
        parts = [part for part in parts if part]
        if parts:
            processes.append(parts[0])
        if len(parts) >= 2:
            parameters.append(parts[1])

    return processes, parameters


def unique_sorted(values: Iterable[str]) -> Tuple[str, ...]:
    cleaned = {clean_extracted_piece(value) for value in values}
    cleaned.discard("")
    return tuple(sorted(cleaned, key=lambda value: (-len(value), value.lower())))


def reference_key(value: object) -> str:
    return normalize_header(value)


def reference_full_key(product: object, module: object, feature_name: object) -> Tuple[str, str, str]:
    return (reference_key(product), reference_key(module), reference_key(feature_name))


def reference_product_feature_key(product: object, feature_name: object) -> Tuple[str, str]:
    return (reference_key(product), reference_key(feature_name))


def reference_row_score(row: Dict[str, str]) -> Tuple[int, int, int, int]:
    actions = normalize_action_text(row.get("Action Required", ""))
    return (
        1 if has_real_value(row.get("ESS Jobs Value", ""), "") else 0,
        1 if has_real_value(row.get("Profile Options Value", ""), "") else 0,
        len(actions),
        sum(1 for column in TARGET_COLUMNS if has_real_value(row.get(column, ""), "")),
    )


def choose_best_reference_row(
    current: Optional[Dict[str, str]],
    candidate: Dict[str, str],
) -> Dict[str, str]:
    if current is None:
        return candidate
    if reference_row_score(candidate) > reference_row_score(current):
        return candidate
    return current


def load_reference_patterns(
    template_path: Path = CANONICAL_TEMPLATE_PATH,
) -> ReferencePatterns:
    if not template_path.exists():
        return ReferencePatterns()

    try:
        reference = pd.read_excel(
            template_path,
            sheet_name="Filtered Features",
            keep_default_na=False,
        )
    except Exception:
        return ReferencePatterns()

    processes: List[str] = []
    parameters: List[str] = []
    rows_by_full_key: Dict[Tuple[str, str, str], Dict[str, str]] = {}
    rows_by_product_feature_key: Dict[Tuple[str, str], Dict[str, str]] = {}
    feature_key_counts: Dict[str, int] = {}
    rows_by_feature_key: Dict[str, Dict[str, str]] = {}

    for _, row in reference.iterrows():
        canonical_row = {
            column: "" if pd.isna(row.get(column, "")) else str(row.get(column, ""))
            for column in TARGET_COLUMNS
        }
        feature_key = reference_key(canonical_row["Feature Name"])
        full_key = reference_full_key(
            canonical_row["Product"],
            canonical_row["Module"],
            canonical_row["Feature Name"],
        )
        product_feature_key = reference_product_feature_key(
            canonical_row["Product"],
            canonical_row["Feature Name"],
        )
        if feature_key:
            feature_key_counts[feature_key] = feature_key_counts.get(feature_key, 0) + 1
            rows_by_feature_key[feature_key] = canonical_row
        if all(product_feature_key):
            rows_by_product_feature_key[product_feature_key] = choose_best_reference_row(
                rows_by_product_feature_key.get(product_feature_key),
                canonical_row,
            )
        if all(full_key):
            rows_by_full_key[full_key] = canonical_row

    rows_by_feature_key = {
        key: row
        for key, row in rows_by_feature_key.items()
        if feature_key_counts.get(key) == 1
    }

    if "ESS Jobs Value" in reference.columns:
        for value in reference["ESS Jobs Value"]:
            value_processes, value_parameters = parse_reference_ess_value(value)
            processes.extend(value_processes)
            parameters.extend(value_parameters)

    action_values = set(REFERENCE_ACTION_VALUES)
    if "Action Required" in reference.columns:
        for value in reference["Action Required"]:
            cleaned = clean_cell(value, missing_value="")
            if cleaned:
                action_values.add(cleaned)

    return ReferencePatterns(
        ess_processes=unique_sorted(processes),
        ess_parameters=unique_sorted(parameters),
        action_values=unique_sorted(action_values),
        rows_by_full_key=rows_by_full_key,
        rows_by_product_feature_key=rows_by_product_feature_key,
        rows_by_feature_key=rows_by_feature_key,
    )


def find_reference_feature_row(
    reference_patterns: Optional[ReferencePatterns],
    product: object,
    module: object,
    feature_name: object,
) -> Optional[Dict[str, str]]:
    if not reference_patterns:
        return None

    full_key = reference_full_key(product, module, feature_name)
    if all(full_key) and full_key in reference_patterns.rows_by_full_key:
        return reference_patterns.rows_by_full_key[full_key]

    product_feature_key = reference_product_feature_key(product, feature_name)
    if all(product_feature_key) and product_feature_key in reference_patterns.rows_by_product_feature_key:
        return reference_patterns.rows_by_product_feature_key[product_feature_key]

    feature_key = reference_key(feature_name)
    if feature_key:
        return reference_patterns.rows_by_feature_key.get(feature_key)
    return None


def find_known_phrase(text: str, phrases: Sequence[str]) -> Optional[str]:
    text_norm = re.sub(r"\s+", " ", text).strip()
    for phrase in phrases:
        pattern = re.compile(
            r"(?<![A-Za-z0-9])" + re.escape(phrase) + r"(?![A-Za-z0-9])",
            flags=re.IGNORECASE,
        )
        match = pattern.search(text_norm)
        if match:
            return preserve_ess_acronyms(clean_extracted_piece(match.group(0)))
    return None


def extract_value_after_phrase(text: str, phrase: str) -> Optional[str]:
    pattern = re.compile(re.escape(phrase), flags=re.IGNORECASE)
    match = pattern.search(text)
    if not match:
        return None

    remainder = text[match.end() :]
    remainder = re.sub(
        r"^\s*(?:is|as|=|:|-|>|to|value|values|of|for|,|;)+\s*",
        "",
        remainder,
        flags=re.IGNORECASE,
    ).strip()
    match = re.match(r"([A-Za-z0-9][A-Za-z0-9_.-]*(?:\s*,\s*[A-Za-z0-9][A-Za-z0-9_.-]*)*)", remainder)
    if not match:
        return None

    value = clean_extracted_piece(match.group(1))
    return value or None


def build_reference_based_ess(
    text: str,
    reference_patterns: ReferencePatterns,
) -> Optional[str]:
    process = find_known_phrase(text, reference_patterns.ess_processes)
    parameter = find_known_phrase(text, reference_patterns.ess_parameters)
    if not process:
        return None

    if not parameter:
        return process

    value = extract_value_after_phrase(text, parameter)
    if not value:
        return process

    values = [clean_extracted_piece(part) for part in re.split(r"\s*,\s*", value)]
    values = [part for part in values if part]
    if not values:
        return process

    return join_ess_entries(format_ess_entry(process, parameter, part) for part in values)


def extract_ess_job_from_text(
    value: object,
    reference_patterns: Optional[ReferencePatterns] = None,
) -> Optional[str]:
    text = compact_text(value)
    if not text:
        return None

    scheduled_process_pattern = extract_scheduled_process_index_pattern(text)
    if scheduled_process_pattern:
        return scheduled_process_pattern

    if reference_patterns:
        reference_based = build_reference_based_ess(text, reference_patterns)
        if reference_based:
            return reference_based

    existing_pattern = normalize_existing_ess_pattern(text)
    if existing_pattern:
        return preserve_ess_acronyms(existing_pattern)

    stop_labels = [
        *ESS_PROCESS_LABELS,
        *ESS_PARAMETER_LABELS,
        *ESS_PARAMETER_VALUE_LABELS,
    ]
    process_value = extract_labeled_value(text, ESS_PROCESS_LABELS, stop_labels)
    parameter_value = extract_labeled_value(text, ESS_PARAMETER_LABELS, stop_labels)
    value_value = extract_labeled_value(text, ESS_PARAMETER_VALUE_LABELS, stop_labels)

    if process_value and parameter_value and value_value:
        return format_ess_entry(process_value, parameter_value, value_value)
    if process_value:
        return preserve_ess_acronyms(process_value)

    return None


def normalize_ess_cell(value: object, missing_value: str = DEFAULT_MISSING_VALUE) -> str:
    text = clean_cell(value, missing_value)
    if text == missing_value:
        return missing_value

    normalized_pattern = normalize_existing_ess_pattern(text)
    if ">" in text:
        return preserve_ess_acronyms(normalized_pattern) if normalized_pattern else missing_value

    return preserve_ess_acronyms(text)


def build_ess_from_additional_info_series(
    frame: pd.DataFrame,
    missing_value: str = DEFAULT_MISSING_VALUE,
    reference_patterns: Optional[ReferencePatterns] = None,
) -> pd.Series:
    info_col = find_column(
        frame.columns,
        "Additional Info",
        ADDITIONAL_INFO_HINTS,
        threshold=0.70,
    )
    if not info_col:
        return empty_series(frame.index, missing_value)

    extracted = frame[info_col].map(
        lambda value: extract_ess_job_from_text(value, reference_patterns)
    )
    return extracted.map(lambda value: value if value else missing_value)


def build_profile_options_series(
    frame: pd.DataFrame,
    missing_value: str = DEFAULT_MISSING_VALUE,
) -> pd.Series:
    direct = find_exact_output_column(frame.columns, "Profile Options Value")
    if direct:
        return frame[direct].map(
            lambda value: _format_direct_profile_value_for_combine(value) or missing_value
        )

    task_col = find_column(frame.columns, "Task Name", PROFILE_TASK_HINTS, threshold=0.70)
    option_col = find_column(
        frame.columns,
        "Profile Option",
        PROFILE_OPTION_HINTS,
        threshold=0.70,
    )
    value_col = find_column(
        frame.columns,
        "Profile Option Value",
        PROFILE_VALUE_HINTS,
        threshold=0.80,
    )

    if not option_col and not value_col:
        return empty_series(frame.index, missing_value)

    task_values = first_present_series(frame, [task_col], missing_value)
    option_values = first_present_series(frame, [option_col, value_col], missing_value)
    profile_value_values = first_present_series(frame, [value_col], missing_value)

    def build(row_index: object) -> str:
        option_value = option_values.loc[row_index]
        if option_value == missing_value:
            return missing_value

        task_value = task_values.loc[row_index]
        if task_value == missing_value:
            task_value = "Profile options"

        row_profile_value = profile_value_values.loc[row_index]
        if row_profile_value == missing_value:
            row_profile_value = "Yes"

        option_entries: List[str] = []
        for single_value in split_parameter_values(option_value):
            profile_code = single_value
            profile_value = row_profile_value
            if "=" in single_value:
                profile_code, profile_value = single_value.split("=", 1)
                profile_code = clean_pattern_piece(profile_code)
                profile_value = clean_pattern_piece(profile_value) or "Yes"

            if is_valid_profile_option_code(profile_code):
                option_entries.append(
                    format_profile_option_entry(task_value, profile_code, profile_value)
                )

        return join_pattern_entries(option_entries) or missing_value

    return pd.Series([build(index) for index in frame.index], index=frame.index)


def build_ess_jobs_series(
    frame: pd.DataFrame,
    missing_value: str = DEFAULT_MISSING_VALUE,
    reference_patterns: Optional[ReferencePatterns] = None,
) -> pd.Series:
    additional_info_values = build_ess_from_additional_info_series(
        frame,
        missing_value,
        reference_patterns,
    )

    direct = find_exact_output_column(frame.columns, "ESS Jobs Value")
    if direct:
        direct_values = frame[direct].map(
            lambda value: normalize_ess_cell(value, missing_value)
        )
        return direct_values.mask(direct_values.eq(missing_value), additional_info_values)

    process_col = find_column(frame.columns, "Process Name", ESS_PROCESS_HINTS, threshold=0.70)
    parameter_col = find_column(
        frame.columns,
        "Parameter Label",
        ESS_PARAMETER_HINTS,
        threshold=0.76,
    )
    value_col = find_column(frame.columns, "Parameter Value", ESS_VALUE_HINTS, threshold=0.80)

    if not process_col:
        return additional_info_values

    process_values = first_present_series(frame, [process_col], missing_value)
    parameter_values = first_present_series(frame, [parameter_col], missing_value)
    value_values = first_present_series(frame, [value_col], missing_value)

    def build(row_index: object) -> str:
        process_value = process_values.loc[row_index]
        if process_value == missing_value:
            return missing_value

        parameter_value = parameter_values.loc[row_index]
        value_value = value_values.loc[row_index]
        if parameter_value == missing_value or value_value == missing_value:
            return preserve_ess_acronyms(process_value)

        return format_ess_entry(process_value, parameter_value, value_value)

    column_values = pd.Series([build(index) for index in frame.index], index=frame.index)
    return column_values.mask(column_values.eq(missing_value), additional_info_values)


def normalize_action_text(value: object) -> set[str]:
    text = normalize_header(value)
    actions: set[str] = set()

    if re.search(r"\bauto(?:matic|matically)?\b", text):
        actions.add("Auto")
    if re.search(r"\bopt\s*in\b|\boptin\b|\benable\b|\benabled\b", text):
        actions.add("OPT IN")
    if re.search(r"\bprofile\s+options?\b|\bprofile\s+values?\b", text):
        actions.add("Profile Options")
    if re.search(r"\bess\b|\bscheduled\s+process\b|\bjob\b", text):
        actions.add("ESS Jobs")

    return actions


def canonical_reference_action(value: object) -> Optional[str]:
    text = clean_cell(value, missing_value="")
    if not text:
        return None

    normalized = re.sub(r"\s*\+\s*", " + ", text.strip())
    for action in REFERENCE_ACTION_VALUES:
        if normalized.lower() == action.lower():
            return format_action_required(normalize_action_text(action), DEFAULT_MISSING_VALUE)

    return None


def canonicalize_action_required(value: object, missing_value: str = DEFAULT_MISSING_VALUE) -> str:
    text = clean_cell(value, missing_value="")
    if not text:
        return missing_value

    canonical = canonical_reference_action(text)
    if canonical:
        return canonical

    text_norm = normalize_header(text)
    has_auto = re.search(r"\bauto(?:matic|matically)?\b", text_norm) is not None
    has_non_auto = re.search(r"\bopt\s*in\b|\boptin\b|\bprofile\b|\bess\b|\bjob\b", text_norm) is not None
    if has_auto and not has_non_auto:
        return "Auto"

    actions = normalize_action_text(text)
    return format_action_required(actions, missing_value) if actions else (text or missing_value)


def is_auto_action_label(value: object) -> bool:
    text = normalize_header(value)
    if not text:
        return False

    non_auto_tokens = ["opt", "profile", "ess", "job"]
    return text in {"auto", "auto enabled", "automatic", "automatically", "already enabled"} or (
        "auto" in text and not any(token in text for token in non_auto_tokens)
    )


def action_required_from_normalized_values(
    source_action: object,
    profile_value: object,
    ess_value: object,
    missing_value: str = DEFAULT_MISSING_VALUE,
) -> str:
    existing_action = clean_cell(source_action, missing_value="")
    if is_auto_action_label(existing_action):
        return existing_action or "Auto"

    actions = normalize_action_text(existing_action)
    has_profile = has_real_value(profile_value, missing_value)
    has_ess = has_real_value(ess_value, missing_value)

    if has_profile:
        actions.add("Profile Options")
    if has_ess:
        actions.add("ESS Jobs")

    return format_action_required(actions, missing_value) if actions else "OPTIN"


def format_action_required(actions: set[str], missing_value: str) -> str:
    if not actions:
        return "Auto"
    if "Auto" in actions and not {"Profile Options", "ESS Jobs"}.intersection(actions):
        return "Auto"

    if {"ESS Jobs", "Profile Options"}.issubset(actions):
        return "OPTIN + ESS Jobs + Profile Options"
    if "Profile Options" in actions:
        return "OPTIN + Profile Options"
    if "ESS Jobs" in actions:
        return "OPTIN + ESS Jobs"
    if "OPT IN" in actions or "OPTIN" in actions:
        return "OPTIN"
    return "Auto"


def build_action_required_series(
    source: pd.DataFrame,
    normalized: pd.DataFrame,
    missing_value: str = DEFAULT_MISSING_VALUE,
) -> pd.Series:
    exact_column = find_exact_output_column(source.columns, "Action Required")
    fuzzy_column = find_column(source.columns, "Action Required", HEADER_HINTS.get("Action Required"))
    action_values = first_present_series(source, [exact_column, fuzzy_column], missing_value)

    output: List[str] = []
    for row_index in source.index:
        output.append(
            action_required_from_normalized_values(
                action_values.loc[row_index],
                normalized.loc[row_index, "Profile Options Value"],
                normalized.loc[row_index, "ESS Jobs Value"],
                missing_value,
            )
        )

    return pd.Series(output, index=source.index)


def normalize_frame(
    source: pd.DataFrame,
    missing_value: str = DEFAULT_MISSING_VALUE,
    reference_patterns: Optional[ReferencePatterns] = None,
) -> pd.DataFrame:
    normalized = pd.DataFrame(index=source.index)

    for column in ["Product", "Module", "Feature Name"]:
        exact_column = find_exact_output_column(source.columns, column)
        fuzzy_column = find_column(source.columns, column, HEADER_HINTS.get(column))
        normalized[column] = first_present_series(
            source,
            [exact_column, fuzzy_column],
            missing_value,
        )

    normalized["Profile Options Value"] = build_profile_options_series(source, missing_value)
    normalized["ESS Jobs Value"] = build_ess_jobs_series(
        source,
        missing_value,
        reference_patterns,
    )
    normalized["Action Required"] = build_action_required_series(
        source,
        normalized,
        missing_value,
    )

    return normalized[TARGET_COLUMNS].fillna(missing_value)


def sheet_score(frame: pd.DataFrame) -> int:
    score = 0
    for column in TARGET_COLUMNS:
        if find_exact_output_column(frame.columns, column):
            score += 3
    for column in DIRECT_COLUMNS:
        if find_column(frame.columns, column, HEADER_HINTS.get(column)):
            score += 1
    if find_column(frame.columns, "Additional Info", ADDITIONAL_INFO_HINTS, threshold=0.70):
        score += 2
    return score


def choose_source_sheet(
    sheets: Dict[str, pd.DataFrame],
    requested_sheet: Optional[str] = None,
) -> Tuple[str, pd.DataFrame]:
    if requested_sheet:
        if requested_sheet not in sheets:
            available = ", ".join(sheets)
            raise ValueError(
                f"Sheet '{requested_sheet}' was not found. Available sheets: {available}"
            )
        return requested_sheet, sheets[requested_sheet]

    ranked = sorted(
        sheets.items(),
        key=lambda item: (sheet_score(item[1]), item[0] == "Filtered Features"),
        reverse=True,
    )
    if not ranked or sheet_score(ranked[0][1]) == 0:
        raise ValueError("No usable source sheet was found in the workbook.")

    return ranked[0]


def select_source_sheets(
    sheets: Dict[str, pd.DataFrame],
    requested_sheet: Optional[str] = None,
) -> List[Tuple[str, pd.DataFrame]]:
    if requested_sheet:
        return [choose_source_sheet(sheets, requested_sheet)]

    selected = [
        (sheet_name, frame)
        for sheet_name, frame in sheets.items()
        if sheet_score(frame) > 0
    ]
    if not selected:
        raise ValueError("No usable source sheet was found in the workbook.")

    return selected


def apply_excel_visible_rows(
    input_path: Path,
    sheets: Dict[str, pd.DataFrame],
) -> Dict[str, pd.DataFrame]:
    """Honor rows hidden by Excel filters before selecting a source sheet."""
    try:
        workbook = load_workbook(input_path, read_only=False, data_only=True)
    except Exception:
        return sheets

    visible_sheets: Dict[str, pd.DataFrame] = {}
    for sheet_name, frame in sheets.items():
        worksheet = workbook[sheet_name] if sheet_name in workbook.sheetnames else None
        if worksheet is None or frame.empty:
            visible_sheets[sheet_name] = frame
            continue

        hidden_indices = [
            excel_row - 2
            for excel_row in range(2, worksheet.max_row + 1)
            if worksheet.row_dimensions[excel_row].hidden
        ]
        hidden_indices = [index for index in hidden_indices if index in frame.index]
        if hidden_indices:
            hidden_excel_rows = [index + 2 for index in hidden_indices]
            log.info(
                "Excel visible-row filter: sheet='%s' | before=%d | hidden_rows=%d | after=%d | excel_rows=%s",
                sheet_name,
                len(frame),
                len(hidden_indices),
                len(frame) - len(hidden_indices),
                hidden_excel_rows,
            )
            visible_sheets[sheet_name] = frame.drop(index=hidden_indices).reset_index(drop=True)
        else:
            log.info(
                "Excel visible-row filter: sheet='%s' | before=%d | hidden_rows=0 | after=%d",
                sheet_name,
                len(frame),
                len(frame),
            )
            visible_sheets[sheet_name] = frame

    return visible_sheets


def default_input_path() -> Path:
    workbooks = sorted(Path.cwd().glob("*.xlsx"))
    if len(workbooks) == 1:
        return workbooks[0]
    if CANONICAL_TEMPLATE_PATH.exists():
        return CANONICAL_TEMPLATE_PATH

    raise ValueError("Provide an input .xlsx file path.")


def output_path_for(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_normalized.xlsx")


def copy_cell_style(source_cell: object, target_cell: object) -> None:
    if not source_cell.has_style:
        return

    target_cell.font = copy(source_cell.font)
    target_cell.fill = copy(source_cell.fill)
    target_cell.border = copy(source_cell.border)
    target_cell.alignment = copy(source_cell.alignment)
    target_cell.number_format = source_cell.number_format
    target_cell.protection = copy(source_cell.protection)


def apply_reference_formatting(
    output_path: Path,
    template_path: Path = CANONICAL_TEMPLATE_PATH,
) -> None:
    if not template_path.exists():
        return

    template_workbook = load_workbook(template_path)
    if "Filtered Features" not in template_workbook.sheetnames:
        return

    output_workbook = load_workbook(output_path)
    output_sheet = output_workbook["Filtered Features"]
    template_sheet = template_workbook["Filtered Features"]

    for column_index in range(1, len(TARGET_COLUMNS) + 1):
        column_letter = output_sheet.cell(row=1, column=column_index).column_letter
        template_dimension = template_sheet.column_dimensions[column_letter]
        if template_dimension.width:
            output_sheet.column_dimensions[column_letter].width = template_dimension.width

        copy_cell_style(
            template_sheet.cell(row=1, column=column_index),
            output_sheet.cell(row=1, column=column_index),
        )
        if template_sheet.max_row >= 2:
            for row_index in range(2, output_sheet.max_row + 1):
                copy_cell_style(
                    template_sheet.cell(row=2, column=column_index),
                    output_sheet.cell(row=row_index, column=column_index),
                )

    output_sheet.auto_filter.ref = (
        f"A1:{output_sheet.cell(row=1, column=len(TARGET_COLUMNS)).column_letter}"
        f"{output_sheet.max_row}"
    )
    output_sheet.freeze_panes = template_sheet.freeze_panes
    output_workbook.save(output_path)


def normalize_file(
    input_path: Path,
    output_path: Optional[Path] = None,
    sheet: Optional[str] = None,
    missing_value: str = DEFAULT_MISSING_VALUE,
    template_path: Path = CANONICAL_TEMPLATE_PATH,
) -> Tuple[Path, str, int]:
    if not input_path.exists():
        raise FileNotFoundError(f"Input file does not exist: {input_path}")

    sheets = pd.read_excel(input_path, sheet_name=None)
    sheets = apply_excel_visible_rows(input_path, sheets)
    source_sheets = select_source_sheets(sheets, sheet)
    reference_patterns = load_reference_patterns(template_path)
    normalized_frames = [
        normalize_frame(source_frame, missing_value, reference_patterns)
        for _, source_frame in source_sheets
    ]
    normalized = pd.concat(normalized_frames, ignore_index=True)
    source_sheet_names = ", ".join(sheet_name for sheet_name, _ in source_sheets)

    final_output_path = output_path or output_path_for(input_path)
    with pd.ExcelWriter(final_output_path, engine="openpyxl") as writer:
        normalized.to_excel(writer, sheet_name="Filtered Features", index=False)

    apply_reference_formatting(final_output_path, template_path)

    return final_output_path, source_sheet_names, len(normalized)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Normalize an Excel workbook into a Filtered Features sheet."
    )
    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        help="Input .xlsx workbook. Defaults to the only .xlsx file in this folder.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output .xlsx workbook. Defaults to <input>_normalized.xlsx.",
    )
    parser.add_argument(
        "-s",
        "--sheet",
        help="Source sheet name. Defaults to the best matching sheet.",
    )
    parser.add_argument(
        "--missing-value",
        default=DEFAULT_MISSING_VALUE,
        help="Value used for missing fields. Defaults to blank.",
    )
    return parser.parse_args()


def normalizer_cli_main() -> None:
    args = parse_args()
    input_path = args.input or default_input_path()
    output_path, source_sheet, row_count = normalize_file(
        input_path=input_path,
        output_path=args.output,
        sheet=args.sheet,
        missing_value=args.missing_value,
    )
    print(f"Source sheet: {source_sheet}")
    print(f"Rows written: {row_count}")
    print(f"Output file: {output_path}")


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  [%(levelname)s]  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("combine")

EXCEL_PATH = os.getenv("EXCEL_PATH", "Halifax_RW_Features List JLReview_Final_041726.xlsx")


def _parse_execution_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--run-only", dest="run_only")
    args, _ = parser.parse_known_args()
    return args


_EXECUTION_ARGS = _parse_execution_args()
RUN_ONLY_RAW = _EXECUTION_ARGS.run_only if _EXECUTION_ARGS.run_only is not None else os.getenv("RUN_ONLY", "")
RUN_ONLY = _norm(RUN_ONLY_RAW).casefold() if "_norm" in globals() else str(RUN_ONLY_RAW or "").strip().casefold()


def _normalize_run_only(value: str) -> str:
    normalized = " ".join(str(value or "").replace("_", " ").replace("-", " ").split()).casefold()
    aliases = {
        "": "all",
        "all": "all",
        "opt in": "optin",
        "optin": "optin",
        "opt-in": "optin",
        "profile options": "profile",
        "profile option": "profile",
        "profile": "profile",
        "ess jobs": "ess",
        "ess job": "ess",
        "ess": "ess",
        "auto": "auto",
    }
    if normalized not in aliases:
        allowed = "All, Opt In, Profile Options, ESS Jobs, Auto"
        raise ValueError(f"Unsupported RUN_ONLY value '{value}'. Allowed values: {allowed}")
    return aliases[normalized]


RUN_MODE = _normalize_run_only(RUN_ONLY)


def _should_run(section: str) -> bool:
    return RUN_MODE == "all" or RUN_MODE == section


def _is_auto_only_mode() -> bool:
    """Return True when execution should run only Auto logging flow."""
    auto_only_flag = os.getenv("AUTO_ONLY", "").strip().casefold() in {"1", "true", "yes", "y"}
    return auto_only_flag or RUN_MODE == "auto"


def _env_int(name: str, default: int) -> int:
    try:
        return int(str(os.getenv(name, default)).strip())
    except Exception:
        return default


def _env_float(name: str, default: float) -> float:
    try:
        return float(str(os.getenv(name, default)).strip())
    except Exception:
        return default


SLOW_INSTANCE_MODE = os.getenv("SLOW_INSTANCE_MODE", "").strip().casefold() in {"1", "true", "yes", "y"}
BROWSER_HEADLESS = os.getenv("PLAYWRIGHT_HEADLESS", "false").strip().casefold() in {"1", "true", "yes", "y"}
BROWSER_SLOW_MO_MS = _env_int("PLAYWRIGHT_SLOW_MO_MS", 0 if BROWSER_HEADLESS else 200)
NAV_TIMEOUT_MS = _env_int("NAV_TIMEOUT_MS", 30_000 if SLOW_INSTANCE_MODE else 15_000)
SHORT_WAIT_MS = _env_int("SHORT_WAIT_MS", 10_000 if SLOW_INSTANCE_MODE else 5_000)
UI_STABILIZE_SEC = _env_float("UI_STABILIZE_SEC", 2.0 if SLOW_INSTANCE_MODE else 1.0)
NAV_RETRIES = max(1, _env_int("NAV_RETRIES", 4 if SLOW_INSTANCE_MODE else 2))
BASE_URL = os.getenv(
    "FUSION_BASE_URL",
    "",
)
USERNAME = os.getenv("FUSION_USERNAME", "")
PASSWORD = os.getenv("FUSION_PASSWORD", "")


def wait_for_oracle_busy_clear(page, timeout: int = 25_000) -> None:
    """Wait for common Oracle/ADF busy indicators to clear."""
    try:
        page.wait_for_function(
            "() => !document.querySelector("
            "  '[role=\"progressbar\"], .xbusy, [aria-busy=\"true\"], .oj-progress'"
            ")",
            timeout=timeout,
        )
    except Exception:
        pass


def login(
    page,
    base_url: str = BASE_URL,
    username: str = USERNAME,
    password: str = PASSWORD,
    timeout: int = 90_000,
    wait_busy_func=wait_for_oracle_busy_clear,
    logger: Optional[logging.Logger] = None,
) -> bool:
    """
    Log in to Oracle Fusion using the login form variants already handled in combine.py.

    Returns True when the Navigator is visible, or raises if no supported login path reaches it.
    """
    active_log = logger or log
    active_log.info("Login")
    page.goto(base_url, timeout=60_000)
    page.wait_for_load_state("domcontentloaded")
    time.sleep(2.0)

    idcs_password_selectors = [
        'input[id="idcs-signin-basic-signin-form-password|input"]',
        'input[class="oj-inputpassword-input oj-text-field-input oj-component-initnode|input"]',
        "#idcs-signin-basic-signin-form-password\\|input",
        "#idcs-signin-basic-signin-form-password|input",
    ]

    login_selectors = [
        (
            'input[id="idcs-signin-basic-signin-form-username|input"]',
            idcs_password_selectors,
            'button:has-text("Next"), button:has-text("Sign In")',
        ),
        (
            "#idcs-signin-basic-signin-form-username",
            idcs_password_selectors,
            '.oj-button-text:has-text("Sign In")',
        ),
        ("input#userid", "input#password", "button#btnActive"),
        ("#userid", "#password", "#btnActive"),
        ('input[name="username"]', 'input[name="password"]', 'button[type="submit"]'),
        ('input[type="email"]', 'input[type="password"]', 'button[type="submit"]'),
    ]

    attempted_login = False
    for username_selector, password_selectors, submit_selector in login_selectors:
        try:
            page.wait_for_selector(username_selector, timeout=8_000)
            if not page.is_visible(username_selector):
                continue
            page.fill(username_selector, username)

            if isinstance(password_selectors, str):
                password_selectors = [password_selectors]

            password_filled = False
            for password_selector in password_selectors:
                try:
                    page.wait_for_selector(password_selector, timeout=3_000)
                    if not page.is_visible(password_selector):
                        continue
                    page.fill(password_selector, password)
                    password_filled = True
                    break
                except Exception:
                    continue

            if not password_filled:
                continue

            page.click(submit_selector)
            active_log.info("  Login submitted via: %s", username_selector)
            attempted_login = True
            break
        except Exception:
            continue

    if not attempted_login:
        try:
            inputs = page.query_selector_all('input:not([type="hidden"])')
            if len(inputs) >= 2:
                inputs[0].fill(username)
                inputs[1].fill(password)
                inputs[1].press("Enter")
                active_log.info("  Login submitted via fallback input detection")
                attempted_login = True
        except Exception:
            pass

    page.wait_for_selector('svg[aria-label="Navigator"], [aria-label="Navigator"]', timeout=timeout)
    wait_busy_func(page)
    active_log.info("  Logged in")
    return True


def _init_run_logs():
    run_stamp = datetime.now().strftime("run_%Y%m%d_%H%M%S")
    run_dir = Path.cwd() / "logs" / run_stamp
    run_dir.mkdir(parents=True, exist_ok=True)

    log_files = {
        "optin": run_dir / "optin_log.txt",
        "profile": run_dir / "profile_options_log.txt",
        "ess": run_dir / "ess_jobs_log.txt",
        "auto": run_dir / "auto_log.txt",
        "feature": run_dir / "feature_wise_log.txt",
        "summary": run_dir / "summary_log.txt",
    }

    for key, p in log_files.items():
        with open(p, "w", encoding="utf-8") as f:
            f.write(f"{key.upper()} LOG\n")
            f.write(f"Run Folder: {run_dir}\n")
            f.write(f"Started At: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 100 + "\n")

    return run_dir, log_files


def _append_realtime(log_path: Path, message: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{ts}] {message}\n")
        f.flush()


def _safe_screenshot_token(value: object, fallback: str = "item", max_len: int = 70) -> str:
    token = re.sub(r"[^A-Za-z0-9]+", "_", str(value or "").strip())
    token = token.strip("_")
    return (token or fallback)[:max_len]


def _status_token(value: object, default: str = "info") -> str:
    text = str(value or default).strip().casefold()
    if text in {"already enabled", "already_enabled"}:
        return "already_enabled"
    if text in {"enabled now", "enabled", "pass", "success", "set"}:
        return "success"
    if text in {"not found", "not_found", "notfound"}:
        return "not_found"
    if text in {"skipped", "skip"}:
        return "skipped"
    if "fail" in text or "error" in text:
        return "error"
    return _safe_screenshot_token(text, default)


def _capture_workflow_screenshot(
    page,
    screenshots_dir: Path,
    category: str,
    step: str,
    index: int,
    item_name: str,
    status: str = "info",
):
    screenshots_dir.mkdir(parents=True, exist_ok=True)
    safe_item = _safe_screenshot_token(item_name)
    safe_status = _status_token(status)
    safe_step = _safe_screenshot_token(step)
    path = screenshots_dir / f"{index:03d}_{category}_{safe_step}_{safe_item}_{safe_status}.png"
    try:
        page.screenshot(path=str(path), full_page=True)
        log.info("Screenshot captured: %s", path)
    except Exception as exc:
        log.warning("Screenshot capture failed for %s: %s", item_name, exc)

SUPPORTED_ACTION_TOKENS = {
    "opt in": "opt_in",
    "optin": "opt_in",
    "ess jobs": "ess_jobs",
    "profile options": "profile_options",
    "auto": "auto",
}

OPTIN_SOURCE = '"""\nOracle SCM – New Features OPT IN Automation\n============================================\nFlow:\n  1. Login\n  2. Setup & Maintenance\n  3. Actions → Go to Offerings\n  4. Click Procurement tile → New Features\n  5. For each feature from Excel (OPT IN only):\n       a. Search feature name in the search box\n       b. Read the Functional Area from the result row\n       c. Check the Enabled column:\n            - a[title="Enabled"] img[src*="accept_ena.png"] → already enabled ✅\n            - a[title="Allows Opt-In"] clickable arrow      → needs enabling\n       d. If needs enabling:\n            - Click the Allows Opt-In arrow\n            - On Edit Features page: find the feature row → tick label.x17j\n            - Click Done → back to New Features\n  6. Generate summary: feature | functional area | status | error\n\nCONFIRMED HTML (New Features page row):\n  Functional Area: 2nd <td> text = "Procure Goods from Preferred Sources During Catalog Shopping"\n  Feature:         3rd <td> text = "Procure Goods from Preferred Sources using Shopping Lists"\n  Enabled col:     <a title="Enabled"> with img[src*="accept_ena.png"] = already on\n  Opt-In col:      <a title="Allows Opt-In"> clickable arrow = needs enabling\n\n  After clicking Allows Opt-In arrow → Edit Features page:\n    Feature row has label.x17j checkbox (same page as parent feature enabling)\n\nUsage:\n    pip install playwright openpyxl python-dotenv\n    playwright install chromium\n    python new_features_optin.py\n"""\n\nimport logging\nimport os\nimport time\nfrom dataclasses import dataclass, field\nfrom datetime import datetime\nfrom pathlib import Path\n\nimport openpyxl\nfrom dotenv import load_dotenv\nfrom playwright.sync_api import sync_playwright, Page, TimeoutError as PWTimeout\n\nload_dotenv()\n\n# ─────────────────────────────────────────────────────────────────────────────\n# CONFIGURATION\n# ─────────────────────────────────────────────────────────────────────────────\n\nBASE_URL   = os.getenv("FUSION_BASE_URL", "https://eghw-dev1.fa.em2.oraclecloud.com/fscmUI/faces/FuseTaskListManagerTop?_afrLoop=60355488926401771&_adf.ctrl-state=xxa77g8uf_498")\nUSERNAME   = os.getenv("FUSION_USERNAME",  "veera.h@oracle.com")\nPASSWORD   = os.getenv("FUSION_PASSWORD",  "Welcome_01")\nEXCEL_PATH = os.getenv("EXCEL_PATH",       "SCM_REDWOOD_FEATURES.xlsx")\nOFFERING   = "Procurement"\n\nHEADLESS = False\nSLOW_MO  = 200\nTIMEOUT  = 45_000\n\n# ─────────────────────────────────────────────────────────────────────────────\n# LOGGING & SCREENSHOTS\n# ─────────────────────────────────────────────────────────────────────────────\n\nlogging.basicConfig(\n    level=logging.INFO,\n    format="%(asctime)s  [%(levelname)s]  %(message)s",\n    datefmt="%H:%M:%S",\n)\nlog = logging.getLogger(__name__)\n\nBASE_OUTPUT_DIR = Path.home() / "Desktop" / "fusion-screenshots"\nBASE_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)\nRUN_OUTPUT_DIR = BASE_OUTPUT_DIR\n\n\ndef init_run_artifacts() -> Path:\n    """Create a per-run folder and attach a file logger inside it."""\n    global RUN_OUTPUT_DIR\n\n    run_ts = datetime.now().strftime("run_%Y%m%d_%H%M%S")\n    RUN_OUTPUT_DIR = BASE_OUTPUT_DIR / run_ts\n    RUN_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)\n\n    log_path = RUN_OUTPUT_DIR / "run.log"\n    fh = logging.FileHandler(log_path, encoding="utf-8")\n    fh.setLevel(logging.INFO)\n    fh.setFormatter(logging.Formatter("%(asctime)s  [%(levelname)s]  %(message)s", "%H:%M:%S"))\n    log.addHandler(fh)\n\n    log.info(f"📁 Run artifacts folder: {RUN_OUTPUT_DIR}")\n    log.info(f"📝 Log file: {log_path}")\n    return RUN_OUTPUT_DIR\n\n\ndef snap(page: Page, label: str):\n    ts = datetime.now().strftime("%H%M%S")\n    safe = "".join(c if c.isalnum() else "_" for c in label)[:40]\n    path = RUN_OUTPUT_DIR / f"{ts}_{safe}.png"\n    try:\n        page.screenshot(path=path, full_page=True)\n        log.info(f"  📸 {path.name}")\n    except Exception:\n        pass\n\n\ndef wait_busy(page: Page, timeout: int = 25_000):\n    try:\n        page.wait_for_function(\n            "() => !document.querySelector("\n            "  \'[role=\\"progressbar\\"], .xbusy, [aria-busy=\\"true\\"], .oj-progress\'"\n            ")",\n            timeout=timeout,\n        )\n    except PWTimeout:\n        pass\n\n\ndef mdown(el):\n    """Bypass ADF onclick=\'return false\' with mousedown sequence."""\n    el.evaluate("""el => {\n        [\'mousedown\',\'mouseup\',\'click\'].forEach(n =>\n            el.dispatchEvent(new MouseEvent(n, {\n                bubbles:true, cancelable:true, view:window,\n                button:0, buttons:1\n            }))\n        );\n    }""")\n\n\n# ─────────────────────────────────────────────────────────────────────────────\n# DATA\n# ─────────────────────────────────────────────────────────────────────────────\n\n@dataclass\nclass Feature:\n    name: str\n    module: str       # from Excel — maps to Functional Area on New Features page\n    row: int\n\n\n@dataclass\nclass FeatureResult:\n    name: str\n    functional_area: str   # read from the page (may differ from Excel module)\n    status: str            # "Already Enabled" | "Enabled Now" | "Not Found" | "Error"\n    error: str = ""\n\n\n# ─────────────────────────────────────────────────────────────────────────────\n# EXCEL READER\n# ─────────────────────────────────────────────────────────────────────────────\n\ndef load_optin_features(path: str) -> list[Feature]:\n    """\n    Read workbook and return Procurement features with exact Action Required = OPT IN.\n    Source columns (by header name): Product, Module, Feature Name, Action Required\n    """\n    log.info(f"📂 Loading from: {path}")\n    wb = openpyxl.load_workbook(path, data_only=True)\n\n    # Prefer the raw source sheet that includes Action Required.\n    source_sheet = "Filtered Features" if "Filtered Features" in wb.sheetnames else wb.sheetnames[0]\n    ws = wb[source_sheet]\n    features: list[Feature] = []\n\n    # Build header index map from first row\n    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())\n    header_map = {str(v).strip().casefold(): idx for idx, v in enumerate(header_row) if v is not None}\n\n    required_headers = ["product", "module", "feature name", "action required"]\n    missing = [h for h in required_headers if h not in header_map]\n    if missing:\n        log.error(\n            f"Missing required columns in sheet \'{source_sheet}\': {missing}. "\n            f"Found: {[str(h) for h in header_row if h is not None]}"\n        )\n        return features\n\n    product_i = header_map["product"]\n    module_i = header_map["module"]\n    feature_i = header_map["feature name"]\n    action_i = header_map["action required"]\n\n    seen: set[tuple[str, str]] = set()\n\n    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):\n        if not any(v for v in row if v is not None and str(v).strip()):\n            continue\n\n        product = str(row[product_i]).strip() if len(row) > product_i and row[product_i] else ""\n        module = str(row[module_i]).strip() if len(row) > module_i and row[module_i] else ""\n        name = str(row[feature_i]).strip() if len(row) > feature_i and row[feature_i] else ""\n        action = str(row[action_i]).strip() if len(row) > action_i and row[action_i] else ""\n\n        # Only Procurement + exact OPT IN (no combined values)\n        if product.casefold() != "procurement":\n            continue\n        if action.casefold() != "opt in":\n            continue\n\n        if not name or not module:\n            continue\n\n        dedupe_key = (module.casefold(), name.casefold())\n        if dedupe_key in seen:\n            continue\n        seen.add(dedupe_key)\n\n        features.append(Feature(name=name, module=module, row=row_idx))\n\n    log.info(f"  ✅ {len(features)} Procurement features loaded (Action Required = OPT IN, exact)")\n    return features\n\n\n# ─────────────────────────────────────────────────────────────────────────────\n# NAVIGATION: Login → Setup & Maintenance → Offerings → New Features\n# ─────────────────────────────────────────────────────────────────────────────\n\ndef login(page: Page):\n    log.info("🔐 Login")\n    page.goto(BASE_URL, timeout=60_000)\n    page.wait_for_load_state("domcontentloaded")\n    time.sleep(2.0)\n\n    for u_sel, p_sel, s_sel in [\n        (\'input[id="idcs-signin-basic-signin-form-username|input"]\',\n         \'input[class="idcs-signin-basic-signin-form-password|input"]\',\n         \'button:has-text("Next"), button:has-text("Sign In")\'),\n        ("#idcs-signin-basic-signin-form-username",\n         "#idcs-signin-basic-signin-form-password|input",\n         ".oj-button-text:has-text(\'Sign In\')"),\n        ("input#userid", "input#password", "button#btnActive"),\n        ("#userid",      "#password",      "#btnActive"),\n        (\'input[name="username"]\', \'input[name="password"]\', \'button[type="submit"]\'),\n    ]:\n        try:\n            page.wait_for_selector(u_sel, timeout=8_000)\n            if not page.is_visible(u_sel):\n                continue\n            page.fill(u_sel, USERNAME)\n            page.fill(p_sel, PASSWORD)\n            page.click(s_sel)\n            log.info(f"  ✓ Login via: {u_sel}")\n            break\n        except Exception:\n            continue\n\n    page.wait_for_selector(\n        \'svg[aria-label="Navigator"], [aria-label="Navigator"]\', timeout=90_000\n    )\n    wait_busy(page)\n    log.info("  ✅ Logged in")\n\n\ndef go_to_setup_and_maintenance(page: Page):\n    log.info("🔧 Setup and Maintenance")\n    for sel in [\'#pt1\\\\:_UIScmil2u\', \'img[title="Settings and Actions"]\',\n                \'img[alt="Settings and Actions"]\']:\n        try:\n            page.click(sel, timeout=8_000)\n            break\n        except Exception:\n            continue\n    time.sleep(0.8)\n    page.click(\'text=Setup and Maintenance\', timeout=10_000)\n    wait_busy(page)\n\n\ndef go_to_offerings(page: Page):\n    log.info("📋 Go to Offerings")\n    for sel in [\'td.xo2:has-text("Actions")\', \'button:has-text("Actions")\',\n                \'a:has-text("Actions")\', \':text-is("Actions")\']:\n        try:\n            page.wait_for_selector(sel, state="visible", timeout=8_000)\n            page.click(sel)\n            break\n        except Exception:\n            continue\n    time.sleep(0.7)\n    for sel in [\'td.xo2:has-text("Go to Offerings")\', \'a:has-text("Go to Offerings")\',\n                \':text-is("Go to Offerings")\']:\n        try:\n            page.wait_for_selector(sel, state="visible", timeout=8_000)\n            page.click(sel)\n            break\n        except Exception:\n            continue\n    wait_busy(page)\n    log.info("  ✅ Offerings loaded")\n\n\ndef go_to_new_features(page: Page):\n    """\n    Click the Procurement tile then click New Features.\n    """\n    log.info(f"🆕 {OFFERING} → New Features")\n\n    # Click Procurement tile\n    for sel in [f\'img[title="{OFFERING}"]\', f\'img[alt="{OFFERING}"]\',\n                f\':text-is("{OFFERING}")\']:\n        try:\n            page.wait_for_selector(sel, state="visible", timeout=10_000)\n            page.click(sel)\n            log.info(f"  ✓ Tile clicked ({sel})")\n            break\n        except Exception:\n            continue\n    time.sleep(1.0)\n\n    # Click New Features\n    for sel in [f\'tr:has-text("{OFFERING}") td.xo2:has-text("New Features")\',\n                f\'tr:has-text("{OFFERING}") a:has-text("New Features")\',\n                \'td.xo2:has-text("New Features")\',\n                \'a:has-text("New Features")\',\n                \':text-is("New Features")\']:\n        try:\n            page.wait_for_selector(sel, state="visible", timeout=8_000)\n            page.click(sel)\n            log.info(f"  ✓ New Features clicked ({sel})")\n            break\n        except Exception:\n            continue\n\n    wait_busy(page)\n\n    # Ensure we are on the "Available Features" tab inside New Features.\n    for sel in [\n        \'div[id*="availableFeatureTab::ti"] a:has-text("Available Features")\',\n        \'a[id*="availableFeatureTab::disAcr"]:has-text("Available Features")\',\n        \'a:has-text("Available Features")\',\n    ]:\n        try:\n            page.wait_for_selector(sel, state="visible", timeout=6_000)\n            tab = page.locator(sel).first\n            mdown(tab)\n            log.info(f"  ✓ Available Features tab clicked ({sel})")\n            break\n        except Exception:\n            continue\n\n    # Give ADF tab switch time to render correct grid before searching.\n    try:\n        page.wait_for_selector(\n            \'div[id*="availableFeatureTab::ti"].p_AFSelected, \'\n            \'a[id*="availableFeatureTab::disAcr"].p_AFSelected\',\n            state="visible", timeout=8_000,\n        )\n    except Exception:\n        pass\n    time.sleep(1.5)\n\n    wait_busy(page)\n    # Wait for the New Features search box to confirm page loaded\n    page.wait_for_selector(\'input[aria-label="Feature"], input[name*="qbeFeature"]\',\n                           state="visible", timeout=TIMEOUT)\n    log.info("  ✅ New Features page loaded")\n\n\n# ─────────────────────────────────────────────────────────────────────────────\n# NEW FEATURES PAGE: Search + Read row status\n# ─────────────────────────────────────────────────────────────────────────────\n\ndef search_feature(page: Page, feature_name: str):\n    """Target the New Features search box reliably, type feature name, press Enter."""\n    log.info(f"  🔍 Search: \'{feature_name[:60]}\'")\n    search_box = None\n\n    # Prefer specific qbeFeature selectors (avoid generic text inputs).\n    for sel in [\n        \'input[aria-label="Feature"]\',\n        \'input[name*="qbeFeature"]\',\n        \'input[id*="qbeFeature"]\',\n    ]:\n        try:\n            page.wait_for_selector(sel, state="visible", timeout=8_000)\n            candidate = page.locator(sel).first\n            candidate.scroll_into_view_if_needed(timeout=2_000)\n            candidate.click(timeout=2_000)\n            search_box = candidate\n            log.info(f"  ✓ Search box targeted ({sel})")\n            break\n        except Exception:\n            continue\n\n    # Last-resort fallback (kept for compatibility with older page variants).\n    if not search_box:\n        fallback = \'input.x25[type="text"]\'\n        page.wait_for_selector(fallback, state="visible", timeout=8_000)\n        search_box = page.locator(fallback).first\n        search_box.click(timeout=2_000)\n        log.info(f"  ✓ Search box targeted ({fallback})")\n\n    # Clear + type + verify value landed in the intended input.\n    search_box.fill("")\n    search_box.type(feature_name, delay=30)\n    current_val = search_box.input_value()\n    if current_val.strip() != feature_name.strip():\n        search_box.click()\n        search_box.press("ControlOrMeta+A")\n        search_box.press("Backspace")\n        search_box.fill(feature_name)\n\n    page.keyboard.press("Enter")\n    time.sleep(2.0)\n    wait_busy(page)\n\n\ndef read_matching_feature_rows(page: Page, feature_name: str) -> list[dict]:\n    """\n    After searching, read all matching result rows and return a list of:\n      {\n        "functional_area": str,   # 2nd td text\n        "feature":         str,   # 3rd td text\n        "is_enabled":      bool,  # True if a[title="Enabled"] present\n        "can_opt_in":      bool,  # informational only (not used for decision)\n        "row_index":       int,   # _afrrk attribute value for targeting\n      }\n    Returns [] if no matching row found.\n\n    CONFIRMED HTML:\n      Enabled:     <a title="Enabled">   img[src*="accept_ena.png"]\n      Allows OI:   <a title="Allows Opt-In"> clickable arrow (has onclick)\n                OR <img title="Allows Opt-In"> static indicator (not clickable)\n    """\n    # Use a normalized matcher to avoid misses due to case/punctuation/spacing.\n    target = feature_name.strip().lower()\n    try:\n        result = page.evaluate("""(featName) => {\n            const norm = (s) => (s || \'\')\n                .toLowerCase()\n                .replace(/\\\\s+/g, \' \')\n                .trim();\n\n            const featNorm = norm(featName);\n            // Restrict to actual ADF data rows for the features grid.\n            const rows = Array.from(document.querySelectorAll(\'tr[_afrrk], tr[_afrowkey]\')).filter(r =>\n                r.querySelectorAll(\'td\').length >= 3\n            );\n\n            const matches = [];\n\n            for (const row of rows) {\n                const tds = row.querySelectorAll(\'td\');\n                if (tds.length < 3) continue;\n\n                const cellText = (cell) => {\n                    if (!cell) return \'\';\n                    const compactRepeatedText = (text) => {\n                        const words = (text || \'\').replace(/\\s+/g, \' \').trim().split(\' \').filter(Boolean);\n                        if (words.length && words.length % 2 === 0) {\n                            const half = words.length / 2;\n                            const left = words.slice(0, half).join(\' \').toLowerCase();\n                            const right = words.slice(half).join(\' \').toLowerCase();\n                            if (left === right) return words.slice(0, half).join(\' \');\n                        }\n                        return (text || \'\').replace(/\\s+/g, \' \').trim();\n                    };\n                    const uniqueParts = [];\n                    const seenParts = new Set();\n                    const addPart = (value) => {\n                        const cleaned = compactRepeatedText(value);\n                        const key = cleaned.toLowerCase();\n                        if (cleaned && !seenParts.has(key)) {\n                            seenParts.add(key);\n                            uniqueParts.push(cleaned);\n                        }\n                    };\n                    const parts = [\n                        cell.innerText,\n                        cell.textContent,\n                        cell.getAttribute(\'title\'),\n                        cell.getAttribute(\'aria-label\'),\n                    ];\n                    cell.querySelectorAll(\'[title], [aria-label]\').forEach(el => {\n                        parts.push(el.getAttribute(\'title\'));\n                        parts.push(el.getAttribute(\'aria-label\'));\n                    });\n                    parts.forEach(addPart);\n                    return compactRepeatedText(uniqueParts.join(\' \'));\n                };\n\n                const cellTexts = Array.from(tds).map(cellText);\n                const isTextMatch = (text) => {\n                    const textNorm = norm(text);\n                    if (!textNorm) return false;\n                    return textNorm === featNorm ||\n                        textNorm.includes(featNorm) ||\n                        featNorm.includes(textNorm);\n                };\n\n                // Locate the actual Feature column by searched feature text. Some ADF\n                // grids include an icon/control cell before Offering, so fixed indexes\n                // can shift and make td[1] point to Offering instead of Functional Area.\n                const featureIndex = cellTexts.findIndex(isTextMatch);\n                if (featureIndex < 0) continue;\n\n                const featureText = cellTexts[featureIndex];\n                const featureNorm = norm(featureText);\n\n                // Robust match (case-insensitive + partial both directions)\n                const matched = isTextMatch(featureText);\n                if (!matched) continue;\n\n                // Functional Area is the column immediately before Feature in this grid.\n                // Walk backward to skip any empty/icon-only cells.\n                let functionalAreaIndex = featureIndex - 1;\n                while (functionalAreaIndex >= 0 && !norm(cellTexts[functionalAreaIndex])) {\n                    functionalAreaIndex--;\n                }\n                const funcArea = functionalAreaIndex >= 0 ? cellTexts[functionalAreaIndex] : \'\';\n\n                // Prefer column-aware checks to avoid cross-column false positives.\n                // Typical order:\n                //  Offering | Functional Area | Feature | ... | Enabled | Allows Opt-In\n                const enabledTd = tds[featureIndex + 4] || tds[6] || null;\n                const allowsTd  = tds[featureIndex + 5] || tds[7] || null;\n\n                // Enabled should ONLY be inferred from explicit enabled controls/icons.\n                // NOTE: qual_checkmark_16 belongs to "Allows Opt-In" in some UIs and must\n                // not be treated as Enabled.\n                const enabledLink = row.querySelector(\'a[title="Enabled"]\');\n                const isEnabled = !!enabledLink || !!(\n                    enabledTd && enabledTd.querySelector(\n                        \'a[title="Enabled"], img[src*="accept_ena.png"], img[alt*="Enabled"]\'\n                    )\n                );\n\n                // Allows Opt-In can appear as clickable arrow anchor OR arrow icon in allows column.\n                const optInAnchor = row.querySelector(\'a[title="Allows Opt-In"]\');\n                const canOptIn = !!optInAnchor || !!(\n                    allowsTd && allowsTd.querySelector(\n                        \'a[title="Allows Opt-In"], img[src*="arrowgo"], img[title="Allows Opt-In"]\'\n                    )\n                );\n\n                const rk = row.getAttribute(\'_afrrk\') || row.getAttribute(\'_afrowkey\');\n\n                matches.push({\n                    functional_area: funcArea,\n                    feature: featureText,\n                    feature_col_index: featureIndex,\n                    functional_area_col_index: functionalAreaIndex,\n                    is_enabled: isEnabled,\n                    can_opt_in: canOptIn,\n                    row_index: rk,\n                    opt_in_anchor_id: optInAnchor ? optInAnchor.id : null,\n                });\n            }\n\n            return matches;\n        }""", target)\n        return result or []\n    except Exception as e:\n        log.debug(f"  read_matching_feature_rows error: {e}")\n        return []\n\n\ndef verify_feature_exists_and_enabled(page: Page, feature_name: str) -> dict:\n    """\n    Verify feature row existence first, then verify enablement indicators.\n    Returns: {"exists": bool, "is_enabled": bool}\n    """\n    log.info(f"  ✅/🔎 Verify row + enabled state: \'{feature_name[:50]}\'")\n\n    # Same style as test1 verify_auto_enabled: target the row by feature text.\n    row = page.locator(f\'tr:has-text("{feature_name[:50]}")\').first\n    try:\n        row.wait_for(state="visible", timeout=8_000)\n    except Exception:\n        return {"exists": False, "is_enabled": False}\n\n    indicator = row.locator(\n        \'a[title="Enabled"], img[src*="accept_ena.png"], \'\n        \'[class*="enabled"], img[alt*="Enabled"], \'\n        \'input[type="checkbox"]:checked, oj-switch[value="true"]\'\n    ).first\n\n    try:\n        indicator.wait_for(state="visible", timeout=4_000)\n        return {"exists": True, "is_enabled": True}\n    except Exception:\n        return {"exists": True, "is_enabled": False}\n\n\n# ─────────────────────────────────────────────────────────────────────────────\n# OPT-IN: Click arrow → Edit Features page → tick checkbox → Done\n# ─────────────────────────────────────────────────────────────────────────────\n\ndef click_opt_in_arrow(page: Page, feature_name: str, anchor_id: str | None, row_index: str | None = None) -> bool:\n    """\n    Click the Allows Opt-In arrow in the feature row.\n    Uses the anchor id if available, otherwise JS row scan.\n\n    CONFIRMED: anchor has onclick="this.focus();return false;"\n    -> use mousedown sequence.\n    """\n    log.info("  → Clicking Allows Opt-In arrow")\n    short = feature_name[:55]\n    clicked = False\n\n    # Strategy 1: use the confirmed anchor id directly\n    if anchor_id:\n        try:\n            escaped = anchor_id.replace(":", "\\\\:")\n            result = page.evaluate(f"""() => {{\n                const a = document.querySelector(\'#{escaped}\');\n                if (!a) return \'not-found\';\n                [\'mousedown\',\'mouseup\',\'click\'].forEach(n =>\n                    a.dispatchEvent(new MouseEvent(n, {{\n                        bubbles:true, cancelable:true, view:window,\n                        button:0, buttons:1\n                    }}))\n                );\n                return \'fired\';\n            }}""")\n            if result == "fired":\n                clicked = True\n                log.info(f"  ✓ Arrow clicked by id: {anchor_id}")\n        except Exception as e:\n            log.debug(f"  Arrow by id failed: {e}")\n\n    # Strategy 1b: if row index is known, click an actionable control in that row\n    # (ADF can render different titles/icons, so keep selectors flexible).\n    if not clicked and row_index:\n        try:\n            result = page.evaluate("""(rk) => {\n                const row = document.querySelector(`tr[_afrrk="${rk}"]`);\n                if (!row) return \'row-not-found\';\n\n                const candidates = [\n                    row.querySelector(\'a[title="Allows Opt-In"]\'),\n                    row.querySelector(\'a[title="Enable"]\'),\n                    row.querySelector(\'a[title*="Enable"]\'),\n                    row.querySelector(\'a:has(img[src*="arrowgo"])\'),\n                    row.querySelector(\'img[src*="arrowgo"]\')?.closest(\'a\'),\n                ].filter(Boolean);\n\n                const a = candidates[0];\n                if (!a) return \'no-action-anchor\';\n\n                [\'mousedown\',\'mouseup\',\'click\'].forEach(n =>\n                    a.dispatchEvent(new MouseEvent(n, {\n                        bubbles:true, cancelable:true, view:window,\n                        button:0, buttons:1\n                    }))\n                );\n                return \'fired:\' + (a.id || \'no-id\');\n            }""", row_index)\n            log.info(f"  Arrow S1b(row): {result}")\n            if result and result.startswith("fired:"):\n                clicked = True\n        except Exception as e:\n            log.debug(f"  Arrow S1b failed: {e}")\n\n    # Strategy 2: JS scan rows for the feature text, find a[title="Allows Opt-In"]\n    if not clicked:\n        try:\n            result = page.evaluate("""(featName) => {\n                const rows = document.querySelectorAll(\'tr[_afrrk]\');\n                for (const row of rows) {\n                    const tds = row.querySelectorAll(\'td\');\n                    const featureText = tds[2]?.textContent?.trim() || \'\';\n                    if (!featureText.includes(featName)) continue;\n                    const a = row.querySelector(\'a[title="Allows Opt-In"]\');\n                    if (!a) return \'no-arrow\';\n                    [\'mousedown\',\'mouseup\',\'click\'].forEach(n =>\n                        a.dispatchEvent(new MouseEvent(n, {\n                            bubbles:true, cancelable:true, view:window,\n                            button:0, buttons:1\n                        }))\n                    );\n                    return \'fired:\' + a.id;\n                }\n                return \'not-found\';\n            }""", short)\n            log.info(f"  Arrow S2: {result}")\n            if result and result.startswith("fired:"):\n                clicked = True\n        except Exception as e:\n            log.debug(f"  Arrow S2 failed: {e}")\n\n    # Strategy 3: physical mouse on arrowgo img\n    if not clicked:\n        try:\n            img = page.locator(\n                f\'tr:has(td:has-text("{short}")) img[src*="arrowgo"]\'\n            ).first\n            box = img.bounding_box()\n            if box:\n                cx, cy = box["x"] + box["width"]/2, box["y"] + box["height"]/2\n                page.mouse.move(cx, cy)\n                time.sleep(0.1)\n                page.mouse.down()\n                time.sleep(0.1)\n                page.mouse.up()\n                log.info(f"  Arrow S3: physical mouse ({cx:.0f},{cy:.0f})")\n                clicked = True\n        except Exception as e:\n            log.debug(f"  Arrow S3 failed: {e}")\n\n    if not clicked:\n        snap(page, f"arrow_fail_{feature_name[:20]}")\n    return clicked\n\n\ndef enable_on_edit_features_page(page: Page, feature_name: str, functional_area: str = "") -> dict:\n    """\n    After clicking the Allows Opt-In arrow we land on the Edit Features page\n    for that functional area. Find the feature row and tick its label.x17j.\n    If the feature row is absent, fall back to the captured Functional Area row.\n\n    CONFIRMED HTML (same page as parent feature enabling):\n      <span class="x2ey">\n        <span style="white-space:normal">Feature Name</span>\n      </span>\n      ...\n      <label class="x17j" for="...sbc1::content">\n    """\n    log.info(f"  → Edit Features page: enabling \'{feature_name[:55]}\'")\n    if functional_area:\n        log.info(f"  → Functional Area fallback: \'{functional_area[:55]}\'")\n\n    # Wait for the Edit Features table to load\n    try:\n        page.wait_for_selector(\'label.x17j\', state="attached", timeout=TIMEOUT)\n    except PWTimeout:\n        log.warning("  ⚠️  Edit Features page did not load (label.x17j not found)")\n        snap(page, f"edit_features_timeout_{feature_name[:20]}")\n        return {"enabled": False, "status": "edit-page-timeout", "matched_by": None}\n    time.sleep(1.0)\n\n    # Use full name for resilient matching (handles long Redwood names).\n    search_text = feature_name.strip()\n    try:\n        match_script = """(payload) => {\n            const norm = (s) => (s || \'\')\n                .toLowerCase()\n                .replace(/[^a-z0-9\\s]/g, \' \')\n                .replace(/\\s+/g, \' \')\n                .trim();\n\n            const tokenSet = (s) => new Set(norm(s).split(\' \').filter(Boolean));\n            const overlapScore = (a, b) => {\n                const A = tokenSet(a), B = tokenSet(b);\n                if (!A.size || !B.size) return 0;\n                let hit = 0;\n                for (const t of A) if (B.has(t)) hit++;\n                return hit / Math.max(A.size, B.size);\n            };\n\n            const readText = (el) => {\n                if (!el) return \'\';\n                const compactRepeatedText = (text) => {\n                    const words = (text || \'\').replace(/\\s+/g, \' \').trim().split(\' \').filter(Boolean);\n                    if (words.length && words.length % 2 === 0) {\n                        const half = words.length / 2;\n                        const left = words.slice(0, half).join(\' \').toLowerCase();\n                        const right = words.slice(half).join(\' \').toLowerCase();\n                        if (left === right) return words.slice(0, half).join(\' \');\n                    }\n                    return (text || \'\').replace(/\\s+/g, \' \').trim();\n                };\n                const uniqueParts = [];\n                const seenParts = new Set();\n                const addPart = (value) => {\n                    const cleaned = compactRepeatedText(value);\n                    const key = cleaned.toLowerCase();\n                    if (cleaned && !seenParts.has(key)) {\n                        seenParts.add(key);\n                        uniqueParts.push(cleaned);\n                    }\n                };\n                const parts = [\n                    el.innerText,\n                    el.textContent,\n                    el.getAttribute?.(\'title\'),\n                    el.getAttribute?.(\'aria-label\'),\n                ];\n                el.querySelectorAll?.(\'[title], [aria-label]\')?.forEach(child => {\n                    parts.push(child.getAttribute(\'title\'));\n                    parts.push(child.getAttribute(\'aria-label\'));\n                });\n                parts.forEach(addPart);\n                return compactRepeatedText(uniqueParts.join(\' \'));\n            };\n\n            const checkedState = (label) => {\n                const inputId = label.getAttribute(\'for\');\n                const inp = inputId ? document.getElementById(inputId) : null;\n                const aria = label.getAttribute(\'aria-checked\') || inp?.getAttribute?.(\'aria-checked\');\n                return !!(inp?.checked || aria === \'true\');\n            };\n\n            const fireClick = (el) => {\n                if (!el) return;\n                try { el.scrollIntoView({block: \'center\', inline: \'center\'}); } catch (e) {}\n                [\'mousedown\', \'mouseup\', \'click\'].forEach(n =>\n                    el.dispatchEvent(new MouseEvent(n, {\n                        bubbles: true, cancelable: true, view: window,\n                        button: 0, buttons: 1\n                    }))\n                );\n            };\n\n            const clickCandidate = (best, matchedBy, candidateCount) => {\n                const label = best.label;\n                const inputId = label.getAttribute(\'for\');\n                const inp = inputId ? document.getElementById(inputId) : null;\n\n                if (checkedState(label)) {\n                    return {\n                        enabled: true,\n                        status: \'already-checked\',\n                        matched_by: matchedBy,\n                        matched_text: best.rowText,\n                        score: best.score,\n                        candidate_count: candidateCount,\n                    };\n                }\n\n                fireClick(label);\n                if (!checkedState(label) && inp) fireClick(inp);\n                if (!checkedState(label)) {\n                    try { label.click(); } catch (e) {}\n                }\n\n                const nowChecked = checkedState(label);\n                return {\n                    enabled: nowChecked,\n                    status: nowChecked ? \'ticked\' : \'click-no-state-change\',\n                    matched_by: matchedBy,\n                    matched_text: best.rowText,\n                    score: best.score,\n                    candidate_count: candidateCount,\n                };\n            };\n\n            const collectCandidates = (targetText) => {\n                const targetNorm = norm(targetText);\n                if (!targetNorm) return [];\n\n                const rows = Array.from(document.querySelectorAll(\'tr\'));\n                const candidates = [];\n\n                for (const row of rows) {\n                const label = row.querySelector(\'label.x17j\');\n                    if (!label) continue;\n\n                    const textNodes = Array.from(row.querySelectorAll(\n                        \'span.x2ey, span[style*="white-space:normal"], td, th, a[title], span[title]\'\n                    ));\n                    const rowText = (textNodes.length ? textNodes.map(readText).join(\' \') : readText(row))\n                        .replace(/\\s+/g, \' \')\n                        .trim();\n                const rowNorm = norm(rowText);\n                if (!rowNorm) continue;\n\n                let score = 0;\n                    if (rowNorm === targetNorm) score = 1.0;\n                    else if (rowNorm.includes(targetNorm) || targetNorm.includes(rowNorm)) score = 0.9;\n                    else score = overlapScore(rowNorm, targetNorm);\n\n                if (score >= 0.45) {\n                    candidates.push({ row, label, rowText, score });\n                }\n            }\n\n                candidates.sort((a, b) => b.score - a.score || b.rowText.length - a.rowText.length);\n                return candidates;\n            };\n\n            const featureCandidates = collectCandidates(payload.feature);\n            if (featureCandidates.length) {\n                return clickCandidate(featureCandidates[0], \'feature\', featureCandidates.length);\n            }\n\n            const functionalAreaCandidates = collectCandidates(payload.functional_area);\n            if (functionalAreaCandidates.length) {\n                return clickCandidate(functionalAreaCandidates[0], \'functional_area\', functionalAreaCandidates.length);\n            }\n\n            return {\n                enabled: false,\n                status: \'not-found\',\n                matched_by: null,\n                feature_candidate_count: 0,\n                functional_area_candidate_count: 0,\n            };\n        }"""\n        scroll_script = """() => {\n            const isScrollable = (el) => {\n                if (!el) return false;\n                const style = window.getComputedStyle(el);\n                const overflowY = style.overflowY || \'\';\n                return /(auto|scroll)/.test(overflowY) && el.scrollHeight > el.clientHeight + 20;\n            };\n            const containers = Array.from(document.querySelectorAll(\n                \'div, table, tbody, [role="grid"], [role="treegrid"], [role="rowgroup"]\'\n            )).filter(isScrollable);\n            containers.sort((a, b) => (b.scrollHeight - b.clientHeight) - (a.scrollHeight - a.clientHeight));\n            const scroller = containers[0] || document.scrollingElement || document.documentElement;\n            const beforeTop = scroller.scrollTop;\n            const beforeWin = window.scrollY;\n            const step = Math.max(260, Math.floor((scroller.clientHeight || window.innerHeight || 600) * 0.75));\n            scroller.scrollTop = Math.min(scroller.scrollTop + step, scroller.scrollHeight);\n            window.scrollBy(0, step);\n            return {\n                moved: scroller.scrollTop !== beforeTop || window.scrollY !== beforeWin,\n                top: scroller.scrollTop,\n                max: Math.max(0, scroller.scrollHeight - scroller.clientHeight),\n                selector: scroller.tagName,\n            };\n        }"""\n        reset_scroll_script = """() => {\n            const isScrollable = (el) => {\n                if (!el) return false;\n                const style = window.getComputedStyle(el);\n                const overflowY = style.overflowY || \'\';\n                return /(auto|scroll)/.test(overflowY) && el.scrollHeight > el.clientHeight + 20;\n            };\n            const containers = Array.from(document.querySelectorAll(\n                \'div, table, tbody, [role="grid"], [role="treegrid"], [role="rowgroup"]\'\n            )).filter(isScrollable);\n            containers.sort((a, b) => (b.scrollHeight - b.clientHeight) - (a.scrollHeight - a.clientHeight));\n            const scroller = containers[0] || document.scrollingElement || document.documentElement;\n            scroller.scrollTop = 0;\n            window.scrollTo(0, 0);\n        }"""\n\n        payload = {"feature": search_text, "functional_area": (functional_area or "").strip()}\n        result = None\n        try:\n            page.evaluate(reset_scroll_script)\n            time.sleep(0.4)\n        except Exception:\n            pass\n\n        for scroll_attempt in range(18):\n            result = page.evaluate(match_script, payload)\n            status_for_attempt = result.get("status") if isinstance(result, dict) else result\n            if status_for_attempt != "not-found":\n                if scroll_attempt:\n                    log.info(f"  ✓ Edit Features match found after auto-scroll attempt {scroll_attempt + 1}")\n                break\n\n            scroll_info = page.evaluate(scroll_script)\n            log.info(\n                "  ↕ Edit Features auto-scroll: "\n                f"attempt={scroll_attempt + 1} moved={scroll_info.get(\'moved\')} "\n                f"top={scroll_info.get(\'top\')}/{scroll_info.get(\'max\')}"\n            )\n            if isinstance(scroll_info, dict) and not scroll_info.get("moved"):\n                try:\n                    page.mouse.wheel(0, 700)\n                    log.info("  ↕ Edit Features auto-scroll: mouse wheel fallback")\n                except Exception:\n                    pass\n            wait_busy(page, timeout=8_000)\n            time.sleep(0.45)\n            if isinstance(scroll_info, dict) and not scroll_info.get("moved"):\n                result = page.evaluate(match_script, payload)\n                status_after_wheel = result.get("status") if isinstance(result, dict) else result\n                if status_after_wheel != "not-found":\n                    log.info(f"  ✓ Edit Features match found after wheel fallback")\n                    break\n\n\n        if isinstance(result, dict):\n            log.info(\n                "  Tick result: "\n                f"{result.get(\'status\')} | candidates={result.get(\'candidate_count\')} "\n                f"| matched_by={result.get(\'matched_by\')} | score={result.get(\'score\')} "\n                f"| match=\'{str(result.get(\'matched_text\', \'\'))[:80]}\'"\n            )\n        else:\n            log.info(f"  Tick result: {result}")\n\n        status = result.get("status") if isinstance(result, dict) else result\n        matched_by = result.get("matched_by") if isinstance(result, dict) else None\n\n        if status == "already-checked":\n            log.info(f"  ℹ️  Already enabled on Edit Features page ({matched_by})")\n            return result\n        elif status == "ticked":\n            log.info(f"  ✅ Checkbox ticked ON ({matched_by})")\n            return result\n        elif status == "not-found":\n            log.warning("  ⚠️  Feature and Functional Area not found on Edit Features page")\n            snap(page, f"notfound_edit_{feature_name[:20]}")\n            return result\n        else:\n            log.warning(f"  ⚠️  Edit Features checkbox not enabled")\n            snap(page, f"notfound_edit_{feature_name[:20]}")\n            return result\n\n    except Exception as e:\n        log.warning(f"  ❌ Error on Edit Features page: {e}")\n        return {"enabled": False, "status": "error", "matched_by": None, "error": str(e)}\n\n\ndef click_done(page: Page):\n    """Click Done — saves Edit Features changes and returns to New Features."""\n    for sel in [\'a[accesskey="o"][role="button"]\', \'a.xrg[role="button"]\',\n                \'a.xrg\', \'a[role="button"]:has-text("Done")\', \'button:has-text("Done")\']:\n        try:\n            page.wait_for_selector(sel, state="visible", timeout=5_000)\n            el = page.locator(sel).first\n            mdown(el)\n            wait_busy(page)\n            log.info("  ✅ Done")\n            return\n        except Exception:\n            continue\n    # Physical mouse fallback\n    try:\n        btn = page.locator(\'a.xrg, button:has-text("Done")\').first\n        box = btn.bounding_box()\n        if box:\n            cx, cy = box["x"] + box["width"]/2, box["y"] + box["height"]/2\n            page.mouse.move(cx, cy)\n            time.sleep(0.1)\n            page.mouse.down()\n            time.sleep(0.1)\n            page.mouse.up()\n            wait_busy(page)\n            log.info("  ✅ Done (physical mouse)")\n            return\n    except Exception:\n        pass\n    log.warning("  ⚠️  Done button not found")\n\n\ndef wait_back_on_new_features(page: Page):\n    """After Done, wait until we\'re back on the New Features search page."""\n    try:\n        page.wait_for_selector(\n            \'input[aria-label="Feature"], input[name*="qbeFeature"]\',\n            state="visible", timeout=TIMEOUT,\n        )\n        wait_busy(page)\n        time.sleep(1.0)\n    except PWTimeout:\n        log.warning("  ⚠️  Did not return to New Features page after Done")\n\n\n# ─────────────────────────────────────────────────────────────────────────────\n# PROCESS ONE FEATURE\n# ─────────────────────────────────────────────────────────────────────────────\n\ndef process_feature(page: Page, feat: Feature) -> FeatureResult:\n    """\n    Search for one feature on the New Features page, check its status,\n    enable it if needed, return a FeatureResult.\n    """\n    log.info(f"\\n  ── {feat.name[:65]}")\n\n    # Search\n    search_feature(page, feat.name)\n\n    # Read all matching result rows (for functional area, per-row status, etc.)\n    matched_rows = read_matching_feature_rows(page, feat.name)\n\n    # Generic text-only row checks can match unrelated elements (e.g., Navigator);\n    # use them only as fallback when row_info can\'t be extracted.\n    quick_verify = {"exists": False, "is_enabled": False}\n    if not matched_rows:\n        quick_verify = verify_feature_exists_and_enabled(page, feat.name)\n\n    if not matched_rows and not quick_verify.get("exists", False):\n        log.warning(f"  ⚠️  Not found in New Features table")\n        snap(page, f"notfound_{feat.name[:20]}")\n        return FeatureResult(\n            name=feat.name,\n            functional_area=feat.module,\n            status="Not Found",\n            error="Feature not found in New Features search results",\n        )\n\n    # Decision logic uses only is_enabled across all matched rows:\n    # - if any row is disabled => process that row\n    # - if all rows enabled => already enabled\n    target_row = None\n    if matched_rows:\n        target_row = next((r for r in matched_rows if not r.get("is_enabled", False)), None)\n        if not target_row:\n            functional_area = matched_rows[0].get("functional_area", feat.module)\n            log.info(f"  Matched rows: {len(matched_rows)}")\n            log.info("  ✅ All matched rows already enabled")\n            return FeatureResult(\n                name=feat.name,\n                functional_area=functional_area,\n                status="Already Enabled",\n            )\n\n    functional_area = (target_row or {}).get("functional_area", feat.module)\n    is_enabled      = (target_row or {}).get("is_enabled", quick_verify.get("is_enabled", False))\n    can_opt_in      = (target_row or {}).get("can_opt_in", False)\n    anchor_id       = (target_row or {}).get("opt_in_anchor_id")\n    row_index       = (target_row or {}).get("row_index")\n\n    log.info(f"  Functional Area: {functional_area}")\n    log.info(f"  Enabled: {is_enabled} | Can Opt-In: {can_opt_in}")\n\n    # Already enabled — nothing to do\n    if is_enabled:\n        log.info("  ✅ Already enabled")\n        return FeatureResult(\n            name=feat.name,\n            functional_area=functional_area,\n            status="Already Enabled",\n        )\n\n    # Disabled row found → try enabling it by clicking its arrow.\n    arrow_clicked = click_opt_in_arrow(page, feat.name, anchor_id, row_index)\n    if not arrow_clicked:\n        return FeatureResult(\n            name=feat.name,\n            functional_area=functional_area,\n            status="Error",\n            error="Could not click Allows Opt-In arrow for disabled matched row",\n        )\n\n    wait_busy(page)\n    time.sleep(2.0)\n\n    # Now on Edit Features page — tick the feature checkbox, or fall back to Functional Area.\n    edit_result = enable_on_edit_features_page(page, feat.name, functional_area)\n    ticked = bool(edit_result.get("enabled")) if isinstance(edit_result, dict) else bool(edit_result)\n    edit_status = edit_result.get("status") if isinstance(edit_result, dict) else ""\n    matched_by = edit_result.get("matched_by") if isinstance(edit_result, dict) else None\n\n    # Click Done to save and return to New Features\n    click_done(page)\n    wait_back_on_new_features(page)\n\n    if ticked:\n        log.info(f"  ✅ Feature enabled successfully via {matched_by or \'feature\'}")\n        return FeatureResult(\n            name=feat.name,\n            functional_area=functional_area,\n            status="Enabled Now",\n        )\n    elif edit_status == "not-found":\n        log.warning("  ⚠️  Skipped: feature and Functional Area not found on Edit Features page")\n        return FeatureResult(\n            name=feat.name,\n            functional_area=functional_area,\n            status="Not Found",\n            error="Feature and Functional Area not found on Edit Features page; clicked Done and skipped",\n        )\n    else:\n        return FeatureResult(\n            name=feat.name,\n            functional_area=functional_area,\n            status="Error",\n            error=f"Arrow clicked but Edit Features enable failed ({edit_status or \'unknown\'})",\n        )\n\n    # Neither enabled nor has opt-in arrow\n    log.warning("  ⚠️  No enabled indicator and no Opt-In arrow found")\n    snap(page, f"nostatus_{feat.name[:20]}")\n    return FeatureResult(\n        name=feat.name,\n        functional_area=functional_area,\n        status="Not Found",\n        error="Row found but no Enabled or Allows Opt-In indicator detected",\n    )\n\n\n# ─────────────────────────────────────────────────────────────────────────────\n# SUMMARY REPORT\n# ─────────────────────────────────────────────────────────────────────────────\n\ndef print_summary(results: list[FeatureResult]):\n    log.info("")\n    log.info("=" * 100)\n    log.info("SUMMARY REPORT")\n    log.info("=" * 100)\n    log.info(f"  {\'#\':<4}  {\'Status\':<18}  {\'Functional Area\':<45}  {\'Feature\':<50}  {\'Error\'}")\n    log.info(f"  {\'─\'*4}  {\'─\'*18}  {\'─\'*45}  {\'─\'*50}  {\'─\'*30}")\n\n    counts = {"Already Enabled": 0, "Enabled Now": 0, "Not Found": 0, "Error": 0}\n\n    for idx, r in enumerate(results, 1):\n        icon = {\n            "Already Enabled": "✅",\n            "Enabled Now":     "🔓",\n            "Not Found":       "⚠️ ",\n            "Error":           "❌",\n        }.get(r.status, "?")\n\n        log.info(\n            f"  {idx:<4}  {icon} {r.status:<16}  "\n            f"{r.functional_area[:44]:<45}  "\n            f"{r.name[:49]:<50}  "\n            f"{r.error[:40] if r.error else \'\'}"\n        )\n        counts[r.status] = counts.get(r.status, 0) + 1\n\n    log.info("")\n    log.info(f"  Total: {len(results)}")\n    log.info(f"  ✅ Already Enabled : {counts.get(\'Already Enabled\', 0)}")\n    log.info(f"  🔓 Enabled Now     : {counts.get(\'Enabled Now\', 0)}")\n    log.info(f"  ⚠️  Not Found       : {counts.get(\'Not Found\', 0)}")\n    log.info(f"  ❌ Error           : {counts.get(\'Error\', 0)}")\n    log.info("=" * 100)\n\n    # Also save to a text file on Desktop\n    report_path = RUN_OUTPUT_DIR / f"summary_{datetime.now().strftime(\'%Y%m%d_%H%M%S\')}.txt"\n    try:\n        with open(report_path, "w") as f:\n            f.write("FEATURE ENABLEMENT SUMMARY\\n")\n            f.write("=" * 100 + "\\n")\n            f.write(f"{\'#\':<4}  {\'Status\':<18}  {\'Functional Area\':<45}  {\'Feature\':<55}  Error\\n")\n            f.write(f"{\'─\'*4}  {\'─\'*18}  {\'─\'*45}  {\'─\'*55}  {\'─\'*30}\\n")\n            for idx, r in enumerate(results, 1):\n                f.write(\n                    f"{idx:<4}  {r.status:<18}  "\n                    f"{r.functional_area[:44]:<45}  "\n                    f"{r.name[:54]:<55}  "\n                    f"{r.error or \'\'}\\n"\n                )\n            f.write("\\n")\n            f.write(f"Total            : {len(results)}\\n")\n            f.write(f"Already Enabled  : {counts.get(\'Already Enabled\', 0)}\\n")\n            f.write(f"Enabled Now      : {counts.get(\'Enabled Now\', 0)}\\n")\n            f.write(f"Not Found        : {counts.get(\'Not Found\', 0)}\\n")\n            f.write(f"Error            : {counts.get(\'Error\', 0)}\\n")\n        log.info(f"\\n  📄 Summary saved: {report_path}")\n    except Exception as e:\n        log.warning(f"  Could not save summary file: {e}")\n\n\n# ─────────────────────────────────────────────────────────────────────────────\n# MAIN\n# ─────────────────────────────────────────────────────────────────────────────\n\ndef main():\n    run_dir = init_run_artifacts()\n    log.info("=" * 70)\n    log.info("Oracle SCM – New Features OPT IN Automation")\n    log.info(f"  Excel : {EXCEL_PATH}")\n    log.info(f"  Output: {run_dir}")\n    log.info("=" * 70)\n\n    features = load_optin_features(EXCEL_PATH)\n    if not features:\n        log.error("No OPT IN features found. Check EXCEL_PATH.")\n        return\n\n    log.info(f"\\n  {len(features)} features to process\\n")\n\n    results: list[FeatureResult] = []\n\n    with sync_playwright() as pw:\n        browser = pw.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO)\n        page    = browser.new_page()\n        page.set_default_timeout(TIMEOUT)\n\n        try:\n            # ── One-time navigation to New Features page ──────────────────\n            login(page)\n            go_to_setup_and_maintenance(page)\n            go_to_offerings(page)\n            go_to_new_features(page)\n\n            # ── Process each feature ──────────────────────────────────────\n            for idx, feat in enumerate(features, start=1):\n                log.info(f"\\n▶  [{idx}/{len(features)}]  {feat.module}")\n\n                try:\n                    result = process_feature(page, feat)\n                    results.append(result)\n\n                except Exception as exc:\n                    log.error(f"  ❌ Unexpected error: {exc}", exc_info=True)\n                    snap(page, f"FAIL_{feat.name[:20]}")\n                    results.append(FeatureResult(\n                        name=feat.name,\n                        functional_area=feat.module,\n                        status="Error",\n                        error=str(exc)[:80],\n                    ))\n                    # Try to get back to New Features page\n                    try:\n                        if not page.locator(\n                            \'input[aria-label="Feature"]\'\n                        ).is_visible(timeout=3_000):\n                            go_to_setup_and_maintenance(page)\n                            go_to_offerings(page)\n                            go_to_new_features(page)\n                    except Exception:\n                        pass\n\n        except Exception as exc:\n            log.error(f"❌ Critical failure: {exc}", exc_info=True)\n            snap(page, "CRITICAL_FAILURE")\n        finally:\n            browser.close()\n\n    print_summary(results)\n\n\nif __name__ == "__main__":\n    main()\n'
PROFILE_SOURCE = 'import logging\nimport os\nimport re\nimport time\nfrom datetime import datetime\nfrom pathlib import Path\n\nfrom openpyxl import load_workbook\nfrom playwright.sync_api import sync_playwright, Page, TimeoutError as PWTimeout\n\nBASE_URL = os.getenv("FUSION_BASE_URL", "https://eghw-dev1.fa.em2.oraclecloud.com")\nUSERNAME = os.getenv("FUSION_USERNAME", "veera.h@oracle.com")\nPASSWORD = os.getenv("FUSION_PASSWORD", "Welcome_01")\n\nHEADLESS = os.getenv("HEADLESS", "false").strip().lower() in ["1", "true", "yes"]\nSLOW_MO = int(os.getenv("SLOW_MO", "400"))\nDEFAULT_TIMEOUT = int(os.getenv("DEFAULT_TIMEOUT", "45000"))\n\nlog = logging.getLogger(__name__)\nSCREENSHOT_DIR = Path.home() / "Desktop" / "fusion-screenshots"\nSCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)\n\n\ndef snap(page: Page, label: str):\n    ts = datetime.now().strftime("%H%M%S")\n    safe = "".join(c if c.isalnum() else "_" for c in label)[:40]\n    path = SCREENSHOT_DIR / f"{ts}_{safe}.png"\n    try:\n        page.screenshot(path=path, full_page=True)\n        log.info(f"  📸 {path.name}")\n    except Exception as e:\n        log.warning(f"  Screenshot failed: {e}")\n\n\ndef wait_busy(page: Page, timeout=30_000):\n    try:\n        page.wait_for_function(\n            "() => !document.querySelector("\n            "  \'[role=\\"progressbar\\"], .xbusy, [aria-busy=\\"true\\"], .oj-progress\'"\n            ")",\n            timeout=timeout,\n        )\n    except PWTimeout:\n        log.warning("  Busy indicator still present — continuing")\n\n\ndef login(page: Page):\n    log.info("🔐 Login")\n    page.goto(BASE_URL, timeout=60_000)\n    page.wait_for_load_state("domcontentloaded")\n    time.sleep(1)\n\n    login_forms = [\n        (\n            "#idcs-signin-basic-signin-form-username",\n            "#idcs-signin-basic-signin-form-password|input",\n            ".oj-button-text:has-text(\'Sign In\')",\n        ),\n        ("input#userid", "input#password", "button#btnActive"),\n        ("#userid", "#password", "#btnActive"),\n        (\'input[name="username"]\', \'input[name="password"]\', \'button[type="submit"]\'),\n        (\'input[type="email"]\', \'input[type="password"]\', \'button[type="submit"]\'),\n    ]\n\n    for u_sel, p_sel, s_sel in login_forms:\n        try:\n            page.wait_for_selector(u_sel, timeout=5_000)\n            if not page.is_visible(u_sel):\n                continue\n            page.fill(u_sel, USERNAME)\n            page.fill(p_sel, PASSWORD)\n            page.click(s_sel)\n            log.info(f"  ✓ Login form: {u_sel}")\n            break\n        except Exception:\n            continue\n\n    page.wait_for_selector(\n        \'svg[aria-label="Navigator"], [aria-label="Navigator"]\',\n        timeout=120_000,\n    )\n    wait_busy(page)\n    log.info("  ✅ Login successful")\n\n\ndef go_to_setup_and_maintenance(page: Page):\n    log.info("🔧 Setup and Maintenance")\n    for sel in [\n        \'#pt1\\\\:_UIScmil2u\',\n        \'img[title="Settings and Actions"]\',\n        \'img[alt="Settings and Actions"]\',\n    ]:\n        try:\n            page.wait_for_selector(sel, state="visible", timeout=8_000)\n            page.click(sel)\n            log.info(f"  ✓ Avatar clicked ({sel})")\n            break\n        except Exception:\n            continue\n\n    time.sleep(0.6)\n    page.click(\n        \'a:has-text("Setup and Maintenance"), \'\n        \'li:has-text("Setup and Maintenance")\',\n        timeout=10_000,\n    )\n    wait_busy(page)\n    log.info("  ✅ Setup and Maintenance open")\n\n# ── Profile Option (Step 10, Workflow 4) ──────────────────────────────────────\n\n\ndef _safe_filename(name: str) -> str:\n    """Convert task name to filesystem-safe filename fragment."""\n    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", (name or "").strip())\n    return cleaned.strip("_") or "profile_options"\n\n\ndef _profile_log_path() -> Path:\n    """Return single consolidated txt log path for full run."""\n    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")\n    return Path.cwd() / f"profile_options_status_{stamp}.txt"\n\n\ndef _init_profile_log(log_path: Path):\n    """Create log file with required column headers."""\n    with open(log_path, "w", encoding="utf-8") as f:\n        f.write("Task Name > ProfileOptionCode = Value | Status\\n")\n\n\ndef _append_profile_log_row(\n    log_path: Path,\n    task_name: str,\n    profile_option: str,\n    profile_option_value: str,\n    status: str,\n)-> str:\n    """Append one profile-option processing row."""\n    normalized_status = "PASS" if str(status).strip().upper() == "PASS" else "FAIL"\n    with open(log_path, "a", encoding="utf-8") as f:\n        f.write(\n            f"{task_name} > {profile_option} = {profile_option_value} | Status={normalized_status}\\n"\n        )\n    return normalized_status\n\n\ndef _append_profile_log_summary(log_path: Path, passed: int, failed: int):\n    """Append pass/fail summary counters."""\n    with open(log_path, "a", encoding="utf-8") as f:\n        f.write("\\n")\n        f.write(f"Total Passed: {passed}\\n")\n        f.write(f"Total Failed: {failed}\\n")\n\ndef _is_tasks_tray_open(page: Page) -> bool:\n    """Return True when Tasks tray appears open and Search link is interactable."""\n    try:\n        return page.evaluate("""() => {\n            const link = document.querySelector(\'li.x1q9 a[title="Search"]\');\n            if (!link) return false;\n            const style = window.getComputedStyle(link);\n            const rect = link.getBoundingClientRect();\n            return style.display !== \'none\' &&\n                   style.visibility !== \'hidden\' &&\n                   Number(style.opacity || 1) > 0 &&\n                   rect.width > 0 && rect.height > 0;\n        }""")\n    except Exception:\n        return False\n\n\ndef _reveal_for_click(page: Page, selector: str) -> bool:\n    """Temporarily reveal hidden ADF nodes so they can be clicked."""\n    try:\n        return page.evaluate("""(sel) => {\n            const el = document.querySelector(sel);\n            if (!el) return false;\n\n            const makeVisible = (node) => {\n                if (!node || !node.style) return;\n                node.style.display = node.style.display || \'block\';\n                node.style.visibility = \'visible\';\n                node.style.opacity = \'1\';\n                node.hidden = false;\n                node.removeAttribute(\'hidden\');\n                node.setAttribute(\'aria-hidden\', \'false\');\n            };\n\n            makeVisible(el);\n            let p = el.parentElement;\n            for (let i = 0; i < 4 && p; i++) {\n                makeVisible(p);\n                p = p.parentElement;\n            }\n\n            el.scrollIntoView({ block: \'center\', inline: \'center\' });\n            return true;\n        }""", selector)\n    except Exception:\n        return False\n\n\ndef _safe_event_click(page: Page, selector: str) -> bool:\n    """Click using JS event sequence, with Playwright fallback."""\n    try:\n        fired = page.evaluate("""(sel) => {\n            const el = document.querySelector(sel);\n            if (!el) return false;\n            [\'mousedown\',\'mouseup\',\'click\'].forEach(n => {\n                el.dispatchEvent(new MouseEvent(n, {\n                    bubbles:true, cancelable:true, view:window,\n                    button:0, buttons:1\n                }));\n            });\n            return true;\n        }""", selector)\n        if fired:\n            return True\n    except Exception:\n        pass\n\n    try:\n        page.locator(selector).first.click(force=True, timeout=2_000)\n        return True\n    except Exception:\n        return False\n\n\ndef _open_tasks_tray(page: Page):\n    """Open Tasks tray even when trigger is present but hidden by JS/CSS state."""\n    selectors = [\n        \'div[_afrptkey*="sdi10"]\',\n        \'div.x1ge.p_AFFirst\',\n        \'[role="tab"]:has-text("Tasks")\',\n        \'a[title="Tasks"]\',\n        \':text("Tasks")\',\n    ]\n\n    # Fast path\n    if _is_tasks_tray_open(page):\n        return\n\n    last_error = "No matching Tasks trigger"\n    for attempt in range(1, 4):\n        for sel in selectors:\n            try:\n                _reveal_for_click(page, sel)\n                if not _safe_event_click(page, sel):\n                    continue\n                time.sleep(0.8)\n                wait_busy(page)\n                if _is_tasks_tray_open(page):\n                    log.info(f"  ✓ Tasks tray opened ({sel}) on attempt {attempt}")\n                    return\n            except Exception as e:\n                last_error = f"{sel}: {e}"\n                continue\n\n        time.sleep(0.7)\n\n    # Diagnostics for mentor/debug conversations\n    diag = page.evaluate("""() => {\n        const taskNodes = Array.from(document.querySelectorAll(\'*\'))\n            .filter(el => (el.textContent || \'\').trim() === \'Tasks\')\n            .slice(0, 5)\n            .map(el => {\n                const s = window.getComputedStyle(el);\n                return {\n                    tag: el.tagName,\n                    cls: el.className,\n                    display: s.display,\n                    visibility: s.visibility,\n                    opacity: s.opacity,\n                    w: Math.round(el.getBoundingClientRect().width),\n                    h: Math.round(el.getBoundingClientRect().height),\n                };\n            });\n        return JSON.stringify(taskNodes);\n    }""")\n    raise RuntimeError(f"Could not open Tasks tray. Last error: {last_error}. Tasks diagnostics: {diag}")\n\n\ndef _click_task_result(page: Page, wanted_tokens: list[str], role_pattern: str) -> str:\n    """Click a task-search result link using token + role based matching."""\n\n    # Give async result rendering a few chances.\n    last_diag = ""\n    for attempt in range(1, 6):\n        wait_busy(page)\n        time.sleep(0.6)\n\n        # 1) Fast/precise path for known class, but with flexible text matching.\n        try:\n            clicked = page.evaluate("""(tokens) => {\n                const norm = s => (s || \'\').toLowerCase().replace(/\\s+/g, \' \').trim();\n                const anchors = Array.from(document.querySelectorAll(\'a.xmx\'));\n                for (const a of anchors) {\n                    const text = norm(a.textContent);\n                    if (tokens.some(t => text.includes(t))) {\n                        a.scrollIntoView({ block: \'center\', inline: \'center\' });\n                        [\'mousedown\',\'mouseup\',\'click\'].forEach(n => {\n                            a.dispatchEvent(new MouseEvent(n, {\n                                bubbles:true, cancelable:true, view:window,\n                                button:0, buttons:1\n                            }));\n                        });\n                        return `clicked-js:${a.id || \'no-id\'}`;\n                    }\n                }\n                return \'\';\n            }""", wanted_tokens)\n            if clicked:\n                return clicked\n        except Exception:\n            pass\n\n        # 2) Playwright role/text fallback (less dependent on class names).\n        try:\n            link = page.get_by_role("link", name=role_pattern).first\n            if link.count() > 0:\n                link.click(timeout=2_000)\n                return f"clicked-role:{role_pattern}"\n        except Exception:\n            pass\n\n        # 3) Generic anchor scan + helper click fallback.\n        try:\n            candidate_selectors = [\n                \'a:has-text("Manage Receiving Profile")\',\n                \'a:has-text("Profile Value")\',\n                \'a\',\n            ]\n            for sel in candidate_selectors:\n                locator = page.locator(sel)\n                cnt = min(locator.count(), 25)\n                for i in range(cnt):\n                    a = locator.nth(i)\n                    text = (a.inner_text(timeout=400) or "").strip().lower()\n                    if all(t in text for t in wanted_tokens):\n                        try:\n                            a.scroll_into_view_if_needed(timeout=1_000)\n                            a.click(timeout=1_500)\n                            return f"clicked-pw:{sel}[{i}]"\n                        except Exception:\n                            # Build a selector fallback when direct click is flaky.\n                            aid = a.get_attribute("id", timeout=400)\n                            if aid and _safe_event_click(page, f"#{aid}"):\n                                return f"clicked-fallback-id:{aid}"\n        except Exception:\n            pass\n\n        # Diagnostic snapshot for logs if we fail all attempts.\n        try:\n            last_diag = page.evaluate("""() => {\n                const norm = s => (s || \'\').replace(/\\s+/g, \' \').trim();\n                const links = Array.from(document.querySelectorAll(\'a\'))\n                    .filter(a => norm(a.textContent).length)\n                    .slice(0, 30)\n                    .map(a => ({\n                        id: a.id || \'\',\n                        cls: a.className || \'\',\n                        text: norm(a.textContent).slice(0, 140)\n                    }));\n                return JSON.stringify(links);\n            }""")\n        except Exception:\n            last_diag = "<diag unavailable>"\n\n    raise RuntimeError(f"Task link not found after retries. Visible links sample: {last_diag}")\n\n\ndef _click_named_task_result(page: Page, task_name: str) -> str:\n    """Click any named task from Search results using robust matching."""\n    norm_name = " ".join(task_name.lower().split())\n    tokens = [t for t in norm_name.split() if t]\n    escaped_pattern = r"(?i)" + r"\\s+".join(tokens)\n    return _click_task_result(page, wanted_tokens=tokens, role_pattern=escaped_pattern)\n\n\ndef _open_task_from_tasks_search(page: Page, task_name: str):\n    """Navigate via Setup and Maintenance Tasks search and open specific task."""\n    go_to_setup_and_maintenance(page)\n    time.sleep(1.5)\n\n    log.info("  → Clicking Tasks panel icon")\n    _open_tasks_tray(page)\n\n    log.info("  → Clicking Search in Tasks tray")\n    search_selectors = [\n        \'li.x1q9 a[title="Search"]\',\n        \'a[title="Search"]\',\n        \'li.x1q9 a:has-text("Search")\',\n    ]\n    clicked = False\n    for s in search_selectors:\n        try:\n            _reveal_for_click(page, s)\n            if _safe_event_click(page, s):\n                clicked = True\n                break\n        except Exception:\n            continue\n    if not clicked:\n        raise RuntimeError("Could not click Search in Tasks tray")\n\n    time.sleep(1.5)\n    wait_busy(page)\n\n    log.info(f"  → Searching task \'{task_name}\'")\n    for sel in [\n        \'input[id*="s9:it1::content"]\',\n        \'input[name*="s9:it1"]\',\n        \'input.x25[size="60"]\',\n        \'input.x25[autocomplete="on"]\',\n    ]:\n        try:\n            page.wait_for_selector(sel, state="visible", timeout=8_000)\n            page.fill(sel, task_name)\n            break\n        except Exception:\n            continue\n\n    try:\n        page.keyboard.press("Enter")\n    except Exception:\n        pass\n    time.sleep(2.0)\n    wait_busy(page)\n\n    log.info(f"  → Opening task result \'{task_name}\'")\n    last_task_err = None\n    for task_attempt in range(1, 4):\n        try:\n            result = _click_named_task_result(page, task_name)\n            log.info(f"  ✓ Task clicked: {result} (attempt {task_attempt})")\n            time.sleep(2.0)\n            wait_busy(page)\n            return\n        except Exception as e:\n            last_task_err = e\n            try:\n                page.keyboard.press("Enter")\n            except Exception:\n                pass\n            time.sleep(1.2)\n            wait_busy(page)\n\n    raise RuntimeError(f"Could not click task result \'{task_name}\': {last_task_err}")\n\n\ndef _load_profile_tasks_from_excel(excel_path: str, sheet_name: str = "Profile Options") -> list[dict]:\n    """Load profile option tasks and code lists from Excel.\n\n    Expected columns:\n      - Profile Options\n      - Profile Option Code (comma-separated)\n      - optional Profile Value\n    """\n    wb = load_workbook(excel_path, data_only=True)\n    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active\n\n    headers = [str(ws.cell(1, c).value or "").strip() for c in range(1, ws.max_column + 1)]\n    header_map = {h.lower(): i + 1 for i, h in enumerate(headers)}\n\n    c_task = header_map.get("profile options")\n    c_codes = header_map.get("profile option code")\n    c_value = header_map.get("profile value")\n    if not c_task or not c_codes:\n        raise RuntimeError("Excel must contain \'Profile Options\' and \'Profile Option Code\' columns")\n\n    rows: list[dict] = []\n    for r in range(2, ws.max_row + 1):\n        task_name = str(ws.cell(r, c_task).value or "").strip()\n        raw_codes = str(ws.cell(r, c_codes).value or "").strip()\n        row_value = str(ws.cell(r, c_value).value or "").strip() if c_value else ""\n        if not task_name or not raw_codes:\n            continue\n        codes = [c.strip() for c in raw_codes.split(",") if c.strip()]\n        if not codes:\n            continue\n        rows.append({\n            "task_name": task_name,\n            "codes": codes,\n            "profile_value": row_value or "Yes",\n        })\n    return rows\n\n\ndef _load_profile_tasks_from_redwood_excel(\n    excel_path: str,\n    sheet_name: str = "1.11 - Final Features List",\n    default_task_name: str = "Manage Administrator Profile Values",\n) -> list[dict]:\n    """Load profile-option tasks from Redwood master sheet format.\n\n    Scope (as requested):\n      - Process only rows where Action Required is exactly "Profile Options"\n      - Parse Profile Options Value entries in format:\n          <task name> > <profile option code> = Yes\n      - Support multiple entries separated by comma/newline (or mixed)\n    """\n    wb = load_workbook(excel_path, data_only=True)\n    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active\n\n    headers = [str(ws.cell(1, c).value or "").strip() for c in range(1, ws.max_column + 1)]\n    header_map = {" ".join(h.lower().split()): i + 1 for i, h in enumerate(headers) if h}\n\n    c_profile_val = header_map.get("profile options value")\n    c_action = header_map.get("action required")\n    if not c_profile_val or not c_action:\n        raise RuntimeError("Excel must contain \'Profile Options Value\' and \'Action Required\' columns")\n\n    rows: list[dict] = []\n\n    def _clean_task_name(v: str) -> str:\n        cleaned = (v or "").strip()\n        cleaned = re.sub(r"\\s+(task|page)\\s*$", "", cleaned, flags=re.IGNORECASE)\n        return cleaned.strip(" -:")\n\n    # Match: <task> > <code> = Yes/Y\n    # Keep code matcher broad enough to include non-ORA prefixes as well.\n    task_code_yes_re = re.compile(\n        r"^(?P<task>[^>]+?)\\s*>\\s*(?P<code>[A-Z0-9_]+)\\s*=\\s*(yes|y)\\s*$",\n        flags=re.IGNORECASE,\n    )\n\n    for r in range(2, ws.max_row + 1):\n        profile_col = str(ws.cell(r, c_profile_val).value or "").strip()\n        action_col = str(ws.cell(r, c_action).value or "").strip()\n\n        # Strict filter: only Action Required == "Profile Options"\n        if " ".join(action_col.lower().split()) != "profile options":\n            continue\n        if not profile_col:\n            continue\n\n        # Support comma/newline-separated multiple patterns in a single cell.\n        normalized = profile_col.replace("\\r\\n", "\\n").replace("\\r", "\\n")\n        entries = [e.strip() for e in normalized.replace("\\n", ",").split(",") if e.strip()]\n\n        grouped_codes: dict[str, list[str]] = {}\n        for entry in entries:\n            cleaned = entry.strip().strip(\'"\').strip("\'").strip("`")\n            m = task_code_yes_re.match(cleaned)\n            if not m:\n                continue\n\n            task_name = _clean_task_name(m.group("task")) or default_task_name\n            code = m.group("code").strip()\n            grouped_codes.setdefault(task_name, [])\n            if code not in grouped_codes[task_name]:\n                grouped_codes[task_name].append(code)\n\n        for task_name, codes in grouped_codes.items():\n            if not codes:\n                continue\n            rows.append(\n                {\n                    "task_name": task_name,\n                    "codes": codes,\n                    "profile_value": "Yes",\n                }\n            )\n\n    return rows\n\n\ndef _resolve_excel_path(configured_path: str) -> str:\n    """Resolve an Excel path for local portability."""\n    candidates: list[Path] = []\n    if configured_path:\n        candidates.append(Path(configured_path).expanduser())\n    candidates.append(Path.cwd() / "SCM_REDWOOD_FEATURES.xlsx")\n    candidates.append(Path(__file__).resolve().parent / "SCM_REDWOOD_FEATURES.xlsx")\n\n    for p in candidates:\n        if p.is_file():\n            return str(p.resolve())\n\n    raise RuntimeError(\n        "Excel file not found. Set PROFILE_EXCEL_PATH explicitly or place SCM_REDWOOD_FEATURES.xlsx in project directory."\n    )\n\n\ndef _load_profile_tasks_from_tsv(tsv_path: str) -> list[dict]:\n    """Load profile option tasks and code lists from messy TSV-like text.\n\n    Extracts only:\n      - task name (Profile Options/Profile Values task)\n      - profile option code(s)\n\n    Ignores boolean/value tokens from source.\n    """\n    with open(tsv_path, "r", encoding="utf-8", errors="ignore") as f:\n        lines = f.readlines()\n\n    # Primary pattern: task text followed by optional "task" and ">" marker.\n    task_pattern = re.compile(\n        r"(Manage\\s+[A-Za-z\\s]+?(?:Profile\\s+Options|Profile\\s+Values))\\s*(?:task)?\\s*>",\n        re.IGNORECASE,\n    )\n    # Fallback pattern: task text anywhere on line (for very noisy rows).\n    task_fallback_pattern = re.compile(\n        r"(Manage\\s+[A-Za-z\\s]+?(?:Profile\\s+Options|Profile\\s+Values))",\n        re.IGNORECASE,\n    )\n    code_pattern = re.compile(r"ORA_[A-Z0-9_]+_ENABLED")\n\n    grouped_codes: dict[str, list[str]] = {}\n    current_task: str | None = None\n\n    for raw in lines:\n        line = " ".join(raw.replace("\\t", " ").split())\n        if not line:\n            continue\n\n        task_match = task_pattern.search(line)\n        if not task_match:\n            task_match = task_fallback_pattern.search(line)\n        if task_match:\n            current_task = " ".join(task_match.group(1).split())\n            grouped_codes.setdefault(current_task, [])\n\n        codes = code_pattern.findall(line)\n        if codes and current_task:\n            existing = set(grouped_codes[current_task])\n            for code in codes:\n                if code not in existing:\n                    grouped_codes[current_task].append(code)\n                    existing.add(code)\n\n    rows: list[dict] = []\n    for task_name, codes in grouped_codes.items():\n        if not codes:\n            continue\n        rows.append({\n            "task_name": task_name,\n            "codes": codes,\n            "profile_value": "Yes",\n        })\n\n    return rows\n\n\ndef _resolve_tsv_path(configured_path: str) -> str:\n    """Resolve TSV path dynamically for cross-system portability.\n\n    Priority:\n      1) Explicit env/config path, if provided and exists\n      2) ./table.tsv (current working directory)\n      3) table.tsv beside this script\n      4) first *.tsv in current working directory\n    """\n    candidates: list[Path] = []\n\n    if configured_path:\n        candidates.append(Path(configured_path).expanduser())\n\n    candidates.append(Path.cwd() / "table.tsv")\n    candidates.append(Path(__file__).resolve().parent / "table.tsv")\n\n    for p in candidates:\n        if p.is_file():\n            return str(p.resolve())\n\n    try:\n        first_tsv = next(Path.cwd().glob("*.tsv"), None)\n        if first_tsv and first_tsv.is_file():\n            return str(first_tsv.resolve())\n    except Exception:\n        pass\n\n    looked = "\\n".join(f" - {str(p)}" for p in candidates)\n    raise RuntimeError(\n        "TSV file not found. Set PROFILE_TSV_PATH explicitly or place table.tsv in project directory.\\n"\n        f"Looked in:\\n{looked}"\n    )\n\n\ndef _click_profile_search_button(page: Page):\n    """Click exact Search button on Manage Receiving Profile Options page."""\n    # Prefer exact-role click to avoid the top-level global search area.\n    try:\n        page.get_by_role("button", name=r"^Search$").first.click(timeout=3_000)\n        return\n    except Exception:\n        pass\n\n    # Robust fallback selectors in ADF pages.\n    for sel in [\n        \'button:has-text("Search")\',\n        \'a[title="Search"]\',\n        \'img[title="Search"]\',\n    ]:\n        try:\n            _reveal_for_click(page, sel)\n            if _safe_event_click(page, sel):\n                return\n        except Exception:\n            continue\n\n    # Last resort\n    page.keyboard.press("Enter")\n\n\ndef _highlight_step(page: Page, selector: str, label: str = ""):\n    """Visually highlight element with blue outline so each automation step is visible."""\n    try:\n        page.evaluate(\n            """({sel, label}) => {\n                const el = document.querySelector(sel);\n                if (!el) return false;\n                el.scrollIntoView({ block: \'center\', inline: \'center\' });\n                const prevOutline = el.style.outline;\n                const prevBoxShadow = el.style.boxShadow;\n                const prevTransition = el.style.transition;\n                el.style.transition = \'all 120ms ease\';\n                el.style.outline = \'3px solid #1e90ff\';\n                el.style.boxShadow = \'0 0 0 4px rgba(30,144,255,0.30)\';\n                if (label) el.setAttribute(\'data-step-label\', label);\n                setTimeout(() => {\n                    el.style.outline = prevOutline;\n                    el.style.boxShadow = prevBoxShadow;\n                    el.style.transition = prevTransition;\n                }, 900);\n                return true;\n            }""",\n            {"sel": selector, "label": label},\n        )\n    except Exception:\n        pass\n    time.sleep(0.5)\n\n\ndef _highlight_first_available(page: Page, selectors: list[str], label: str = "") -> str | None:\n    """Highlight first matching selector and return it."""\n    for sel in selectors:\n        try:\n            if page.locator(sel).count() > 0:\n                _highlight_step(page, sel, label)\n                return sel\n        except Exception:\n            continue\n    return None\n\n\ndef _enter_profile_code_visibly(page: Page, selectors: list[str], code: str) -> str:\n    """Enter profile code in a clearly visible, step-by-step way and verify typed value."""\n    last_err = None\n    for sel in selectors:\n        try:\n            field = page.locator(sel).first\n            field.wait_for(state="visible", timeout=10_000)\n            _highlight_step(page, sel, label=f"Typing: {code}")\n\n            # Hard clear so every loop starts fresh and visibly.\n            field.click(timeout=2_000)\n            mod = "Meta" if os.name == "posix" else "Control"\n            page.keyboard.press(f"{mod}+A")\n            page.keyboard.press("Backspace")\n            field.fill("")\n            time.sleep(0.35)\n\n            # Slow typing so user can see each value entering.\n            field.type(code, delay=90)\n            time.sleep(0.45)\n\n            # Validate typed value; retry once if mismatch.\n            typed = (field.input_value(timeout=2_000) or "").strip()\n            if typed != code:\n                field.click(timeout=2_000)\n                page.keyboard.press(f"{mod}+A")\n                page.keyboard.press("Backspace")\n                field.type(code, delay=90)\n                time.sleep(0.45)\n                typed = (field.input_value(timeout=2_000) or "").strip()\n\n            if typed != code:\n                raise RuntimeError(f"typed=\'{typed}\' expected=\'{code}\'")\n\n            log.info(f"  ✓ Profile code entered visibly ({sel})")\n            return sel\n        except Exception as e:\n            last_err = e\n            continue\n\n    raise RuntimeError(f"Could not enter profile code \'{code}\' in any selector: {last_err}")\n\n\ndef _click_save(page: Page):\n    """Click exact Save button (not Save and Close)."""\n    try:\n        page.get_by_role("button", name=r"^Save$").first.click(timeout=3_000)\n        return\n    except Exception:\n        pass\n\n    for sel in [\n        \'button:has-text("Save")\',\n        \'button[id*="APsv"]\',\n        \'button[accesskey="S"]\',\n    ]:\n        try:\n            locator = page.locator(sel)\n            cnt = min(locator.count(), 10)\n            for i in range(cnt):\n                btn = locator.nth(i)\n                text = (btn.inner_text(timeout=500) or "").strip().lower()\n                if text == "save":\n                    btn.click(timeout=2_000)\n                    return\n        except Exception:\n            continue\n\n    raise RuntimeError("Save button not found")\n\n\ndef _click_ok(page: Page):\n    """Click OK button after profile value selection."""\n    try:\n        page.get_by_role("button", name=r"^OK$").first.click(timeout=3_000)\n        return\n    except Exception:\n        pass\n\n    for sel in [\n        \'button:has-text("OK")\',\n        \'a:has-text("OK")\',\n        \'button[id*="ok"]\',\n    ]:\n        try:\n            locator = page.locator(sel)\n            cnt = min(locator.count(), 10)\n            for i in range(cnt):\n                btn = locator.nth(i)\n                text = (btn.inner_text(timeout=500) or "").strip().lower()\n                if text == "ok":\n                    btn.click(timeout=2_000)\n                    return\n        except Exception:\n            continue\n\n    # Some ADF pages don\'t show separate OK after dropdown selection.\n    log.info("  ℹ OK button not found; continuing to Save")\n\n\ndef _assert_search_results_for_code(page: Page, code: str):\n    """Ensure profile-option search returned a usable result row for this code."""\n    no_result_markers = [\n        "No results found",\n        "No results found.",\n    ]\n    for marker in no_result_markers:\n        if page.get_by_text(marker, exact=False).count() > 0:\n            raise RuntimeError(f"Search returned no results for code \'{code}\'")\n\n    # Validate at least a partial code token appears in search results section.\n    # Many UIs truncate long values, so use meaningful prefix token.\n    token = code[:18]\n    try:\n        code_hits = page.locator(f"td:has-text(\'{token}\')")\n        if code_hits.count() == 0:\n            # Fallback generic text search.\n            if page.get_by_text(token, exact=False).count() == 0:\n                raise RuntimeError(f"Could not verify search result row for \'{code}\'")\n    except Exception as e:\n        raise RuntimeError(f"Search verification failed for \'{code}\': {e}")\n\n\ndef _ensure_profile_value_control_supported(page: Page):\n    """Fail when Profile Value control is not a supported Yes/No dropdown."""\n    dropdown_selectors = [\n        \'select[id*="soc2::content"]\',\n        \'select[name*="soc2"]\',\n        \'select.x2h\',\n    ]\n    input_selectors = [\n        \'input[id*="soc2::content"]\',\n        \'input[name*="soc2"]\',\n        \'input[aria-label*="Profile Value"]\',\n        \'textarea[id*="soc2::content"]\',\n        \'textarea[name*="soc2"]\',\n    ]\n\n    has_dropdown = any(page.locator(sel).count() > 0 for sel in dropdown_selectors)\n    if not has_dropdown:\n        if any(page.locator(sel).count() > 0 for sel in input_selectors):\n            raise RuntimeError("Profile Value is rendered as input/text field (expected dropdown)")\n        raise RuntimeError("Profile Value dropdown control not found")\n\n\ndef _validate_dropdown_has_yes_no(page: Page):\n    """Ensure dropdown options include Yes/No semantics (Yes/No or Y/N)."""\n    value_selectors = [\n        \'select[id*="soc2::content"]\',\n        \'select[name*="soc2"]\',\n        \'select.x2h\',\n    ]\n\n    last_err = None\n    for sel in value_selectors:\n        try:\n            page.locator(sel).first.wait_for(state="visible", timeout=3_000)\n            options = page.evaluate(\n                """(s) => {\n                    const d = document.querySelector(s);\n                    if (!d) return [];\n                    return Array.from(d.options || []).map(o => ({\n                        value: (o.value || \'\').trim().toLowerCase(),\n                        label: (o.text || \'\').trim().toLowerCase(),\n                    }));\n                }""",\n                sel,\n            ) or []\n\n            values = {o.get("value", "") for o in options}\n            labels = {o.get("label", "") for o in options}\n            has_yes = ("y" in values) or ("yes" in labels)\n            has_no = ("n" in values) or ("no" in labels)\n\n            if not (has_yes and has_no):\n                raise RuntimeError(\n                    f"Unsupported dropdown options for Profile Value. Found: {sorted(labels)} / {sorted(values)}"\n                )\n            return\n        except Exception as e:\n            last_err = e\n            continue\n\n    raise RuntimeError(f"Could not validate Profile Value dropdown options: {last_err}")\n\n\ndef _set_profile_value_with_verification(page: Page, profile_value: str):\n    """Set dropdown to expected value and verify selection; raise if not applied."""\n    desired_is_yes = profile_value.strip().lower() in ["yes", "y"]\n    desired_value = "Y" if desired_is_yes else "N"\n\n    value_selectors = [\n        \'select[id*="soc2::content"]\',\n        \'select[name*="soc2"]\',\n        \'select.x2h\',\n    ]\n\n    last_err = None\n    for sel in value_selectors:\n        try:\n            dropdown = page.locator(sel).first\n            dropdown.wait_for(state="visible", timeout=8_000)\n\n            # Try multiple set strategies for ADF dropdown variability.\n            for _ in range(3):\n                try:\n                    dropdown.select_option(value=desired_value)\n                except Exception:\n                    try:\n                        dropdown.select_option(label="Yes" if desired_is_yes else "No")\n                    except Exception:\n                        pass\n\n                time.sleep(0.4)\n\n                selected_value = ""\n                selected_label = ""\n                try:\n                    selected_value = (dropdown.input_value(timeout=1_500) or "").strip()\n                except Exception:\n                    selected_value = ""\n                try:\n                    selected_label = page.evaluate(\n                        """(sel) => {\n                            const d = document.querySelector(sel);\n                            if (!d) return \'\';\n                            const idx = d.selectedIndex;\n                            if (idx < 0) return \'\';\n                            return (d.options[idx]?.text || \'\').trim();\n                        }""",\n                        sel,\n                    ) or ""\n                except Exception:\n                    selected_label = ""\n\n                selected_value_low = selected_value.lower()\n                selected_label_low = selected_label.lower()\n                if desired_is_yes and (selected_value_low == "y" or selected_label_low == "yes"):\n                    log.info("  ✓ Dropdown verified as Yes")\n                    return\n                if (not desired_is_yes) and (selected_value_low == "n" or selected_label_low == "no"):\n                    log.info("  ✓ Dropdown verified as No")\n                    return\n\n            raise RuntimeError(\n                f"Dropdown did not persist expected value \'{profile_value}\' after retries"\n            )\n        except Exception as e:\n            last_err = e\n            continue\n\n    raise RuntimeError(f"Profile Value dropdown unavailable or not settable: {last_err}")\n\ndef set_profile_options(\n    page: Page,\n    profile_codes: list[str],\n    profile_value: str = "Yes",\n    task_name: str = "Manage Receiving Profile Options",\n    log_path: Path | None = None,\n) -> tuple[int, int]:\n    """\n    Set one or more profile options via Setup and Maintenance → Tasks → Search\n    → Manage Administrator Profile Values.\n\n    Handles the confirmed live HTML:\n      Tasks icon: div[_afrptkey*="sdi10"] (ADF tab — no onclick blocker)\n      Search link: li.x1q9 a[title="Search"]\n      Search input: input[id*="s9:it1::content"]\n      Task link: a.xmx with text "Manage Administrator Profile Values"\n      Profile code input: input[aria-label=" Profile Option Code"]\n      Dropdown: select[id*="soc2::content"] with value="Y" for Yes\n      Save button: button[id*="APscl"]\n    """\n    log.info(f"⚙️  Set profile options: {profile_codes}")\n\n    _open_task_from_tasks_search(page, task_name)\n\n    passed_count = 0\n    failed_count = 0\n\n    if log_path is None:\n        log_path = _profile_log_path()\n        _init_profile_log(log_path)\n\n    for code in profile_codes:\n        try:\n            log.info(f"  → Processing: {code}")\n\n            # Enter profile option code\n            log.info(f"  → Entering profile code: {code}")\n            code_selectors = [\'input[aria-label=" Profile Option Code"]\',\n                              \'input[aria-label*="Profile Option Code"]\',\n                              \'input[id*="qryId1:value00::content"]\',\n                              \'input[name*="qryId1:value00"]\',\n                              \'input.x25[autocomplete="off"][size="30"]\']\n            _highlight_first_available(page, code_selectors, label=f"Profile code: {code}")\n            _enter_profile_code_visibly(page, code_selectors, code)\n\n            _highlight_first_available(\n                page,\n                [\'button:has-text("Search")\', \'button[title="Search"]\', \'a[title="Search"]\'],\n                label="Search",\n            )\n            _click_profile_search_button(page)\n            time.sleep(2.5)\n            wait_busy(page)\n\n            # Ensure search result exists; if not, this code must be marked FAIL.\n            _assert_search_results_for_code(page, code)\n\n            # Fail this code if control is text/input or unsupported dropdown.\n            _ensure_profile_value_control_supported(page)\n            _validate_dropdown_has_yes_no(page)\n\n            # Set profile value via dropdown\n            log.info(f"  → Setting value to: {profile_value}")\n            value_selectors = [\'select[id*="soc2::content"]\', \'select[name*="soc2"]\', \'select.x2h\']\n            _highlight_first_available(page, value_selectors, label=f"Profile Value: {profile_value}")\n            _set_profile_value_with_verification(page, profile_value)\n\n            _highlight_first_available(page, [\'button:has-text("OK")\', \'a:has-text("OK")\'], label="OK")\n            _click_ok(page)\n            time.sleep(1.0)\n            wait_busy(page)\n\n            snap(page, f"profile_set_{code[:20]}")\n\n            # Save after each entry before processing next code\n            log.info("  → Save")\n            _highlight_first_available(page, [\'button:has-text("Save")\', \'button[id*="APsv"]\'], label="Save")\n            try:\n                _click_save(page)\n                log.info("  ✓ Saved")\n            except Exception:\n                # If page/context got closed (navigation/session), avoid crashing on fallback.\n                if page.is_closed():\n                    log.warning("  ⚠ Page closed right after Save attempt; skipping fallback click")\n                else:\n                    raise\n\n            wait_busy(page)\n            time.sleep(1.5)\n            log.info(f"  ✅ Profile option {code} saved")\n\n            final_status = _append_profile_log_row(log_path, task_name, code, profile_value, "PASS")\n            if final_status == "PASS":\n                passed_count += 1\n            else:\n                failed_count += 1\n\n        except Exception as e:\n            log.error(f"  ❌ Failed for profile option {code}: {e}")\n            final_status = _append_profile_log_row(log_path, task_name, code, profile_value, "FAIL")\n            if final_status == "PASS":\n                passed_count += 1\n            else:\n                failed_count += 1\n\n    return passed_count, failed_count\n\n\ndef main():\n    """Standalone runner for profile-option automation only."""\n    raw_codes = os.getenv("PROFILE_CODES", "").strip()\n    profile_value = os.getenv("PROFILE_VALUE", "Yes").strip() or "Yes"\n    excel_path = os.getenv(\n        "PROFILE_EXCEL_PATH",\n        "",\n    ).strip()\n    excel_sheet = os.getenv("PROFILE_EXCEL_SHEET", "1.11 - Final Features List").strip() or "1.11 - Final Features List"\n    input_mode = (os.getenv("PROFILE_INPUT_MODE", "redwood_excel").strip() or "redwood_excel").lower()\n    tsv_path = os.getenv(\n        "PROFILE_TSV_PATH",\n        "",\n    ).strip()\n\n    if raw_codes:\n        profile_codes = [c.strip() for c in raw_codes.split(",") if c.strip()]\n        task_rows = [{\n            "task_name": os.getenv("PROFILE_TASK_NAME", "Manage Administrator Profile Values").strip() or "Manage Administrator Profile Values",\n            "codes": profile_codes,\n            "profile_value": profile_value,\n        }]\n    elif input_mode == "tsv":\n        tsv_path = _resolve_tsv_path(tsv_path)\n        log.info(f"Using TSV path: {tsv_path}")\n        task_rows = _load_profile_tasks_from_tsv(tsv_path)\n        if not task_rows:\n            raise RuntimeError(f"No valid profile option rows found in TSV: {tsv_path}")\n    elif input_mode == "excel":\n        excel_path = _resolve_excel_path(excel_path)\n        task_rows = _load_profile_tasks_from_excel(excel_path, excel_sheet)\n        if not task_rows:\n            raise RuntimeError(f"No valid profile option rows found in Excel: {excel_path} [{excel_sheet}]")\n    else:\n        excel_path = _resolve_excel_path(excel_path)\n        task_name = os.getenv("PROFILE_TASK_NAME", "Manage Administrator Profile Values").strip() or "Manage Administrator Profile Values"\n        task_rows = _load_profile_tasks_from_redwood_excel(excel_path, excel_sheet, default_task_name=task_name)\n        if not task_rows:\n            raise RuntimeError(\n                f"No profile-option rows found in Redwood sheet: {excel_path} [{excel_sheet}]. "\n                "Check \'Action Required\' contains \'profile option\' and ORA_* exists in Profile Options Value column."\n            )\n\n    log.info("=" * 70)\n    log.info("Task Tray Profile Automation Test (Excel multi-profile loop)")\n    log.info(f"  PROFILE_ROWS: {len(task_rows)}")\n    log.info("=" * 70)\n\n    log_path = _profile_log_path()\n    _init_profile_log(log_path)\n    total_passed = 0\n    total_failed = 0\n    summary_written = False\n\n    with sync_playwright() as pw:\n        browser = pw.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO)\n        context = browser.new_context(ignore_https_errors=True)\n        page = context.new_page()\n        page.set_default_timeout(DEFAULT_TIMEOUT)\n\n        try:\n            login(page)\n            for idx, row in enumerate(task_rows, start=1):\n                log.info("-" * 70)\n                log.info(\n                    f"[{idx}/{len(task_rows)}] Task=\'{row[\'task_name\']}\' | Codes={len(row[\'codes\'])} | Value={row[\'profile_value\']}"\n                )\n                try:\n                    passed, failed = set_profile_options(\n                        page,\n                        row["codes"],\n                        row["profile_value"],\n                        task_name=row["task_name"],\n                        log_path=log_path,\n                    )\n                    total_passed += passed\n                    total_failed += failed\n                except Exception as row_err:\n                    log.error(\n                        f"❌ Row failed for task \'{row[\'task_name\']}\'. Marking all its profile options as FAIL. Error: {row_err}"\n                    )\n                    for code in row.get("codes", []):\n                        final_status = _append_profile_log_row(\n                            log_path,\n                            row.get("task_name", "Manage Administrator Profile Values"),\n                            code,\n                            row.get("profile_value", "Yes"),\n                            "FAIL",\n                        )\n                        if final_status == "PASS":\n                            total_passed += 1\n                        else:\n                            total_failed += 1\n                    continue\n\n            _append_profile_log_summary(log_path, total_passed, total_failed)\n            summary_written = True\n            log.info(f"📝 Consolidated profile option log generated: {log_path}")\n            log.info("✅ Profile option automation test completed")\n        except Exception as e:\n            log.error(f"❌ Profile option automation test failed: {e}", exc_info=True)\n            try:\n                snap(page, "TASK_TRAY_PROFILE_TEST_FAIL")\n            except Exception:\n                pass\n            raise\n        finally:\n            if not summary_written:\n                try:\n                    _append_profile_log_summary(log_path, total_passed, total_failed)\n                    log.info(f"📝 Partial consolidated profile option log generated: {log_path}")\n                except Exception:\n                    pass\n            browser.close()\n\n\nif __name__ == "__main__":\n    main()\n\n'
ESS_SOURCE = '"""\nOracle Fusion – Navigate to Scheduled Processes\n================================================\nStandalone script that:\n  1. Logs in to Oracle Fusion\n  2. Clicks the Navigator (hamburger) menu (top-left)\n  3. Scrolls down and expands the Tools section\n  4. Clicks "Scheduled Processes"\n  5. Clicks "Schedule New Process"\n  6. Searches for a process (optional)\n\nAll selectors derived from confirmed live HTML.\n\nUsage:\n    pip install playwright python-dotenv\n    playwright install chromium\n    python scheduled_processes.py\n"""\n\nimport logging\nimport os\nimport re\nimport sys\nimport time\nfrom pathlib import Path\nfrom datetime import datetime\nfrom difflib import SequenceMatcher\nfrom typing import Dict, List, Literal, Optional, TypedDict\n\nimport openpyxl\nfrom dotenv import load_dotenv\nfrom playwright.sync_api import sync_playwright, Page, TimeoutError as PWTimeout\n\nload_dotenv()\n\n# ─────────────────────────────────────────────────────────────────────────────\n# CONFIGURATION\n# ─────────────────────────────────────────────────────────────────────────────\n\nBASE_URL        = os.getenv("FUSION_BASE_URL", "https://eghw-dev1.fa.em2.oraclecloud.com/fscmUI/faces/FuseTaskListManagerTop?_afrLoop=60355488926401771&_adf.ctrl-state=xxa77g8uf_498")\nUSERNAME        = os.getenv("FUSION_USERNAME", "veera.h@oracle.com")\nPASSWORD        = os.getenv("FUSION_PASSWORD", "Welcome_01")\n\n# Optional: Process name to search for after opening "Schedule New Process"\nPROCESS_SEARCH  = os.getenv("PROCESS_SEARCH", "")  # e.g., "Import Supplier"\nADDITIONAL_INFO_EXCEL_FILE = os.getenv("ADDITIONAL_INFO_EXCEL_FILE", "SCM_REDWOOD_FEATURES.xlsx")\nADDITIONAL_INFO_SHEETS = os.getenv("ADDITIONAL_INFO_SHEETS", "1.11 - Final Features List")\n\n# Playwright settings\nHEADLESS        = False\nSLOW_MO         = 200\nDEFAULT_TIMEOUT = 90_000\n\n# ─────────────────────────────────────────────────────────────────────────────\n# LOGGING & SCREENSHOTS\n# ─────────────────────────────────────────────────────────────────────────────\n\nlogging.basicConfig(\n    level=logging.INFO,\n    format="%(asctime)s  [%(levelname)s]  %(message)s",\n    datefmt="%H:%M:%S",\n)\nlog = logging.getLogger(__name__)\n\nSCREENSHOT_DIR = Path.home() / "Desktop" / "fusion-screenshots"\nSCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)\n\nESS_JOB_LOG_FILE = Path.cwd() / "ess_job_execution_log.txt"\n\n\ndef snap(page: Page, label: str):\n    """Take a screenshot for debugging."""\n    ts = datetime.now().strftime("%H%M%S")\n    safe = "".join(c if c.isalnum() else "_" for c in label)[:40]\n    path = SCREENSHOT_DIR / f"{ts}_{safe}.png"\n    try:\n        page.screenshot(path=path, full_page=True)\n        log.info(f"  📸 {path.name}")\n    except Exception as e:\n        log.warning(f"  Screenshot failed: {e}")\n\n\ndef dump_html(page: Page, label: str, locator=None):\n    """Dump page or locator HTML for dialog-level debugging."""\n    ts = datetime.now().strftime("%H%M%S")\n    safe = "".join(c if c.isalnum() else "_" for c in label)[:40]\n    path = SCREENSHOT_DIR / f"{ts}_{safe}.html"\n    try:\n        html = locator.inner_html() if locator else page.content()\n        path.write_text(html, encoding="utf-8")\n        log.info(f"  🧾 {path.name}")\n    except Exception as e:\n        log.warning(f"  HTML dump failed: {e}")\n\n\ndef sanitize_parameter_value(value: object) -> str:\n    """Remove commas from parameter values before UI fill/logging."""\n    return str(value or "").replace(",", "").strip()\n\n\ndef wait_busy(page: Page, timeout=30_000):\n    """Wait until Oracle\'s busy spinner/disabled overlays disappear."""\n    try:\n        page.wait_for_function(\n            "() => !document.querySelector("\n            "  \'[role=\\"progressbar\\"], .xbusy, [aria-busy=\\"true\\"], .oj-progress, .AFBusyIndicator\'"\n            ")",\n            timeout=timeout,\n        )\n    except PWTimeout:\n        log.warning("  ⚠️  Busy indicator still present — continuing anyway")\n\n\ndef js_mousedown(page: Page, el, label="element"):\n    """\n    Fire mousedown + mouseup + click sequence for ADF compatibility.\n    ADF often blocks standard clicks via onclick="return false".\n    """\n    try:\n        el.evaluate("""el => {\n            [\'mousedown\',\'mouseup\',\'click\'].forEach(name => {\n                el.dispatchEvent(new MouseEvent(name, {\n                    bubbles: true, cancelable: true,\n                    view: window, button: 0, buttons: 1,\n                    clientX: 1, clientY: 1\n                }));\n            });\n        }""")\n        log.info(f"  ✅ Clicked {label} via mousedown sequence")\n        return True\n    except Exception as e:\n        log.debug(f"  mousedown on {label} failed: {e}")\n        return False\n\n\ndef debug_dom_info(page: Page, description: str):\n    """Dump relevant DOM info for diagnosis."""\n    debug_info = page.evaluate("""() => ({\n        url: window.location.href,\n        title: document.title,\n        inIframe: window !== window.top,\n        overlay: !!document.querySelector(\'.AFPopup, .dialog, [role="dialog"]\'),\n        busy: !!document.querySelector(\'[role="progressbar"], .xbusy, .oj-progress\')\n    })""")\n    log.info(f"  🕵️ {description}: {debug_info}")\n    return debug_info\n\n\n# ─────────────────────────────────────────────────────────────────────────────\n# PARAMETER HANDLING (INLINED)\n# ─────────────────────────────────────────────────────────────────────────────\n\nParameter = Dict[str, str]\n\n\nclass FillPlan(TypedDict):\n    action: Literal["SUBMIT_ONLY", "FILL_AND_SUBMIT"]\n    parameters_to_fill: List[Parameter]\n    extra_input_labels: List[str]\n    submission_allowed: bool\n    message: str\n\n\ndef normalize_label(label: str) -> str:\n    text = str(label).strip().lower()\n    text = re.sub(r"[^a-z0-9\\s]", " ", text)\n    return " ".join(text.split())\n\n\ndef values_match(actual: str, expected: str) -> bool:\n    a = normalize_label(actual)\n    b = normalize_label(expected)\n    if not a or not b:\n        return False\n    return a == b or a in b or b in a\n\n\ndef similarity_score(a: str, b: str) -> float:\n    na = normalize_label(a)\n    nb = normalize_label(b)\n    if not na or not nb:\n        return 0.0\n    if na == nb:\n        return 1.0\n    if na in nb or nb in na:\n        return 0.95\n    return SequenceMatcher(None, na, nb).ratio()\n\n\ndef work_area_match(excel_area: str, ui_area: str) -> bool:\n    a = normalize_label(excel_area)\n    b = normalize_label(ui_area)\n    if not a or not b:\n        return False\n    if a == b or a in b or b in a:\n        return True\n\n    def strip_setup_prefix(v: str) -> str:\n        v = v.replace("setup and maintenance", "setup")\n        if v.startswith("setup "):\n            v = v[len("setup "):].strip()\n        return v\n\n    aa = strip_setup_prefix(a)\n    bb = strip_setup_prefix(b)\n    return bool(aa and bb and (aa == bb or aa in bb or bb in aa))\n\n\ndef load_process_params_from_additional_info(file_path: str) -> List[Dict[str, str]]:\n    """\n    Extract rows where `Action Required` equals "ESS Jobs" and parse only\n    those `ESS Jobs Value` cells that exactly follow the pattern:\n      Process name > Parameter label > Parameter value\n\n    A single cell may contain multiple patterns separated by commas and/or\n    line breaks. Any text that does not match the strict three-part pattern is\n    ignored.\n\n    Expected columns (case-insensitive):\n      - ESS Jobs Value\n      - Action Required\n    """\n\n    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)\n    entries: List[Dict[str, str]] = []\n    allowed_sheets = {\n        s.strip() for s in str(ADDITIONAL_INFO_SHEETS or "").split(",") if s.strip()\n    }\n\n    pattern = re.compile(r"([^>\\r\\n]+)>\\s*([^>\\r\\n]+)>\\s*([^>\\r\\n]+)")\n\n    def _norm(v: object) -> str:\n        return " ".join(str(v or "").replace("\\xa0", " ").split()).strip()\n\n    def _extract_matches(value_text: str) -> List[Dict[str, str]]:\n        matches: List[Dict[str, str]] = []\n        for m in pattern.finditer(value_text):\n            process_name = _norm(m.group(1))\n            param_label = _norm(m.group(2))\n            param_value = _norm(m.group(3))\n            if process_name and param_label and param_value:\n                matches.append(\n                    {\n                        "process_name": process_name,\n                        "param_label": param_label,\n                        "param_value": param_value,\n                    }\n                )\n        return matches\n\n    try:\n        sheet_titles = {ws.title for ws in wb.worksheets}\n        if allowed_sheets and not (allowed_sheets & sheet_titles):\n            log.warning(\n                "  ⚠️ ADDITIONAL_INFO_SHEETS value %r not found. Falling back to all sheets: %s",\n                ADDITIONAL_INFO_SHEETS,\n                sorted(sheet_titles),\n            )\n            allowed_sheets.clear()\n\n        dedup = set()  # avoid duplicate (process, label, value, sheet, row) rows\n\n        for ws in wb.worksheets:\n            if allowed_sheets and ws.title not in allowed_sheets:\n                continue\n\n            headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]\n            normalized_headers = [str(h).strip().lower() if h is not None else "" for h in headers]\n\n            ess_job_header_aliases = {"ess jobs value", "ess job value"}\n            ess_job_present = any(h in ess_job_header_aliases for h in normalized_headers)\n\n            if not ess_job_present or "action required" not in normalized_headers:\n                continue\n\n            # Prefer plural header; fall back to singular.\n            if "ess jobs value" in normalized_headers:\n                info_col = normalized_headers.index("ess jobs value") + 1\n            else:\n                info_col = normalized_headers.index("ess job value") + 1\n            action_col = normalized_headers.index("action required") + 1\n\n            for row_idx in range(2, ws.max_row + 1):\n                action_required = _norm(ws.cell(row_idx, action_col).value)\n                if action_required.lower() != "ess jobs":\n                    continue\n\n                raw = ws.cell(row_idx, info_col).value\n                if raw is None:\n                    log.info(\n                        "  ℹ️ ESS Jobs Value parse [sheet=%s row=%s]: skipped (ESS Jobs row but ESS Jobs Value empty)",\n                        ws.title,\n                        row_idx,\n                    )\n                    continue\n\n                text = str(raw).replace("\\xa0", " ").strip()\n                if not text:\n                    log.info(\n                        "  ℹ️ ESS Jobs Value parse [sheet=%s row=%s]: skipped (ESS Jobs row but ESS Jobs Value blank)",\n                        ws.title,\n                        row_idx,\n                    )\n                    continue\n\n                matches = _extract_matches(text)\n                if not matches:\n                    log.info(\n                        "  ℹ️ ESS Jobs Value parse [sheet=%s row=%s]: ignored (no valid \'A>B>C\' pattern found)",\n                        ws.title,\n                        row_idx,\n                    )\n                    continue\n\n                for match in matches:\n                    key = (\n                        ws.title,\n                        row_idx,\n                        match["process_name"].lower(),\n                        match["param_label"].lower(),\n                        match["param_value"].lower(),\n                    )\n                    if key in dedup:\n                        log.info(\n                            "  ℹ️ ESS Jobs Value parse [sheet=%s row=%s]: duplicate entry skipped",\n                            ws.title,\n                            row_idx,\n                        )\n                        continue\n                    dedup.add(key)\n                    entries.append(\n                        {\n                            "sheet": ws.title,\n                            "row": str(row_idx),\n                            "action_required": action_required,\n                            "process_name": match["process_name"],\n                            "param_label": match["param_label"],\n                            "param_value": match["param_value"],\n                        }\n                    )\n                    log.info(\n                        "  ✅ ESS Jobs Value parsed [sheet=%s row=%s]: process=\'%s\' | label=\'%s\' | value=\'%s\'",\n                        ws.title,\n                        row_idx,\n                        match["process_name"],\n                        match["param_label"],\n                        match["param_value"],\n                    )\n        return entries\n    finally:\n        wb.close()\n\n\ndef find_best_additional_info_entry(entries: List[Dict[str, str]], requested_process_name: str) -> Optional[Dict[str, str]]:\n    requested = (requested_process_name or "").strip()\n    if not entries:\n        return None\n    if not requested:\n        return entries[0]\n\n    for e in entries:\n        if values_match(e.get("process_name", ""), requested):\n            return e\n\n    best = None\n    best_score = 0.0\n    for e in entries:\n        score = similarity_score(e.get("process_name", ""), requested)\n        if score > best_score:\n            best_score = score\n            best = e\n    return best\n\n\ndef load_parameters_from_excel(\n    file_path: str,\n    process_name: str,\n    work_area: str,\n    sheet_name: str = "Parameters",\n) -> List[Parameter]:\n    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)\n    try:\n        if sheet_name not in wb.sheetnames:\n            raise ValueError(\n                f"Sheet \'{sheet_name}\' not found in \'{file_path}\'. Available sheets: {wb.sheetnames}"\n            )\n\n        ws = wb[sheet_name]\n        rows = list(ws.iter_rows(min_row=2, values_only=True))\n\n        matched: List[Parameter] = []\n        process_only_matches: List[Parameter] = []\n        best_process_score = 0.0\n\n        for i, row in enumerate(rows, start=2):\n            if not any(row):\n                continue\n\n            col_process, col_area, col_label, col_value = (list(row) + [None, None, None, None])[:4]\n            process_cell = str(col_process or "")\n            process_matches = values_match(process_cell, process_name)\n            process_score = similarity_score(process_cell, process_name)\n            area_matches = work_area_match(str(col_area or ""), work_area)\n\n            if process_matches and area_matches:\n                label = str(col_label).strip() if col_label is not None else ""\n                value = str(col_value).strip() if col_value is not None else ""\n                if not label:\n                    log.warning("Row %d: Parameter Label is empty — skipped.", i)\n                    continue\n                matched.append({"label": label, "value": value})\n                best_process_score = max(best_process_score, process_score)\n            elif process_matches:\n                label = str(col_label).strip() if col_label is not None else ""\n                value = str(col_value).strip() if col_value is not None else ""\n                if label:\n                    process_only_matches.append({"label": label, "value": value})\n\n        if matched:\n            log.info(\n                "Excel: loaded %d rows for process=\'%s\', work_area=\'%s\' (best score=%.2f)",\n                len(matched), process_name, work_area, best_process_score,\n            )\n            return matched\n\n        if process_only_matches:\n            log.warning(\n                "Excel: no rows matched work_area=\'%s\'. Falling back to process-only match for process=\'%s\' (%d rows).",\n                work_area, process_name, len(process_only_matches),\n            )\n            return process_only_matches\n\n        log.info("Excel: loaded 0 rows for process=\'%s\', work_area=\'%s\'", process_name, work_area)\n        return []\n    finally:\n        wb.close()\n\n\ndef load_unique_process_names_from_excel(\n    file_path: str,\n    sheet_name: str = "Parameters",\n) -> List[str]:\n    """Load distinct process names from first Excel column."""\n    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)\n    try:\n        if sheet_name not in wb.sheetnames:\n            raise ValueError(\n                f"Sheet \'{sheet_name}\' not found in \'{file_path}\'. Available sheets: {wb.sheetnames}"\n            )\n\n        ws = wb[sheet_name]\n        seen = set()\n        names: List[str] = []\n        for row in ws.iter_rows(min_row=2, values_only=True):\n            if not row:\n                continue\n            raw = str(row[0] or "").strip()\n            norm = normalize_label(raw)\n            if not raw or not norm or norm in seen:\n                continue\n            seen.add(norm)\n            names.append(raw)\n        return names\n    finally:\n        wb.close()\n\n\ndef resolve_process_name_with_excel(\n    requested_process_name: str,\n    excel_file_path: str,\n    sheet_name: str = "Parameters",\n    min_autocorrect_score: float = 0.80,\n) -> str:\n    """\n    Validate/autocorrect process name against Excel first column.\n    Returns best process name to be used for UI search/entry.\n    """\n    requested = (requested_process_name or "").strip()\n    if not requested:\n        return requested\n\n    try:\n        excel_processes = load_unique_process_names_from_excel(excel_file_path, sheet_name)\n    except Exception as exc:\n        log.warning("  ⚠️ Excel process-name validation skipped: %s", exc)\n        return requested\n\n    if not excel_processes:\n        log.warning("  ⚠️ Excel process-name validation skipped: no process names found in first column")\n        return requested\n\n    # Exact/contains style match first\n    for proc in excel_processes:\n        if values_match(proc, requested):\n            if normalize_label(proc) != normalize_label(requested):\n                log.info("  ✅ Process autocorrect (exact/contains): \'%s\' -> \'%s\'", requested, proc)\n            else:\n                log.info("  ✅ Process name validated with Excel: \'%s\'", requested)\n            return proc\n\n    # Fuzzy best match fallback\n    best_name = ""\n    best_score = 0.0\n    for proc in excel_processes:\n        score = similarity_score(proc, requested)\n        if score > best_score:\n            best_score = score\n            best_name = proc\n\n    if best_name and best_score >= min_autocorrect_score:\n        log.info(\n            "  ✅ Process autocorrect (fuzzy %.2f): \'%s\' -> \'%s\'",\n            best_score,\n            requested,\n            best_name,\n        )\n        return best_name\n\n    log.warning(\n        "  ⚠️ Process name \'%s\' not confidently matched in Excel (best=\'%s\', score=%.2f). Using original input.",\n        requested,\n        best_name or "(none)",\n        best_score,\n    )\n    return requested\n\n\ndef discover_screen_labels(page: Page) -> List[str]:\n    label_selectors = [\n        "label.af_outputLabel_label",\n        "label[id$=\'::label\']",\n        "oj-label label",\n        ".oj-label-group label",\n        "form label",\n    ]\n\n    ignore_norm_labels = {\n        normalize_label(x)\n        for x in [\n            "Type", "Job", "Job Set", "Name", "Description", "Saved Search", "View",\n            "Flat List", "Hierarchy", "Status", "Schedule Start",\n            "Do you want to review the parameters before submitting?",\n            "Notify me when this process ends", "Schedule", "Submission Notes",\n            "Use Existing Index",\n        ]\n    }\n\n    for selector in label_selectors:\n        elements = page.query_selector_all(selector)\n        if not elements:\n            continue\n        labels: List[str] = []\n        for el in elements:\n            text = (el.inner_text() or "").strip().rstrip(":")\n            if text and normalize_label(text) not in ignore_norm_labels:\n                labels.append(text)\n        if labels:\n            log.info("Screen discovery: found %d labels via \'%s\'", len(labels), selector)\n            return labels\n    log.warning("Screen discovery: no labels found with known selectors.")\n    return []\n\n\ndef build_discovered_label_map(discovered_labels: List[str]) -> Dict[str, str]:\n    label_map: Dict[str, str] = {}\n    for label in discovered_labels:\n        raw = str(label).strip()\n        if raw:\n            label_map[normalize_label(raw)] = raw\n    return label_map\n\n\ndef generate_runtime_fill_plan(\n    process_name: str,\n    discovered_labels: List[str],\n    input_parameters: List[Parameter],\n    strict_mode: bool = False,\n    validation_mode: Literal["STRICT", "WARN_AND_HOLD", "BEST_EFFORT"] = "WARN_AND_HOLD",\n) -> FillPlan:\n    discovered_map = build_discovered_label_map(discovered_labels)\n    input_map: Dict[str, str] = {}\n    original_input_labels: Dict[str, str] = {}\n\n    for item in input_parameters:\n        label = str(item.get("label", "")).strip()\n        value = str(item.get("value", ""))\n        if not label:\n            continue\n        norm = normalize_label(label)\n        input_map[norm] = value\n        original_input_labels[norm] = label\n\n    if not discovered_map:\n        return {\n            "action": "SUBMIT_ONLY",\n            "parameters_to_fill": [],\n            "extra_input_labels": [],\n            "submission_allowed": True,\n            "message": f"Process \'{process_name}\' has no parameters. Submit directly.",\n        }\n\n    parameters_to_fill: List[Parameter] = []\n    for norm_label, original_label in discovered_map.items():\n        best_match_norm = None\n        best_match_score = 0.0\n        if norm_label in input_map:\n            value = input_map[norm_label]\n            if value.strip():\n                parameters_to_fill.append({"label": original_label, "value": value})\n            continue\n\n        for input_norm_label, value in input_map.items():\n            score = similarity_score(norm_label, input_norm_label)\n            if score > best_match_score:\n                best_match_score = score\n                best_match_norm = input_norm_label\n            if values_match(norm_label, input_norm_label) and value.strip():\n                parameters_to_fill.append({"label": original_label, "value": value})\n                best_match_norm = None\n                break\n\n        if best_match_norm and best_match_score >= 0.78:\n            v = input_map.get(best_match_norm, "")\n            if v.strip():\n                parameters_to_fill.append({"label": original_label, "value": v})\n\n    extra_input_labels = [original_input_labels[norm] for norm in input_map if norm not in discovered_map]\n    return {\n        "action": "FILL_AND_SUBMIT",\n        "parameters_to_fill": parameters_to_fill,\n        "extra_input_labels": extra_input_labels,\n        "submission_allowed": True,\n        "message": f"Auto-fill {len(parameters_to_fill)} parameter(s). Extra input ignored: {extra_input_labels or \'none\'}.",\n    }\n\n\ndef fill_parameters_on_screen(page: Page, parameters_to_fill: List[Parameter]) -> Dict[str, str]:\n    results: Dict[str, str] = {}\n    for item in parameters_to_fill:\n        label_text = str(item.get("label", "")).strip()\n        raw_value = str(item.get("value", ""))\n        value = sanitize_parameter_value(raw_value)\n        if not label_text:\n            continue\n        try:\n            if value != raw_value.strip():\n                log.info("  ℹ️ Sanitized parameter \'%s\' value: \'%s\' -> \'%s\'", label_text, raw_value, value)\n            label_el = page.get_by_text(label_text, exact=False).first\n            for_attr = label_el.get_attribute("for") if label_el else None\n            if for_attr:\n                input_el = page.locator(f\'[id="{for_attr}"]\').first\n            else:\n                input_el = label_el.locator(\n                    "xpath=following::input[1] | xpath=following::textarea[1] | xpath=following::select[1]"\n                ).first\n\n            input_el.wait_for(state="visible", timeout=6_000)\n            input_el.click()\n            tag = (input_el.evaluate("el => el.tagName.toLowerCase()") or "").strip()\n            if tag in {"input", "textarea"}:\n                input_el.fill("")\n                input_el.type(value, delay=50)\n            else:\n                input_el.evaluate(\n                    """(el, v) => {\n                        const target = el.matches(\'input,textarea\') ? el : el.querySelector(\'input,textarea\');\n                        if (!target) throw new Error(\'No input/textarea found for parameter field\');\n                        target.focus();\n                        target.value = \'\';\n                        target.value = v;\n                        target.dispatchEvent(new Event(\'input\', { bubbles: true }));\n                        target.dispatchEvent(new Event(\'change\', { bubbles: true }));\n                    }""",\n                    value,\n                )\n            input_el.press("Tab")\n            results[label_text] = f"FILLED (value=\'{value}\')"\n        except Exception as exc:\n            results[label_text] = f"FAILED (value=\'{value}\'): {exc}"\n            log.error("Failed to fill \'%s\': %s", label_text, exc)\n    return results\n\n\ndef run_parameter_handling(\n    page: Page,\n    process_name: str,\n    job_request_id: str,\n    work_area: str,\n    page_name: str = "Scheduled Processes",\n    input_parameters: Optional[List[Parameter]] = None,\n    excel_path: Optional[str] = None,\n    sheet_name: str = "Parameters",\n    strict_mode: bool = False,\n    validation_mode: Literal["STRICT", "WARN_AND_HOLD", "BEST_EFFORT"] = "WARN_AND_HOLD",\n) -> Dict[str, object]:\n    if input_parameters is None:\n        if not excel_path:\n            raise ValueError("Provide either input_parameters or excel_path.")\n        input_parameters = load_parameters_from_excel(\n            file_path=excel_path,\n            process_name=process_name,\n            work_area=work_area,\n            sheet_name=sheet_name,\n        )\n\n    discovered_labels = discover_screen_labels(page)\n    plan = generate_runtime_fill_plan(\n        process_name=process_name,\n        discovered_labels=discovered_labels,\n        input_parameters=input_parameters,\n        strict_mode=strict_mode,\n        validation_mode=validation_mode,\n    )\n\n    fill_results: Dict[str, str] = {}\n    if plan["action"] == "FILL_AND_SUBMIT" and plan["parameters_to_fill"]:\n        fill_results = fill_parameters_on_screen(page, plan["parameters_to_fill"])\n\n    context_key = f"{work_area}|{page_name}|{process_name}"\n    return {\n        "job_request_id": job_request_id,\n        "process_name": process_name,\n        "work_area": work_area,\n        "page_name": page_name,\n        "context_key": context_key,\n        "discovered_labels": discovered_labels,\n        "input_labels": [str(p.get("label", "")) for p in input_parameters],\n        "fill_plan_action": plan["action"],\n        "submission_allowed": plan["submission_allowed"],\n        "parameters_to_fill": plan["parameters_to_fill"],\n        "extra_input_labels": plan["extra_input_labels"],\n        "fill_results": fill_results,\n        "message": plan["message"],\n    }\n\n\n# ─────────────────────────────────────────────────────────────────────────────\n# STEP 1 – LOGIN\n# ─────────────────────────────────────────────────────────────────────────────\n\ndef login(page: Page):\n    log.info("🔐 Step 1 – Login")\n    page.goto(BASE_URL, timeout=60_000)\n    page.wait_for_load_state("domcontentloaded")\n    time.sleep(1)\n\n    LOGIN_FORMS = [\n        ("#idcs-signin-basic-signin-form-username", "#idcs-signin-basic-signin-form-password", ".oj-button-text:has-text(\'Sign In\')"),\n        ("input#userid", "input#password", "button#btnActive"),\n        ("#userid", "#password", "#btnActive"),\n        (\'input[name="username"]\', \'input[name="password"]\', \'button[type="submit"]\'),\n        (\'input[type="email"]\', \'input[type="password"]\', \'button[type="submit"]\'),\n    ]\n\n    logged_in = False\n    for u_sel, p_sel, s_sel in LOGIN_FORMS:\n        try:\n            page.wait_for_selector(u_sel, timeout=5_000)\n            if not page.is_visible(u_sel):\n                continue\n            page.fill(u_sel, USERNAME)\n            page.fill(p_sel, PASSWORD)\n            page.click(s_sel)\n            logged_in = True\n            log.info(f"  ✓ Login form matched: {u_sel}")\n            break\n        except Exception:\n            continue\n\n    if not logged_in:\n        inputs = page.query_selector_all("input")\n        if len(inputs) >= 2:\n            inputs[0].fill(USERNAME)\n            inputs[1].fill(PASSWORD)\n            inputs[1].press("Enter")\n        else:\n            raise RuntimeError("Cannot identify login fields")\n\n    log.info("  ⏳ Waiting for Navigator icon…")\n    page.wait_for_selector(\'svg[aria-label="Navigator"], [aria-label="Navigator"]\', timeout=120_000)\n    wait_busy(page)\n    log.info("  ✅ Login successful")\n\n\n# ─────────────────────────────────────────────────────────────────────────────\n# STEP 2 – CLICK NAVIGATOR (HAMBURGER MENU)\n# ─────────────────────────────────────────────────────────────────────────────\n\ndef click_navigator(page: Page):\n    """Click the top-left hamburger/Navigator menu."""\n    log.info("🧭 Step 2 – Click Navigator (hamburger menu)")\n    \n    NAVIGATOR_SELECTORS = [\n        \'a#pt1\\\\:_UISmmLink\',                    # Exact ID from your HTML\n        \'a[title="Navigator"]\',                   # Title attribute\n        \'a.TabletNavigatorIcon\',                  # Class-based fallback\n        \'svg[aria-label="Navigator"]\',            # SVG fallback\n    ]\n    \n    clicked = False\n    for sel in NAVIGATOR_SELECTORS:\n        try:\n            page.wait_for_selector(sel, state="visible", timeout=15_000)\n            nav_link = page.locator(sel).first\n            \n            # ADF onclick blocker — use mousedown sequence\n            js_mousedown(page, nav_link, "Navigator hamburger")\n            clicked = True\n            log.info(f"  ✓ Navigator clicked ({sel})")\n            break\n        except Exception as e:\n            log.debug(f"  Selector \'{sel}\' failed: {e}")\n            continue\n    \n    if not clicked:\n        snap(page, "navigator_click_failed")\n        debug_dom_info(page, "Navigator click failed")\n        raise RuntimeError("Could not click Navigator (hamburger) menu")\n    \n    wait_busy(page)\n    time.sleep(1.0)  # Let menu animation complete\n    log.info("  ✅ Navigator menu opened")\n\n\n# ─────────────────────────────────────────────────────────────────────────────\n# STEP 3 – EXPAND TOOLS SECTION\n# ─────────────────────────────────────────────────────────────────────────────\n\ndef expand_tools_section(page: Page):\n    """Scroll down and expand the Tools section in the Navigator."""\n    log.info("🔧 Step 3 – Expand Tools section")\n    \n    # The Tools header div that expands/collapses\n    TOOLS_HEADER_SELECTORS = [\n        \'div#pt1\\\\:_UISnvr\\\\:0\\\\:nvgpgl2_groupNode_tools\',  # Exact ID\n        \'div.navmenu-header:has-text("Tools")\',              # Class + text\n        \'div[title="Tools"].navmenu-header\',                 # Title + class\n    ]\n    \n    expanded = False\n    for sel in TOOLS_HEADER_SELECTORS:\n        try:\n            tools_header = page.locator(sel).first\n            tools_header.wait_for(state="visible", timeout=15_000)\n            tools_header.scroll_into_view_if_needed()\n            time.sleep(0.3)\n            \n            # Check if already expanded (child items visible)\n            is_expanded = page.evaluate("""(sel) => {\n                const header = document.querySelector(sel);\n                if (!header) return false;\n                const parent = header.closest(\'[id*="groupNode_tools"]\');\n                if (!parent) return false;\n                const content = parent.querySelector(\'[id*="nvgpgl3_groupNode_tools"]\');\n                return content && content.style.visibility !== \'hidden\';\n            }""", sel)\n            \n            if is_expanded:\n                log.info("  ✓ Tools section already expanded")\n                expanded = True\n                break\n            \n            # Click the header to expand (not the expand arrow itself)\n            js_mousedown(page, tools_header, "Tools section header")\n            time.sleep(0.8)\n            wait_busy(page)\n            \n            # Verify expansion\n            is_now_expanded = page.evaluate("""() => {\n                const content = document.querySelector(\'[id*="nvgpgl3_groupNode_tools"]\');\n                return content && content.style.visibility !== \'hidden\';\n            }""")\n            \n            if is_now_expanded:\n                log.info("  ✓ Tools section expanded")\n                expanded = True\n                break\n                \n        except Exception as e:\n            log.debug(f"  Tools expansion via \'{sel}\' failed: {e}")\n            continue\n    \n    if not expanded:\n        # Fallback: click the expand arrow directly\n        try:\n            expand_arrow = page.locator(\'a#pt1\\\\:_UISnvr\\\\:0\\\\:nvgcil_groupNode_tools\').first\n            expand_arrow.wait_for(state="visible", timeout=10_000)\n            js_mousedown(page, expand_arrow, "Tools expand arrow")\n            time.sleep(0.8)\n            wait_busy(page)\n            log.info("  ✓ Tools expanded via arrow click")\n            expanded = True\n        except Exception as e:\n            snap(page, "tools_expand_failed")\n            debug_dom_info(page, "Tools expand failed")\n            raise RuntimeError(f"Could not expand Tools section: {e}")\n    \n    log.info("  ✅ Tools section ready")\n\n\n# ─────────────────────────────────────────────────────────────────────────────\n# STEP 4 – CLICK SCHEDULED PROCESSES\n# ─────────────────────────────────────────────────────────────────────────────\n\ndef click_scheduled_processes(page: Page):\n    """Click the Scheduled Processes link under Tools."""\n    log.info("📅 Step 4 – Click Scheduled Processes")\n    \n    SCHED_PROC_SELECTORS = [\n        \'a#pt1\\\\:_UISnvr\\\\:0\\\\:nv_itemNode_tools_scheduled_processes_fuse_plus\',  # Exact ID\n        \'a[title="Scheduled Processes"]\',                                          # Title\n        \'a.x3a2:has-text("Scheduled Processes")\',                                  # Class + text\n    ]\n    \n    clicked = False\n    for sel in SCHED_PROC_SELECTORS:\n        try:\n            link = page.locator(sel).first\n            link.wait_for(state="visible", timeout=15_000)\n            link.scroll_into_view_if_needed()\n            time.sleep(0.2)\n            \n            js_mousedown(page, link, "Scheduled Processes link")\n            clicked = True\n            log.info(f"  ✓ Scheduled Processes clicked ({sel})")\n            break\n        except Exception as e:\n            log.debug(f"  Selector \'{sel}\' failed: {e}")\n            continue\n    \n    if not clicked:\n        snap(page, "scheduled_processes_click_failed")\n        debug_dom_info(page, "Scheduled Processes click failed")\n        raise RuntimeError("Could not click Scheduled Processes link")\n    \n    wait_busy(page, timeout=45_000)  # Page load can be slow\n    time.sleep(2.0)\n    log.info("  ✓ Scheduled Processes page loaded")\n\n\ndef _normalize_work_area_text(value: str) -> str:\n    """Normalize UI breadcrumb/work-area text to canonical form for Excel matching."""\n    txt = " ".join((value or "").split()).strip()\n    if not txt:\n        return ""\n    txt = txt.replace("Setup and Maintenance", "Setup")\n    txt = txt.replace(" > ", ": ").replace("/", ": ")\n    txt = txt.replace(" :", ":").replace(": ", ":")\n\n    # If breadcrumb contains many parts, prefer the segment starting with Setup.\n    low = txt.lower()\n    if "setup" in low and ":" in txt:\n        parts = [p.strip() for p in txt.split(":") if p.strip()]\n        if len(parts) >= 2:\n            # Keep only first two parts e.g. Setup:Procurement\n            txt = f"{parts[0]}:{parts[1]}"\n    return txt\n\n\ndef capture_work_area_from_breadcrumb(page: Page) -> str:\n    """Capture work area from Oracle breadcrumb (example: \'Setup:Procurement\')."""\n    breadcrumb_selectors = [\n        \'[id*="_UISbc" i]\',\n        \'[id*="breadcrumb" i]\',\n        \'[aria-label*="breadcrumb" i]\',\n        \'nav[aria-label*="breadcrumb" i]\',\n        \'.breadcrumb\',\n        \'a:has-text("Setup")\',\n    ]\n\n    # 1) Try direct breadcrumb containers/selectors.\n    for sel in breadcrumb_selectors:\n        try:\n            txt = page.locator(sel).first.inner_text(timeout=2500).strip()\n            norm = _normalize_work_area_text(txt)\n            if norm and ("setup" in norm.lower() or "procurement" in norm.lower()):\n                log.info("  ℹ️ Captured breadcrumb work area: \'%s\'", txt)\n                log.info("  ℹ️ Normalized work area for Excel: \'%s\'", norm)\n                return norm\n        except Exception:\n            continue\n\n    # 2) Fallback: scan full page text and regex out setup/work-area phrase.\n    try:\n        body_text = page.evaluate("() => (document.body && document.body.innerText) ? document.body.innerText : \'\'")\n        body_text = " ".join((body_text or "").split())\n        import re\n\n        patterns = [\n            r\'(Setup\\s*(?:and\\s*Maintenance)?\\s*[:>]\\s*[A-Za-z][A-Za-z\\s&-]+)\',\n            r\'(Setup\\s*[:>]\\s*Procurement)\',\n        ]\n        for pat in patterns:\n            m = re.search(pat, body_text, flags=re.IGNORECASE)\n            if m:\n                raw = m.group(1).strip()\n                norm = _normalize_work_area_text(raw)\n                if norm:\n                    log.info("  ℹ️ Captured breadcrumb work area (text scan): \'%s\'", raw)\n                    log.info("  ℹ️ Normalized work area for Excel: \'%s\'", norm)\n                    return norm\n    except Exception:\n        pass\n\n    fallback_work_area = "Tools"\n    log.warning("  ⚠️ Breadcrumb work area not found. Using fallback WORK_AREA=\'%s\'", fallback_work_area)\n    return _normalize_work_area_text(fallback_work_area)\n\n\n# ─────────────────────────────────────────────────────────────────────────────\n# STEP 5 – CLICK SCHEDULE NEW PROCESS\n# ─────────────────────────────────────────────────────────────────────────────\n\ndef click_schedule_new_process(page: Page):\n    """Click the \'Schedule New Process\' button."""\n    log.info("➕ Step 5 – Click Schedule New Process")\n    \n    # The button is an <a> with role="button" and onclick blocker\n    # It contains a span.xrk with the text\n    try:\n        # Try to find the anchor wrapper first\n        new_proc_btn = page.locator(\n            \'a[role="button"]:has-text("Schedule New Process"), \'\n            \'a.xrg:has-text("Schedule New Process")\'\n        ).first\n        new_proc_btn.wait_for(state="visible", timeout=20_000)\n        js_mousedown(page, new_proc_btn, "Schedule New Process button")\n        log.info("  ✓ Schedule New Process clicked (anchor)")\n        \n    except Exception as e:\n        log.debug(f"  Anchor click failed: {e}")\n        # Fallback: find the span and click its parent\n        try:\n            span = page.locator(\'span.xrk:has-text("Schedule New Process")\').first\n            span.wait_for(state="visible", timeout=10_000)\n            parent = span.evaluate_handle("el => el.parentElement")\n            js_mousedown(page, parent, "Schedule New Process (via span parent)")\n            log.info("  ✓ Schedule New Process clicked (via span)")\n        except Exception as e2:\n            log.debug(f"  Span fallback failed: {e2}")\n            # Last resort: keyboard navigation\n            page.keyboard.press("Tab")\n            time.sleep(0.2)\n            page.keyboard.press("Enter")\n            log.info("  ✓ Fallback: used keyboard to activate button")\n    \n    wait_busy(page, timeout=30_000)\n    time.sleep(1.5)\n    log.info("  ✅ Schedule New Process screen loaded")\n\n\ndef select_process_in_schedule_dialog_direct(page: Page, process_name: str, direct: bool = False):\n    """\n    Fill Name directly in Schedule New Process, verify value, then keep dialog ready for OK.\n\n    If `direct` is True, the given process_name is used as-is without looking it up in\n    Additional Info. Otherwise, we resolve it from Additional Info table first.\n    """\n    requested_process = (process_name or "").strip()\n\n    if direct:\n        target_process = requested_process\n        if not target_process:\n            raise RuntimeError("Process name not provided for direct selection")\n    else:\n        additional_info_path = str((Path.cwd() / ADDITIONAL_INFO_EXCEL_FILE).resolve())\n        try:\n            parsed_entries = load_process_params_from_additional_info(additional_info_path)\n            if not parsed_entries:\n                raise RuntimeError("No valid ESS Job entries found in Additional Info")\n\n            best_entry = find_best_additional_info_entry(parsed_entries, requested_process)\n            if not best_entry or not str(best_entry.get("process_name", "")).strip():\n                raise RuntimeError("No valid ESS Job process name found in Additional Info")\n\n            target_process = str(best_entry.get("process_name", "")).strip()\n            log.info(\n                "  ✅ Process from Additional Info: \'%s\' (sheet=%s row=%s)",\n                target_process,\n                best_entry.get("sheet", ""),\n                best_entry.get("row", ""),\n            )\n        except Exception as exc:\n            raise RuntimeError(f"Additional Info process parse failed: {exc}")\n\n    if not target_process.strip():\n        raise RuntimeError("No process name resolved for Schedule New Process Name field")\n\n    log.info("🧩 Step 6 – Enter process name directly in Schedule New Process")\n    snap(page, "s4_schedule_new_process_blank")\n\n    name_input_selectors = [\n        \'input[role="combobox"][name*="selectOneChoice2"]\',\n        \'input[id*="selectOneChoice2"][id$="::content"]\',\n        \'input[aria-autocomplete="list"][role="combobox"]\',\n        \'xpath=//label[contains(normalize-space(),"Name")]/following::input[1]\',\n    ]\n\n    name_input = None\n    for sel in name_input_selectors:\n        try:\n            el = page.locator(sel).first\n            el.wait_for(state="visible", timeout=10_000)\n            name_input = el\n            break\n        except Exception:\n            continue\n\n    if name_input is None:\n        raise RuntimeError("Could not find Name input on Schedule New Process dialog")\n\n    name_input.click()\n    name_input.fill("")\n    name_input.type(target_process, delay=40)\n    page.keyboard.press("Tab")\n    wait_busy(page)\n    time.sleep(0.8)\n\n    # Verify entered value\n    typed_value = (name_input.input_value() or "").strip()\n    if target_process.lower() not in typed_value.lower():\n        # one retry\n        name_input.click()\n        name_input.fill("")\n        name_input.type(target_process, delay=40)\n        page.keyboard.press("Tab")\n        wait_busy(page)\n        time.sleep(0.8)\n        typed_value = (name_input.input_value() or "").strip()\n\n    if target_process.lower() not in typed_value.lower():\n        log.warning(\n            "  ⚠️ Name verification failed. Expected contains \'%s\', actual \'%s\'",\n            target_process,\n            typed_value,\n        )\n        log.warning("  ℹ️ Description: no process available")\n        return ""\n\n    # Optional safety check: description may load with delay in some pods.\n    desc_text = ""\n    desc_locator = page.locator(\'xpath=//label[contains(normalize-space(),"Description")]/following::*[1]\').first\n    for _ in range(6):\n        try:\n            wait_busy(page, timeout=3_000)\n        except Exception:\n            pass\n        try:\n            desc_text = (desc_locator.inner_text(timeout=1_000) or "").strip()\n        except Exception:\n            desc_text = ""\n        if desc_text:\n            break\n        time.sleep(0.5)\n\n    if not desc_text:\n        log.warning("  ⚠️ Process description is empty after name entry; continuing because Name matched.")\n    else:\n        log.info("  ℹ️ Description resolved: \'%s\'", desc_text)\n\n    log.info("  ✅ Name field populated and verified: \'%s\'", typed_value)\n    snap(page, "s2_process_selected_in_schedule_dialog")\n    return target_process\n\n\ndef open_process_details_from_schedule_dialog(page: Page):\n    """Click OK in Schedule New Process dialog to open Process Details (screenshot3)."""\n    log.info("🧩 Step 8 – Open Process Details (screenshot2 -> screenshot3)")\n\n    # Scope all actions to the Schedule New Process popup/dialog first.\n    schedule_dialog = page.locator(\'[role="dialog"]:has-text("Schedule New Process")\').first\n    try:\n        schedule_dialog.wait_for(state="visible", timeout=8_000)\n    except Exception:\n        # keep backward-compatible behavior if role/title differs in some pods\n        schedule_dialog = None\n\n    clicked = False\n    ok_selectors = [\n        \'button:has-text("OK")\',\n        \'a[role="button"]:has-text("OK")\',\n        \'a:has-text("OK")\',\n        \'span:has-text("OK")\',\n        \'xpath=//*[normalize-space()="OK"]\',\n    ]\n\n    # Try multiple attempts because ADF overlays can intercept clicks.\n    for attempt in range(3):\n        try:\n            wait_busy(page, timeout=8_000)\n        except Exception:\n            pass\n\n        # Prefer dialog-scoped OKs first.\n        scopes = [schedule_dialog] if schedule_dialog else []\n        scopes.append(page)\n\n        for scope in scopes:\n            for sel in ok_selectors:\n                try:\n                    candidates = scope.locator(sel)\n                    count = min(candidates.count(), 8)\n                    for idx in range(count - 1, -1, -1):\n                        try:\n                            el = candidates.nth(idx)\n                            if not el.is_visible(timeout=600):\n                                continue\n\n                            # Click cascade: normal click -> js mousedown -> focus+Enter\n                            clicked_this = False\n                            try:\n                                el.click(timeout=1_800)\n                                clicked_this = True\n                            except Exception:\n                                try:\n                                    js_mousedown(page, el, "Schedule New Process OK")\n                                    clicked_this = True\n                                except Exception:\n                                    try:\n                                        el.focus(timeout=500)\n                                        page.keyboard.press("Enter")\n                                        clicked_this = True\n                                    except Exception:\n                                        clicked_this = False\n\n                            if not clicked_this:\n                                continue\n\n                            # Success signals for Process Details.\n                            success = False\n                            success_selectors = [\n                                \'text="Process Details"\',\n                                \'text="Basic Options"\',\n                                \'text="Parameters"\',\n                            ]\n                            for s in success_selectors:\n                                try:\n                                    page.wait_for_selector(s, timeout=3_000)\n                                    success = True\n                                    break\n                                except Exception:\n                                    continue\n\n                            if not success:\n                                try:\n                                    page.get_by_role("button", name="Submit").first.wait_for(\n                                        state="visible", timeout=3_000\n                                    )\n                                    success = True\n                                except Exception:\n                                    success = False\n\n                            if success:\n                                clicked = True\n                                break\n                        except Exception:\n                            continue\n                    if clicked:\n                        break\n                except Exception:\n                    continue\n            if clicked:\n                break\n\n        if clicked:\n            break\n\n        time.sleep(0.6)\n\n    if not clicked:\n        snap(page, "s8_ok_click_failed")\n        dump_html(page, "s8_ok_click_failed")\n        raise RuntimeError("Could not click OK button / trigger navigation to Process Details")\n\n    # Robust page-ready detection across UI variants.\n    ready = False\n    ready_selectors = [\n        \'text="Process Details"\',\n        \'text="Basic Options"\',\n        \'text="Parameters"\',\n    ]\n    for sel in ready_selectors:\n        try:\n            page.wait_for_selector(sel, timeout=8_000)\n            ready = True\n            break\n        except Exception:\n            continue\n\n    if not ready:\n        try:\n            page.get_by_role("button", name="Submit").first.wait_for(state="visible", timeout=8_000)\n            ready = True\n        except Exception:\n            ready = False\n\n    if not ready:\n        raise RuntimeError("OK clicked but Process Details indicators not found")\n\n    wait_busy(page)\n    # step-3 checkpoint already validated by ready-indicator checks above\n    snap(page, "s3_process_details_opened")\n\n\ndef extract_selected_process_name_from_process_details(page: Page) -> str:\n    """Read process name shown on Process Details header section."""\n    name_selectors = [\n        \'xpath=//label[contains(normalize-space(),"Name")]/following::*[1]\',\n        \'xpath=//*[contains(normalize-space(),"Process Details")]/following::label[contains(normalize-space(),"Name")]/following::*[1]\',\n        \'xpath=//span[contains(@id,"name") or contains(@class,"name")][normalize-space()]\',\n    ]\n    for sel in name_selectors:\n        try:\n            txt = page.locator(sel).first.inner_text(timeout=4000).strip()\n            if txt:\n                return " ".join(txt.split())\n        except Exception:\n            continue\n    return ""\n\n\ndef fill_parameters_for_selected_process_and_submit(page: Page, process_name: str, work_area: str):\n    """Extract process name + parameter labels, map with Excel, fill, then submit."""\n    log.info("🧾 Step 9 – Extract, map with Excel, fill parameters, and submit")\n\n    remembered_process_name = (process_name or "").strip()\n    ui_process_name = extract_selected_process_name_from_process_details(page)\n    effective_process_name = remembered_process_name or ui_process_name\n    if not effective_process_name:\n        raise RuntimeError("Could not determine process name for parameter mapping")\n    log.info("  ℹ️ Process remembered from selection: \'%s\'", remembered_process_name or "(none)")\n    log.info("  ℹ️ Process extracted from Process Details: \'%s\'", ui_process_name or "(none)")\n    log.info("  ℹ️ Process used for Excel mapping: \'%s\'", effective_process_name)\n    log.info("  ℹ️ Work area used for Excel mapping: \'%s\'", work_area)\n\n    excel_path = str((Path.cwd() / "ESS_parameters_structure.xlsx").resolve())\n    additional_info_path = str((Path.cwd() / ADDITIONAL_INFO_EXCEL_FILE).resolve())\n\n    additional_info_input_params: List[Parameter] = []\n    try:\n        parsed_entries = load_process_params_from_additional_info(additional_info_path)\n        best_entry = find_best_additional_info_entry(parsed_entries, effective_process_name)\n        if best_entry:\n            additional_info_input_params = [\n                {\n                    "label": best_entry.get("param_label", "").strip(),\n                    "value": best_entry.get("param_value", "").strip(),\n                }\n            ]\n            log.info(\n                "  ✅ Parameter from Additional Info: \'%s\' = \'%s\' (sheet=%s row=%s)",\n                best_entry.get("param_label", ""),\n                best_entry.get("param_value", ""),\n                best_entry.get("sheet", ""),\n                best_entry.get("row", ""),\n            )\n    except Exception as exc:\n        log.warning("  ⚠️ Additional Info parameter parse skipped: %s", exc)\n\n    if additional_info_input_params:\n        result = run_parameter_handling(\n            page=page,\n            process_name=effective_process_name,\n            job_request_id="AUTO",\n            work_area=work_area,\n            page_name="Scheduled Processes",\n            input_parameters=additional_info_input_params,\n            strict_mode=False,\n            validation_mode="WARN_AND_HOLD",\n        )\n    else:\n        result = run_parameter_handling(\n            page=page,\n            process_name=effective_process_name,\n            job_request_id="AUTO",\n            work_area=work_area,\n            page_name="Scheduled Processes",\n            excel_path=excel_path,\n            sheet_name="Parameters",\n            strict_mode=False,\n            validation_mode="WARN_AND_HOLD",\n        )\n\n    discovered = result.get("discovered_labels", [])\n    mapped = result.get("parameters_to_fill", [])\n    extras = result.get("extra_input_labels", [])\n    log.info("  📋 UI parameter labels discovered: %s", discovered)\n    log.info("  🔗 Parameters mapped from Excel: %s", mapped)\n    if extras:\n        log.warning("  ⚠️ Excel labels not present on UI (ignored): %s", extras)\n    log.info("  🧪 Fill results: %s", result.get("fill_results", {}))\n\n    # Explicitly ignore/take no action on checkbox parameter from screenshot flow.\n    log.info("  ℹ️ Ignoring checkbox: \'Use Existing Index\'")\n\n    # Submit (must click the Process Details header Submit shown near Process Options/Advanced)\n    try:\n        log.info("  ➤ Clicking Process Details header Submit")\n\n        process_details = page.locator(\'[role="dialog"]:has-text("Process Details")\').first\n        submit_candidates = [\n            process_details.locator(\'xpath=.//button[normalize-space()="Submit"]\'),\n            process_details.locator(\'xpath=.//*[contains(normalize-space(),"Process Options")]/following::*[normalize-space()="Submit"][1]\'),\n            process_details.get_by_role("button", name="Submit"),\n            page.locator(\'xpath=//*[contains(normalize-space(),"Process Options")]/following::*[normalize-space()="Submit"][1]\'),\n            page.get_by_role("button", name="Submit"),\n        ]\n\n        submit_btn = None\n        for group in submit_candidates:\n            try:\n                if group.count() <= 0:\n                    continue\n                limit = min(group.count(), 5)\n                for idx in range(limit):\n                    cand = group.nth(idx)\n                    if cand.is_visible(timeout=1000):\n                        submit_btn = cand\n                        break\n                if submit_btn is not None:\n                    break\n            except Exception:\n                continue\n\n        if submit_btn is None:\n            raise RuntimeError("Header Submit button not found in Process Details")\n\n        clicked = False\n        try:\n            submit_btn.click(timeout=3_000)\n            clicked = True\n        except Exception:\n            pass\n\n        if not clicked:\n            try:\n                js_mousedown(page, submit_btn, "Process Details header Submit")\n                clicked = True\n            except Exception:\n                pass\n\n        if not clicked:\n            submit_btn.focus(timeout=1_000)\n            page.keyboard.press("Enter")\n            clicked = True\n\n        if not clicked:\n            raise RuntimeError("Unable to click Process Details header Submit")\n\n        wait_busy(page, timeout=30_000)\n        time.sleep(1.0)\n\n        # success confirmation: queue/info message OR dialog close OR post-submit indicators.\n        submit_confirmed = False\n        success_checks = [\n            \'text="queued up for submission"\',\n            \'text="This process will be queued up for submission"\',\n            \'text="Request ID"\',\n            \'text="Confirmation"\',\n        ]\n        for sel in success_checks:\n            try:\n                page.locator(sel).first.wait_for(state="visible", timeout=4_000)\n                submit_confirmed = True\n                break\n            except Exception:\n                continue\n\n        if not submit_confirmed:\n            try:\n                process_details.wait_for(state="hidden", timeout=4_000)\n                submit_confirmed = True\n            except Exception:\n                pass\n\n        if submit_confirmed:\n            log.info("  ✅ Submit success confirmation detected")\n        else:\n            log.info("  ✅ Submit clicked (no explicit confirmation selector detected)")\n\n        # Post-submit confirmation popup: click OK automatically.\n        try:\n            log.info("  ➤ Checking for post-submit confirmation popup")\n            ok_btn = None\n            popup_candidates = [\n                page.locator(\'[role="dialog"]:has-text("Confirmation")\').first,\n                page.locator(\'[role="dialog"]:has-text("Process Details")\').first,\n                page.locator(\'[role="dialog"]\').last,\n            ]\n\n            # Try dialog-scoped OK first\n            for popup in popup_candidates:\n                try:\n                    popup.wait_for(state="visible", timeout=2_000)\n                    cand = popup.get_by_role("button", name="OK").first\n                    if cand.count() > 0 and cand.is_visible(timeout=1_000):\n                        ok_btn = cand\n                        break\n                    cand2 = popup.locator(\'xpath=.//*[normalize-space()="OK"]\').first\n                    if cand2.count() > 0 and cand2.is_visible(timeout=1_000):\n                        ok_btn = cand2\n                        break\n                except Exception:\n                    continue\n\n            # Fallback: any visible global OK button\n            if ok_btn is None:\n                try:\n                    global_ok = page.get_by_role("button", name="OK")\n                    if global_ok.count() > 0:\n                        for idx in range(global_ok.count() - 1, -1, -1):\n                            cand = global_ok.nth(idx)\n                            if cand.is_visible(timeout=800):\n                                ok_btn = cand\n                                break\n                except Exception:\n                    pass\n\n            if ok_btn is not None:\n                log.info("  ℹ️ Submit confirmation popup detected")\n                clicked_ok = False\n                try:\n                    ok_btn.click(timeout=2_000)\n                    clicked_ok = True\n                except Exception:\n                    pass\n\n                if not clicked_ok:\n                    try:\n                        js_mousedown(page, ok_btn, "Submit confirmation OK")\n                        clicked_ok = True\n                    except Exception:\n                        pass\n\n                if not clicked_ok:\n                    ok_btn.focus(timeout=1_000)\n                    page.keyboard.press("Enter")\n                    clicked_ok = True\n\n                if clicked_ok:\n                    wait_busy(page, timeout=15_000)\n                    time.sleep(0.6)\n                    log.info("  ✅ Confirmation OK clicked")\n                else:\n                    log.warning("  ⚠️ Confirmation popup found but OK click failed")\n            else:\n                log.info("  ℹ️ No post-submit confirmation popup detected")\n        except Exception as popup_exc:\n            log.warning(f"  ⚠️ Post-submit OK handling skipped: {popup_exc}")\n    except Exception as e:\n        log.warning(f"  ⚠️ Submit not completed automatically: {e}")\n\n\ndef fill_parameters_for_entry_and_submit(page: Page, entry: Dict[str, str], work_area: str):\n    """\n    Fill parameters using a single ESS Job entry (process_name + param_label + param_value).\n    Ensures multiple rows for the same process use their distinct parameter values.\n    """\n    process_name = entry.get("process_name", "").strip()\n    param_label = entry.get("param_label", "").strip()\n    param_value = entry.get("param_value", "").strip()\n    sanitized_param_value = sanitize_parameter_value(param_value)\n    if not (process_name and param_label):\n        raise RuntimeError("Entry missing process_name or param_label for parameter fill")\n\n    log.info(\n        "  🔗 Using entry param: %s > %s > %s",\n        process_name,\n        param_label,\n        sanitized_param_value,\n    )\n\n    result = run_parameter_handling(\n        page=page,\n        process_name=process_name,\n        job_request_id="AUTO",\n        work_area=work_area,\n        page_name="Scheduled Processes",\n        input_parameters=[{"label": param_label, "value": sanitized_param_value}],\n        strict_mode=False,\n        validation_mode="WARN_AND_HOLD",\n    )\n\n    submit_status = "FAILED"\n    submit_message = "Submit not attempted"\n\n    log.info("  📋 UI parameter labels discovered: %s", result.get("discovered_labels", []))\n    log.info("  🔗 Parameters mapped from entry: %s", result.get("parameters_to_fill", []))\n    extras = result.get("extra_input_labels", [])\n    if extras:\n        log.warning("  ⚠️ Entry labels not present on UI (ignored): %s", extras)\n    log.info("  🧪 Fill results: %s", result.get("fill_results", {}))\n    log.info("  ℹ️ Ignoring checkbox: \'Use Existing Index\'")\n\n    # Submit via the same logic as the primary flow\n    try:\n        log.info("  ➤ Clicking Process Details header Submit")\n        process_details = page.locator(\'[role="dialog"]:has-text("Process Details")\').first\n        submit_candidates = [\n            process_details.locator(\'xpath=.//button[normalize-space()="Submit"]\'),\n            process_details.locator(\'xpath=.//*[contains(normalize-space(),"Process Options")]/following::*[normalize-space()="Submit"][1]\'),\n            process_details.get_by_role("button", name="Submit"),\n            page.locator(\'xpath=//*[contains(normalize-space(),"Process Options")]/following::*[normalize-space()="Submit"][1]\'),\n            page.get_by_role("button", name="Submit"),\n        ]\n\n        submit_btn = None\n        for group in submit_candidates:\n            try:\n                if group.count() <= 0:\n                    continue\n                limit = min(group.count(), 5)\n                for idx in range(limit):\n                    cand = group.nth(idx)\n                    if cand.is_visible(timeout=1000):\n                        submit_btn = cand\n                        break\n                if submit_btn is not None:\n                    break\n            except Exception:\n                continue\n\n        if submit_btn is None:\n            raise RuntimeError("Header Submit button not found in Process Details")\n\n        clicked = False\n        try:\n            submit_btn.click(timeout=3_000)\n            clicked = True\n        except Exception:\n            pass\n\n        if not clicked:\n            try:\n                js_mousedown(page, submit_btn, "Process Details header Submit")\n                clicked = True\n            except Exception:\n                pass\n\n        if not clicked:\n            submit_btn.focus(timeout=1_000)\n            page.keyboard.press("Enter")\n            clicked = True\n\n        if not clicked:\n            raise RuntimeError("Unable to click Process Details header Submit")\n\n        wait_busy(page, timeout=30_000)\n        time.sleep(1.0)\n\n        submit_confirmed = False\n        success_checks = [\n            \'text="queued up for submission"\',\n            \'text="This process will be queued up for submission"\',\n            \'text="Request ID"\',\n            \'text="Confirmation"\',\n        ]\n        for sel in success_checks:\n            try:\n                page.locator(sel).first.wait_for(state="visible", timeout=4_000)\n                submit_confirmed = True\n                break\n            except Exception:\n                continue\n\n        if not submit_confirmed:\n            try:\n                process_details.wait_for(state="hidden", timeout=4_000)\n                submit_confirmed = True\n            except Exception:\n                pass\n\n        if submit_confirmed:\n            log.info("  ✅ Submit success confirmation detected")\n            submit_status = "SUCCESS"\n            submit_message = "Submit confirmed"\n        else:\n            log.info("  ✅ Submit clicked (no explicit confirmation selector detected)")\n            submit_status = "SUCCESS"\n            submit_message = "Submit clicked (no explicit confirmation selector)"\n\n        try:\n            log.info("  ➤ Checking for post-submit confirmation popup")\n            ok_btn = None\n            popup_candidates = [\n                page.locator(\'[role="dialog"]:has-text("Confirmation")\').first,\n                page.locator(\'[role="dialog"]:has-text("Process Details")\').first,\n                page.locator(\'[role="dialog"]\').last,\n            ]\n\n            for popup in popup_candidates:\n                try:\n                    popup.wait_for(state="visible", timeout=2_000)\n                    cand = popup.get_by_role("button", name="OK").first\n                    if cand.count() > 0 and cand.is_visible(timeout=1_000):\n                        ok_btn = cand\n                        break\n                    cand2 = popup.locator(\'xpath=.//*[normalize-space()="OK"]\').first\n                    if cand2.count() > 0 and cand2.is_visible(timeout=1_000):\n                        ok_btn = cand2\n                        break\n                except Exception:\n                    continue\n\n            if ok_btn is None:\n                try:\n                    global_ok = page.get_by_role("button", name="OK")\n                    if global_ok.count() > 0:\n                        for idx in range(global_ok.count() - 1, -1, -1):\n                            cand = global_ok.nth(idx)\n                            if cand.is_visible(timeout=800):\n                                ok_btn = cand\n                                break\n                except Exception:\n                    pass\n\n            if ok_btn is not None:\n                log.info("  ℹ️ Submit confirmation popup detected")\n                clicked_ok = False\n                try:\n                    ok_btn.click(timeout=2_000)\n                    clicked_ok = True\n                except Exception:\n                    pass\n\n                if not clicked_ok:\n                    try:\n                        js_mousedown(page, ok_btn, "Submit confirmation OK")\n                        clicked_ok = True\n                    except Exception:\n                        pass\n\n                if not clicked_ok:\n                    ok_btn.focus(timeout=1_000)\n                    page.keyboard.press("Enter")\n                    clicked_ok = True\n\n                if clicked_ok:\n                    wait_busy(page, timeout=15_000)\n                    time.sleep(0.6)\n                    log.info("  ✅ Confirmation OK clicked")\n                else:\n                    log.warning("  ⚠️ Confirmation popup found but OK click failed")\n            else:\n                log.info("  ℹ️ No post-submit confirmation popup detected")\n        except Exception as popup_exc:\n            log.warning(f"  ⚠️ Post-submit OK handling skipped: {popup_exc}")\n    except Exception as e:\n        log.warning(f"  ⚠️ Submit not completed automatically: {e}")\n        submit_status = "FAILED"\n        submit_message = str(e)\n\n    enriched_result = dict(result)\n    enriched_result.update(\n        {\n            "submit_status": submit_status,\n            "submit_message": submit_message,\n            "used_param_label": param_label,\n            "used_param_value_original": param_value,\n            "used_param_value_sanitized": sanitized_param_value,\n        }\n    )\n    return enriched_result\n\n# ─────────────────────────────────────────────────────────────────────────────\n# MAIN WORKFLOW\n# ─────────────────────────────────────────────────────────────────────────────\n\ndef main():\n    log.info("=" * 70)\n    log.info("Oracle Fusion – Navigate to Scheduled Processes")\n    log.info(f"  Base URL      : {BASE_URL}")\n    log.info(f"  Username      : {USERNAME}")\n    log.info(f"  Process Search: \'{PROCESS_SEARCH}\'" if PROCESS_SEARCH else "  Process Search: (none)")\n    log.info("=" * 70)\n\n    with sync_playwright() as pw:\n        browser = pw.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO)\n        context = browser.new_context(ignore_https_errors=True)\n        page = context.new_page()\n        page.set_default_timeout(DEFAULT_TIMEOUT)\n        \n        # Log console messages for ADF debugging\n        page.on("console", lambda msg: log.debug(f"Console [{msg.type}]: {msg.text}"))\n        page.on("pageerror", lambda err: log.error(f"Page error: {err}"))\n\n        try:\n            # ── Execute the workflow ────────────────────────────────────────\n            login(page)                              # Step 1\n            click_navigator(page)                    # Step 2\n            initial_work_area = capture_work_area_from_breadcrumb(page)\n            expand_tools_section(page)               # Step 3\n            click_scheduled_processes(page)          # Step 4\n            captured_work_area = capture_work_area_from_breadcrumb(page)\n            if captured_work_area.lower() == _normalize_work_area_text("Tools").lower() and initial_work_area:\n                # If Scheduled Processes page doesn\'t expose breadcrumb, reuse earlier captured setup context.\n                captured_work_area = initial_work_area\n\n            # Load all ESS Job entries once, then iterate over each.\n            additional_info_path = str((Path.cwd() / ADDITIONAL_INFO_EXCEL_FILE).resolve())\n            parsed_entries = load_process_params_from_additional_info(additional_info_path)\n            if not parsed_entries:\n                raise RuntimeError("No valid ESS Job entries found in Additional Info/ESS Job Value column")\n\n            job_log_lines: List[str] = []\n            job_log_lines.append("ESS JOB EXECUTION LOG")\n            job_log_lines.append("=" * 90)\n            job_log_lines.append(f"Generated On: {datetime.now().strftime(\'%Y-%m-%d %H:%M:%S\')}")\n            job_log_lines.append(f"Total Parsed ESS Entries: {len(parsed_entries)}")\n            job_log_lines.append("-" * 90)\n\n            for idx, entry in enumerate(parsed_entries, start=1):\n                log.info("=" * 70)\n                log.info("🚀 ESS Job %d/%d — %s", idx, len(parsed_entries), entry.get("process_name", ""))\n                process_name = str(entry.get("process_name", "")).strip()\n                param_label = str(entry.get("param_label", "")).strip()\n                raw_param_value = str(entry.get("param_value", "")).strip()\n                sanitized_param_value = sanitize_parameter_value(raw_param_value)\n\n                job_status = "FAILED"\n                submit_status = "FAILED"\n                fill_results = {}\n                failure_reason = ""\n                try:\n                    click_schedule_new_process(page)         # Step 5 for this entry\n\n                    selected_process_name = select_process_in_schedule_dialog_direct(\n                        page,\n                        process_name,\n                        direct=True,\n                    )  # Step 6/7\n                    if not selected_process_name:\n                        job_status = "SKIPPED"\n                        submit_status = "SKIPPED"\n                        failure_reason = "Process unavailable in selection dialog"\n                    else:\n                        open_process_details_from_schedule_dialog(page)                        # Step 8\n                        per_job_result = fill_parameters_for_entry_and_submit(\n                            page,\n                            entry,\n                            captured_work_area,\n                        )  # Step 9\n                        fill_results = per_job_result.get("fill_results", {})\n                        submit_status = str(per_job_result.get("submit_status", "FAILED"))\n                        job_status = "SUCCESS" if submit_status == "SUCCESS" else "FAILED"\n                        if submit_status != "SUCCESS":\n                            failure_reason = str(per_job_result.get("submit_message", "Submit failed"))\n                except Exception as job_exc:\n                    failure_reason = str(job_exc)\n                    log.error("❌ ESS Job %d failed: %s", idx, job_exc)\n\n                fill_status_text = "; ".join([f"{k}={v}" for k, v in fill_results.items()]) if fill_results else "NO_FILL_RESULT"\n                job_log_lines.append(\n                    f"JOB {idx}/{len(parsed_entries)} | process=\'{process_name}\' | param_label=\'{param_label}\' | "\n                    f"param_value_raw=\'{raw_param_value}\' | param_value_used=\'{sanitized_param_value}\' | "\n                    f"fill_status={fill_status_text} | submit_status={submit_status} | overall_status={job_status}"\n                )\n                if failure_reason:\n                    job_log_lines.append(f"  reason: {failure_reason}")\n\n            job_log_lines.append("-" * 90)\n            ESS_JOB_LOG_FILE.write_text("\\n".join(job_log_lines) + "\\n", encoding="utf-8")\n            log.info("🧾 Per-job execution log saved: %s", ESS_JOB_LOG_FILE)\n            \n            # ── Success ─────────────────────────────────────────────────────\n            log.info("")\n            log.info("=" * 70)\n            log.info("✅ DONE — Scheduled Processes screen ready")\n            log.info("   You can now:")\n            log.info("   • Verify the Request ID parameter was populated")\n            log.info("   • Confirm submission result/confirmation")\n            log.info("=" * 70)\n            \n\n        except Exception as exc:\n            log.error(f"❌ FAILED: {exc}", exc_info=True)\n            snap(page, "workflow_failed")\n            debug_dom_info(page, "Final state on failure")\n        finally:\n            log.info("🧪 Browser left open for manual testing.")\n            try:\n                if sys.stdin and sys.stdin.isatty():\n                    input("Press Enter to close the browser and end the script...")\n                else:\n                    log.info("Non-interactive terminal detected; skipping pause prompt.")\n            except (KeyboardInterrupt, EOFError):\n                log.info("Manual close interrupted; shutting down browser gracefully.")\n            browser.close()\n\n\nif __name__ == "__main__":\n    main()\n'


def _patch_profile_value_controls(profile_mod) -> None:
    dropdown_selectors = [
        'select[id*="soc2::content"]',
        'select[name*="soc2"]',
        'select.x2h',
        'select[aria-label*="Profile Value"]',
        'select[title*="Profile Value"]',
        'select',
    ]
    text_selectors = [
        'textarea[id*="soc2::content"]',
        'textarea[name*="soc2"]',
        'textarea[aria-label*="Profile Value"]',
        'input[id*="soc2::content"]:not([type="hidden"])',
        'input[name*="soc2"]:not([type="hidden"])',
        'input[aria-label*="Profile Value"]:not([type="hidden"])',
        'oj-input-text input:not([type="hidden"])',
        'oj-text-area textarea',
        'textarea',
        'input[type="text"]:not([type="hidden"])',
        'input:not([type]):not([type="hidden"])',
    ]

    def desired_yn(profile_value: str) -> str:
        return "Y" if str(profile_value or "").strip().lower() in {"yes", "y", "true", "1"} else "N"

    def first_visible_enabled(page, selectors):
        last_err = None
        for sel in selectors:
            try:
                locator = page.locator(sel)
                count = min(locator.count(), 20)
                for index in range(count):
                    candidate = locator.nth(index)
                    try:
                        if not candidate.is_visible(timeout=500):
                            continue
                        if not candidate.is_enabled(timeout=500):
                            continue
                        readonly = candidate.evaluate(
                            "el => !!(el.readOnly || el.disabled || el.getAttribute('aria-readonly') === 'true')"
                        )
                        if readonly:
                            continue
                        return candidate
                    except Exception as exc:
                        last_err = exc
                        continue
            except Exception as exc:
                last_err = exc
                continue
        if last_err:
            raise RuntimeError(str(last_err))
        return None

    def dropdown_options(locator):
        try:
            return locator.evaluate(
                """el => Array.from(el.options || []).map(o => ({
                    value: (o.value || '').trim().toLowerCase(),
                    label: (o.text || '').trim().toLowerCase(),
                }))"""
            ) or []
        except Exception:
            return []

    def dropdown_has_yes_no(locator) -> bool:
        options = dropdown_options(locator)
        values = {option.get("value", "") for option in options}
        labels = {option.get("label", "") for option in options}
        return ("y" in values or "yes" in labels) and ("n" in values or "no" in labels)

    def dropdown_selected_yn(locator) -> str:
        try:
            selected = locator.evaluate(
                """el => {
                    const idx = el.selectedIndex;
                    const opt = idx >= 0 ? el.options[idx] : null;
                    return {
                        value: ((el.value || opt?.value || '') + '').trim().toUpperCase(),
                        label: ((opt?.text || '') + '').trim().toUpperCase(),
                    };
                }"""
            ) or {}
        except Exception:
            return ""

        value = selected.get("value", "")
        label = selected.get("label", "")
        if value in {"Y", "YES"} or label == "YES":
            return "Y"
        if value in {"N", "NO"} or label == "NO":
            return "N"
        return ""

    def first_supported_dropdown(page):
        for sel in dropdown_selectors:
            try:
                locator = page.locator(sel)
                count = min(locator.count(), 20)
                for index in range(count):
                    candidate = locator.nth(index)
                    try:
                        if (
                            candidate.is_visible(timeout=500)
                            and candidate.is_enabled(timeout=500)
                            and dropdown_has_yes_no(candidate)
                        ):
                            return candidate
                    except Exception:
                        continue
            except Exception:
                continue
        return None

    def first_profile_value_dropdown(page, require_editable: bool = True):
        for sel in dropdown_selectors:
            try:
                locator = page.locator(sel)
                count = min(locator.count(), 40)
                for index in range(count):
                    candidate = locator.nth(index)
                    try:
                        if not candidate.is_visible(timeout=500):
                            continue
                        if require_editable and not candidate.is_enabled(timeout=500):
                            continue
                        if not dropdown_has_yes_no(candidate):
                            continue
                        return candidate
                    except Exception:
                        continue
            except Exception:
                continue
        return None

    def page_shows_profile_value(page, target: str) -> bool:
        """Return True when the current page visibly shows Profile Value already set."""
        dropdown = first_profile_value_dropdown(page, require_editable=False)
        if dropdown and dropdown_selected_yn(dropdown) == target:
            return True

        text_control = first_yn_text_control(page, allow_blank_profile_value=False)
        if text_control and get_text_value(text_control).strip().upper() == target:
            return True

        expected_label = "YES" if target == "Y" else "NO"
        try:
            return bool(
                page.evaluate(
                    """(expectedLabel) => {
                        const norm = s => (s || '').replace(/\s+/g, ' ').trim().toUpperCase();
                        const rows = Array.from(document.querySelectorAll('tr, [role="row"]'));
                        for (const row of rows) {
                            const text = norm(row.textContent);
                            if (text.includes('PROFILE VALUE') && text.includes(expectedLabel)) return true;
                            if (text.includes('SITE') && text.includes(expectedLabel)) return true;
                        }
                        const selected = Array.from(document.querySelectorAll('select option:checked'))
                            .some(o => norm(o.textContent) === expectedLabel || norm(o.value) === expectedLabel[0]);
                        return selected;
                    }""",
                    expected_label,
                )
            )
        except Exception:
            return False

    def get_text_value(locator) -> str:
        try:
            return (locator.input_value(timeout=1_500) or "").strip()
        except Exception:
            try:
                return (locator.evaluate("el => (el.value || el.textContent || '').trim()") or "").strip()
            except Exception:
                return ""

    def first_yn_text_control(page, allow_blank_profile_value: bool = False):
        for sel in text_selectors:
            try:
                locator = page.locator(sel)
                count = min(locator.count(), 40)
                for index in range(count):
                    candidate = locator.nth(index)
                    try:
                        if not candidate.is_visible(timeout=500):
                            continue
                        if not candidate.is_enabled(timeout=500):
                            continue
                        readonly = candidate.evaluate(
                            "el => !!(el.readOnly || el.disabled || el.getAttribute('aria-readonly') === 'true')"
                        )
                        if readonly:
                            continue

                        value = get_text_value(candidate).strip().upper()
                        if value in {"Y", "N"}:
                            return candidate
                        if allow_blank_profile_value and not value:
                            label = (
                                candidate.evaluate(
                                    "el => (el.getAttribute('aria-label') || el.getAttribute('title') || el.name || el.id || '').toLowerCase()"
                                )
                                or ""
                            )
                            if "profile" in label and "value" in label:
                                return candidate
                    except Exception:
                        continue
            except Exception:
                continue
        return None

    def set_text_value(locator, value: str) -> None:
        try:
            locator.scroll_into_view_if_needed(timeout=1_500)
        except Exception:
            pass

        filled = False
        try:
            locator.click(timeout=1_500)
            modifier = "Meta" if os.name == "posix" else "Control"
            locator.page.keyboard.press(f"{modifier}+A")
            locator.page.keyboard.press("Backspace")
            locator.fill(value, timeout=2_000)
            filled = True
        except Exception:
            filled = False

        if not filled:
            locator.evaluate(
                """(el, value) => {
                    el.focus && el.focus();
                    el.value = value;
                    el.textContent = value;
                    for (const name of ['input', 'change', 'blur']) {
                        el.dispatchEvent(new Event(name, { bubbles: true }));
                    }
                }""",
                value,
            )

        try:
            locator.evaluate(
                """el => {
                    for (const name of ['input', 'change', 'blur']) {
                        el.dispatchEvent(new Event(name, { bubbles: true }));
                    }
                }"""
            )
        except Exception:
            pass

    def profile_values_fetching(page) -> bool:
        try:
            return bool(
                page.evaluate(
                    """() => {
                        const text = (document.body?.innerText || '').replace(/\s+/g, ' ').toUpperCase();
                        if (text.includes('FETCHING DATA') || text.includes('FETCHING')) return true;
                        return !!document.querySelector(
                            '[role="progressbar"], .xbusy, [aria-busy="true"], .oj-progress, .AFBusyIndicator'
                        );
                    }"""
                )
            )
        except Exception:
            return False

    def profile_value_ready(page, target: str | None = None) -> bool:
        if first_supported_dropdown(page):
            return True
        if first_profile_value_dropdown(page, require_editable=False):
            return True
        if first_yn_text_control(page, allow_blank_profile_value=True):
            return True
        if target and page_shows_profile_value(page, target):
            return True
        return False

    def wait_for_profile_value_ready(page, target: str | None = None, timeout_ms: int = 25_000) -> None:
        profile_mod.log.info("  → Waiting for Profile Values section to finish fetching")
        deadline = profile_mod.time.time() + (timeout_ms / 1000)
        attempts = 0
        while profile_mod.time.time() < deadline:
            attempts += 1
            try:
                profile_mod.wait_busy(page, timeout=1_200)
            except Exception:
                pass
            if profile_value_ready(page, target) and not profile_values_fetching(page):
                profile_mod.log.info("  ✓ Profile Values section ready after %s check(s)", attempts)
                return
            profile_mod.time.sleep(0.35)

        if target and page_shows_profile_value(page, target):
            profile_mod.log.info("  ✓ Profile Value visible as %s after wait timeout", target)
            return
        raise RuntimeError("Profile Value control not ready after waiting for profile-values fetch")

    def ensure_profile_value_control_supported(page):
        wait_for_profile_value_ready(page)

        if first_supported_dropdown(page):
            return

        text_control = first_yn_text_control(page, allow_blank_profile_value=True)
        if text_control:
            current = get_text_value(text_control).strip().upper()
            if current and current not in {"Y", "N"}:
                raise RuntimeError(
                    f"Profile Value text control contains unsupported value '{current}'. Expected Y or N."
                )
            return

        if first_profile_value_dropdown(page, require_editable=False):
            return
        raise RuntimeError("Profile Value control not found or is hidden/disabled")

    def validate_dropdown_has_yes_no(page):
        wait_for_profile_value_ready(page)

        dropdown = first_supported_dropdown(page)
        if not dropdown:
            text_control = first_yn_text_control(page, allow_blank_profile_value=True)
            if text_control:
                current = get_text_value(text_control).strip().upper()
                if not current or current in {"Y", "N"}:
                    return
                raise RuntimeError(
                    f"Profile Value text control contains unsupported value '{current}'. Expected Y or N."
                )
            dropdown = first_profile_value_dropdown(page, require_editable=False)
            if dropdown:
                return
            raise RuntimeError("Profile Value control not found or is hidden/disabled")

        return

    def set_profile_value_with_verification(page, profile_value: str):
        target = desired_yn(profile_value)
        wait_for_profile_value_ready(page, target=target)

        dropdown = first_supported_dropdown(page)
        if dropdown:
            desired_is_yes = target == "Y"
            for _ in range(3):
                try:
                    dropdown.select_option(value=target)
                except Exception:
                    try:
                        dropdown.select_option(label="Yes" if desired_is_yes else "No")
                    except Exception:
                        pass
                profile_mod.time.sleep(0.4)
                selected_value = ""
                selected_label = ""
                try:
                    selected_value = (dropdown.input_value(timeout=1_500) or "").strip().upper()
                except Exception:
                    selected_value = ""
                try:
                    selected_label = (
                        dropdown.evaluate(
                            "el => el.selectedIndex >= 0 ? (el.options[el.selectedIndex]?.text || '').trim().toUpperCase() : ''"
                        )
                        or ""
                    )
                except Exception:
                    selected_label = ""
                if selected_value == target or selected_label in {target, "YES" if target == "Y" else "NO"}:
                    profile_mod.log.info("  ✓ Profile Value dropdown verified as %s", target)
                    return
            raise RuntimeError(f"Dropdown did not persist expected Profile Value '{target}'")

        readonly_dropdown = first_profile_value_dropdown(page, require_editable=False)
        if readonly_dropdown and dropdown_selected_yn(readonly_dropdown) == target:
            profile_mod.log.info("  ✓ Profile Value already verified as %s", target)
            return

        text_control = first_yn_text_control(page, allow_blank_profile_value=True)
        if not text_control:
            if page_shows_profile_value(page, target):
                profile_mod.log.info("  ✓ Profile Value already visible as %s", target)
                return
            raise RuntimeError("Profile Value control not found or is hidden/disabled")

        current = get_text_value(text_control).strip().upper()
        if current and current not in {"Y", "N"}:
            raise RuntimeError(
                f"Profile Value text control contains unsupported value '{current}'. Expected Y or N."
            )

        set_text_value(text_control, target)
        profile_mod.time.sleep(0.4)
        verified = get_text_value(text_control).strip().upper()
        if verified != target:
            set_text_value(text_control, target)
            profile_mod.time.sleep(0.4)
            verified = get_text_value(text_control).strip().upper()
        if verified != target:
            raise RuntimeError(
                f"Profile Value text control did not persist expected value '{target}'. Found '{verified}'."
            )
        profile_mod.log.info("  ✓ Profile Value text control verified as %s", target)

    profile_mod._ensure_profile_value_control_supported = ensure_profile_value_control_supported
    profile_mod._validate_dropdown_has_yes_no = validate_dropdown_has_yes_no
    profile_mod._set_profile_value_with_verification = set_profile_value_with_verification


def _load_inline_modules():
    optin_mod = types.ModuleType("optin_inline")
    profile_mod = types.ModuleType("profile_inline")
    ess_mod = types.ModuleType("ess_inline")

    sys.modules[optin_mod.__name__] = optin_mod
    sys.modules[profile_mod.__name__] = profile_mod
    sys.modules[ess_mod.__name__] = ess_mod

    exec(OPTIN_SOURCE, optin_mod.__dict__)
    exec(PROFILE_SOURCE, profile_mod.__dict__)
    exec(ESS_SOURCE, ess_mod.__dict__)
    optin_mod.login = lambda page: login(
        page,
        base_url=getattr(optin_mod, "BASE_URL", BASE_URL),
        username=getattr(optin_mod, "USERNAME", USERNAME),
        password=getattr(optin_mod, "PASSWORD", PASSWORD),
        timeout=getattr(optin_mod, "TIMEOUT", 45_000) * 2,
        wait_busy_func=getattr(optin_mod, "wait_busy", wait_for_oracle_busy_clear),
        logger=getattr(optin_mod, "log", log),
    )
    profile_mod.login = optin_mod.login
    ess_mod.login = optin_mod.login
    _patch_profile_value_controls(profile_mod)
    return optin_mod, profile_mod, ess_mod


def _header_map(ws):
    headers = [str(c.value or "").strip() for c in ws[1]]
    return {" ".join(h.lower().split()): i for i, h in enumerate(headers) if h}


def _norm(s: str) -> str:
    return " ".join((s or "").split()).strip()


def _extract_action_flags(action_text: str) -> set[str]:
    """
    Parse Action Required and return which of the supported actions are present.

    Supports any single action, pair, or all-three combinations, in any order,
    with separators like +, comma, slash, ampersand, and mixed casing/spaces.
    """
    norm = _norm(str(action_text or "")).casefold()
    if not norm:
        return set()

    # Select All means run all available categories.
    if "select all" in norm:
        return {"opt_in", "ess_jobs", "profile_options", "auto"}

    flags = set()
    for token, flag in SUPPORTED_ACTION_TOKENS.items():
        if token in norm:
            flags.add(flag)
    return flags


def _parse_profile_entries(cell_text: str):
    out = []
    if not cell_text:
        return out
    pattern = re.compile(
        r"^(?P<task>[^>]+?)\s*>\s*(?P<code>_|[A-Z][A-Z0-9_]*[A-Z0-9])\s*=\s*(?P<value>[^,\n\r]+)$",
        re.I,
    )
    for entry in split_existing_pattern_entries(str(cell_text).replace("\r", "\n")):
        normalized_entry = entry.strip().strip('"').strip("'")
        m = pattern.match(normalized_entry)
        if not m:
            if normalized_entry:
                out.append(
                    {
                        "source_value": normalized_entry,
                        "skip_reason": "Profile Options Value does not match required pattern: Task > PROFILE_CODE = Value",
                    }
                )
            continue
        task = _norm(m.group("task")) or "Manage Administrator Profile Values"
        code = normalize_profile_option_code(m.group("code"))
        value = _norm(m.group("value")) or "Yes"
        if not is_valid_profile_option_code(code):
            out.append(
                {
                    "task_name": task,
                    "code": code,
                    "value": value,
                    "source_value": normalized_entry,
                    "skip_reason": "Profile option code does not match required code pattern",
                }
            )
            continue
        out.append(
            {
                "task_name": task,
                "code": code,
                "value": value,
                "source_value": f"{task} > {code} = {value}",
            }
        )
    return out


def _parse_ess_entries(cell_text: str):
    out = []
    if not cell_text:
        return out
    for entry in split_existing_pattern_entries(str(cell_text).replace("\r", "\n")):
        pieces = [_norm(piece) for piece in entry.split(">")]
        pieces = [piece for piece in pieces if piece]
        if len(pieces) < 3:
            continue
        if not is_valid_ess_entry(pieces[0], pieces[1], pieces[2]):
            continue
        param_label = normalize_ess_parameter_label_for_process(pieces[0], pieces[1])
        out.append(
            {
                "process_name": pieces[0],
                "param_label": param_label,
                "param_value": pieces[2],
                "source_value": f"{pieces[0]} > {param_label} > {pieces[2]}",
            }
        )
    return out


def _execution_key(*parts: object) -> tuple[str, ...]:
    """Case/space-insensitive key used to remember completed Excel actions."""
    return tuple(_norm(str(part or "")).casefold() for part in parts)


def _dedupe_profile_entries(entries: list[dict]) -> list[dict]:
    """Keep only executable, unique profile option entries from uploaded Excel."""
    unique_entries: list[dict] = []
    seen = set()
    for entry in entries:
        if _norm(entry.get("skip_reason", "")):
            continue

        task = entry.get("task_name") or "Manage Administrator Profile Values"
        code = entry.get("code")
        value = entry.get("value") or "Yes"
        if not code:
            continue

        dedupe_key = (*_execution_key(task, code), _profile_value_key(value))
        if dedupe_key in seen:
            continue
        seen.add(dedupe_key)
        unique_entries.append(entry)
    return unique_entries


def _dedupe_ess_entries(entries: list[dict]) -> list[dict]:
    """Keep only executable, unique ESS job entries from uploaded Excel."""
    unique_entries: list[dict] = []
    seen = set()
    for entry in entries:
        process_name = entry.get("process_name", "")
        param_label = entry.get("param_label", "")
        param_value = entry.get("param_value", "")
        if not (process_name and param_label and param_value):
            continue

        job_key = _execution_key(process_name, param_label, param_value)
        if job_key in seen:
            continue
        seen.add(job_key)
        unique_entries.append(entry)
    return unique_entries


def _profile_value_key(value: object) -> str:
    text = _norm(str(value or "")).casefold()
    if text in {"y", "yes", "true", "1"}:
        return "y"
    if text in {"n", "no", "false", "0"}:
        return "n"
    return text


def _detect_and_dismiss_warning_popup(page) -> str:
    """
    Detect *blocking* warning/error popup dialogs only (not page-level banners/text)
    and dismiss them best-effort.
    Returns popup text when a blocking dialog is found, else empty string.
    """
    dialog_locators = [
        page.locator('[role="dialog"]:has-text("Warning")').first,
        page.locator('[role="alertdialog"]:has-text("Warning")').first,
        page.locator('[role="dialog"]:has-text("Error")').first,
        page.locator('[role="alertdialog"]:has-text("Error")').first,
    ]

    popup = None
    popup_text = ""
    for dlg in dialog_locators:
        try:
            if dlg.is_visible(timeout=700):
                popup = dlg
                try:
                    popup_text = (dlg.inner_text(timeout=700) or "").strip()
                except Exception:
                    popup_text = "Blocking warning popup detected"
                break
        except Exception:
            continue

    if popup is None:
        return ""

    dismiss_candidates = [
        popup.get_by_role("button", name="OK").first,
        popup.get_by_role("button", name="Close").first,
        popup.locator('button:has-text("OK")').first,
        popup.locator('button:has-text("Close")').first,
        popup.locator('a:has-text("OK")').first,
        popup.locator('a:has-text("Close")').first,
        popup.locator('[aria-label="Close"]').first,
    ]
    for btn in dismiss_candidates:
        try:
            if btn.is_visible(timeout=500):
                btn.click(timeout=1_000)
                break
        except Exception:
            continue

    return popup_text or "Blocking warning popup detected"


def _is_new_features_search_ready(page) -> bool:
    selectors = [
        'input[aria-label="Feature"]',
        'input[name*="qbeFeature"]',
        'input[id*="qbeFeature"]',
    ]
    for sel in selectors:
        try:
            if page.locator(sel).first.is_visible(timeout=1200):
                return True
        except Exception:
            continue
    return False


def _normalized_excel_output_path(input_path: Path) -> Path:
    configured_path = os.getenv("NORMALIZED_EXCEL_PATH", "").strip()
    if configured_path:
        return Path(configured_path)
    return input_path.with_name(f"{input_path.stem}_normalized.xlsx")


def _resolve_required_reference_template_path() -> Path:
    """Resolve and validate the reference workbook path used for normalization."""
    configured = os.getenv("NORMALIZATION_TEMPLATE", "").strip()
    template_path = Path(configured) if configured else CANONICAL_TEMPLATE_PATH

    if not template_path.exists():
        source = "NORMALIZATION_TEMPLATE" if configured else "default CANONICAL_TEMPLATE_PATH"
        raise FileNotFoundError(
            "Reference workbook is required before normalization but was not found: "
            f"{template_path} (source={source}). "
            "Place the reference workbook in project root as 'SCM_REDWOOD_FEATURES.xlsx' "
            "or set NORMALIZATION_TEMPLATE to an existing file path."
        )

    return template_path


def _build_normalized_feature_frame(excel_path: str) -> tuple[pd.DataFrame, str]:
    """
    Convert customer workbook layouts into the canonical Filtered Features shape.

    The output keeps combine.py's existing Profile Options and ESS Jobs cell
    patterns unchanged:
    - Profile Options Value: Task name > PROFILE_CODE = value
    - ESS Jobs Value: Process name > Parameter label > Parameter value
    """
    input_path = Path(excel_path)
    if not input_path.exists():
        raise FileNotFoundError(f"Excel file does not exist: {input_path}")

    requested_sheet = os.getenv("EXCEL_SHEET", "").strip() or None
    template_path = _resolve_required_reference_template_path()
    log.info("Normalization reference workbook resolved: '%s'", template_path)

    sheets = pd.read_excel(input_path, sheet_name=None, keep_default_na=False)
    for sheet_name, frame in sheets.items():
        log.info(
            "Excel loaded into memory: file='%s' | sheet='%s' | rows=%d",
            input_path,
            sheet_name,
            len(frame),
        )
    sheets = apply_excel_visible_rows(input_path, sheets)
    source_sheets = select_source_sheets(sheets, requested_sheet)
    for sheet_name, frame in source_sheets:
        log.info(
            "Excel source sheet selected: sheet='%s' | rows=%d",
            sheet_name,
            len(frame),
        )
    reference_patterns = load_reference_patterns(template_path)
    log.info(
        "Reference patterns loaded: ess_processes=%d | ess_parameters=%d | action_values=%d",
        len(reference_patterns.ess_processes),
        len(reference_patterns.ess_parameters),
        len(reference_patterns.action_values),
    )
    normalized_frames = []
    for sheet_name, source_frame in source_sheets:
        before_count = len(source_frame)
        normalized_frame = _normalize_frame_for_combine(source_frame.fillna(""), reference_patterns)
        log.info(
            "Excel normalized sheet: sheet='%s' | before=%d | after=%d",
            sheet_name,
            before_count,
            len(normalized_frame),
        )
        normalized_frames.append(normalized_frame)
    normalized = pd.concat(normalized_frames, ignore_index=True)
    log.info("Excel normalized concat: source_sheets=%d | rows=%d", len(normalized_frames), len(normalized))
    source_sheet_names = ", ".join(sheet_name for sheet_name, _ in source_sheets)

    log.info(
        "Excel normalized in memory: file='%s' | source_sheets='%s' | rows=%d",
        input_path,
        source_sheet_names,
        len(normalized),
    )
    return normalized[TARGET_COLUMNS], source_sheet_names


def normalize_excel_file(excel_path: str) -> Path:
    """Write a normalized Excel workbook and return the path used by the script."""
    input_path = Path(excel_path)
    output_path = _normalized_excel_output_path(input_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    normalized, source_sheet = _build_normalized_feature_frame(excel_path)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        normalized.to_excel(writer, sheet_name="Filtered Features", index=False)

    log.info(
        "Normalized Excel written: source_sheets='%s' | output='%s' | rows=%d",
        source_sheet,
        output_path,
        len(normalized),
    )
    return output_path


def _read_normalized_feature_frame(normalized_excel_path: Path) -> pd.DataFrame:
    normalized = pd.read_excel(
        normalized_excel_path,
        sheet_name="Filtered Features",
        keep_default_na=False,
    )
    log.info(
        "Normalized Excel loaded into memory: file='%s' | sheet='Filtered Features' | rows=%d",
        normalized_excel_path,
        len(normalized),
    )
    missing_columns = [column for column in TARGET_COLUMNS if column not in normalized.columns]
    if missing_columns:
        raise RuntimeError(
            f"Normalized Excel file is missing required columns: {missing_columns}"
        )

    log.info("Script input Excel: '%s'", normalized_excel_path)
    return normalized[TARGET_COLUMNS].fillna("")


def _source_value(row, column: str | None) -> str:
    if not column:
        return ""
    return _norm(str(row.get(column, "") or ""))


def _source_entry_value(row, column: str | None) -> str:
    """Return a cell value with multiple entries normalized to comma separators."""
    if not column:
        return ""
    value = row.get(column, "")
    if pd.isna(value):
        return ""

    text = str(value).replace("\r\n", "\n").replace("\r", "\n").strip()
    lines = [_norm(line) for line in text.split("\n")]
    return ", ".join(line for line in lines if line)


def _source_raw_text(row, column: str | None) -> str:
    if not column:
        return ""
    value = row.get(column, "")
    if pd.isna(value):
        return ""
    return str(value).replace("\r\n", "\n").replace("\r", "\n").strip()


def _best_column(frame: pd.DataFrame, target: str, hints=None, threshold: float = 0.72) -> str | None:
    exact = find_exact_output_column(frame.columns, target)
    if exact:
        return exact
    return find_column(frame.columns, target, hints, threshold=threshold)


def _split_comma_entries(value: str) -> list[str]:
    return [_norm(part) for part in re.split(r"\s*(?:,|\n|\r)+\s*", str(value or "")) if _norm(part)]


def _build_profile_value_for_combine(row, task_col: str | None, option_col: str | None, value_col: str | None) -> str:
    option = _source_value(row, option_col)
    if not option:
        return ""

    task = _source_value(row, task_col) or "Manage Administrator Profile Values"
    value = _source_value(row, value_col) or "Yes"

    entries = []
    for option_entry in _split_comma_entries(option):
        entry_value = value
        if "=" in option_entry:
            code, existing_value = option_entry.split("=", 1)
            option_entry = _norm(code)
            entry_value = _norm(existing_value) or entry_value

        if option_entry and is_valid_profile_option_code(option_entry):
            entries.append(format_profile_option_entry(task, option_entry, entry_value))

    return ", ".join(entries)


def _build_ess_value_for_combine(row, process_col: str | None, parameter_col: str | None, value_col: str | None) -> str:
    process = _source_value(row, process_col)
    if not process:
        return ""

    parameter = _source_value(row, parameter_col)
    value = _source_value(row, value_col)
    if parameter and value:
        return format_ess_entry(process, parameter, value)
    return preserve_ess_acronyms(process)


def _extract_labeled_piece(text: str, labels: list[str]) -> str:
    label_expr = "|".join(re.escape(label) for label in sorted(labels, key=len, reverse=True))
    stop_expr = (
        r"ess job name|ess job|scheduled process name|scheduled process|process name|"
        r"parameter label|parameter name|parameter value|paramter label|paramter name|"
        r"paramter value|parametr label|parametr name|parametr value"
    )
    match = re.search(
        rf"\b(?:{label_expr})\b\s*(?:is|as|=|:|-)?\s*"
        rf"(.+?)(?=\s*(?:,|;|\||\n|\band\b)?\s*\b(?:{stop_expr})\b\s*(?:is|as|=|:|-)?|$)",
        text,
        flags=re.I,
    )
    return _norm(match.group(1).strip(" .;,")) if match else ""


def _format_ess_entries_for_combine(value: str) -> str:
    entries = []
    for entry in split_existing_pattern_entries(value):
        pieces = [clean_pattern_piece(piece) for piece in entry.split(">")]
        pieces = [piece for piece in pieces if piece]
        if len(pieces) >= 3 and is_valid_ess_entry(pieces[0], pieces[1], pieces[2]):
            entries.append(format_ess_entry(pieces[0], pieces[1], pieces[2]))
    return preserve_ess_acronyms(", ".join(entries)) if entries else ""


def _extract_ess_from_additional_info_for_combine(text: str, reference_patterns=None) -> str:
    raw_text = str(text or "").replace("\r", "\n").strip()
    if not raw_text:
        return ""
    if ">" in raw_text and _parse_ess_entries(raw_text):
        return _format_ess_entries_for_combine(raw_text)

    test_py_value = extract_ess_job_from_text(raw_text, reference_patterns)
    if test_py_value:
        return _format_ess_entries_for_combine(test_py_value)

    process = _extract_labeled_piece(
        raw_text,
        ["ESS Job Name", "ESS Job", "Scheduled Process Name", "Scheduled Process", "Process Name"],
    )
    parameter = _extract_labeled_piece(
        raw_text,
        ["Parameter Label", "Parameter Name", "Paramter Label", "Paramter Name", "Parametr Label", "Parametr Name"],
    )
    value = _extract_labeled_piece(
        raw_text,
        ["Parameter Value", "Paramter Value", "Parametr Value"],
    )
    if process and parameter and value:
        return format_ess_entry(process, parameter, value)
    return ""


def _extract_profile_from_any_text_for_combine(text: str) -> str:
    raw_text = str(text or "").replace("\r", "\n").strip()
    if not raw_text:
        return ""

    # Already formatted profile pattern
    if ">" in raw_text and "=" in raw_text:
        normalized = normalize_profile_options_cell(raw_text, "")
        if normalized:
            return normalized

    # Fallback: ORA_* code discovery anywhere in text
    return _format_direct_profile_value_for_combine(raw_text)


def _row_text_for_global_fallback(row, excluded_columns: set[str]) -> str:
    parts = []
    for column_name, value in row.items():
        col = str(column_name)
        if col in excluded_columns:
            continue
        if pd.isna(value):
            continue
        text = str(value).replace("\r\n", "\n").replace("\r", "\n").strip()
        if text:
            parts.append(text)
    return "\n".join(parts)


def _is_auto_action(action: str) -> bool:
    return is_auto_action_label(action)


def _categorize_action_for_combine(
    action: str,
    feature_name: str,
    profile_value: str,
    ess_value: str,
) -> str:
    return action_required_from_normalized_values(action, profile_value, ess_value, "")


def _normalize_frame_for_combine(source_frame: pd.DataFrame, reference_patterns=None) -> pd.DataFrame:
    product_col = _best_column(source_frame, "Product", HEADER_HINTS.get("Product"))
    module_col = _best_column(source_frame, "Module", HEADER_HINTS.get("Module"))
    feature_col = _best_column(source_frame, "Feature Name", HEADER_HINTS.get("Feature Name"))
    action_col = _best_column(source_frame, "Action Required", HEADER_HINTS.get("Action Required"))

    direct_profile_col = find_exact_output_column(source_frame.columns, "Profile Options Value")
    task_col = _best_column(source_frame, "Task Name", PROFILE_TASK_HINTS, threshold=0.70)
    option_col = _best_column(source_frame, "Profile Option", PROFILE_OPTION_HINTS, threshold=0.70)
    profile_value_col = _best_column(source_frame, "Profile Option Value", PROFILE_VALUE_HINTS, threshold=0.80)
    if direct_profile_col:
        if option_col == direct_profile_col:
            option_col = None
        if profile_value_col == direct_profile_col:
            profile_value_col = None

    direct_ess_col = find_exact_output_column(source_frame.columns, "ESS Jobs Value")
    process_col = _best_column(source_frame, "Process Name", ESS_PROCESS_HINTS, threshold=0.70)
    parameter_col = _best_column(source_frame, "Parameter Label", ESS_PARAMETER_HINTS, threshold=0.76)
    ess_value_col = _best_column(source_frame, "Parameter Value", ESS_VALUE_HINTS, threshold=0.80)
    additional_info_col = _best_column(source_frame, "Additional Info", ADDITIONAL_INFO_HINTS, threshold=0.70)

    excluded_for_global_scan = {
        column
        for column in [
            product_col,
            module_col,
            feature_col,
            action_col,
            direct_profile_col,
            task_col,
            option_col,
            profile_value_col,
            direct_ess_col,
            process_col,
            parameter_col,
            ess_value_col,
            additional_info_col,
        ]
        if column
    }

    rows = []
    for _, row in source_frame.iterrows():
        source_action = _source_value(row, action_col)
        product = _source_value(row, product_col)
        module = _source_value(row, module_col)
        feature_name = _source_value(row, feature_col)
        reference_row = find_reference_feature_row(
            reference_patterns,
            product,
            module,
            feature_name,
        )
        if reference_row:
            rows.append(
                {
                    "Product": clean_cell(reference_row.get("Product", product), ""),
                    "Module": clean_cell(reference_row.get("Module", module), ""),
                    "Feature Name": clean_cell(reference_row.get("Feature Name", feature_name), ""),
                    "Profile Options Value": _format_direct_profile_value_for_combine(
                        clean_cell(reference_row.get("Profile Options Value", ""), "")
                    ),
                    "ESS Jobs Value": normalize_ess_cell(
                        clean_cell(reference_row.get("ESS Jobs Value", ""), ""),
                        "",
                    ),
                    "Action Required": canonicalize_action_required(
                        reference_row.get("Action Required", ""),
                        "",
                    ),
                }
            )
            continue

        profile_value = _source_entry_value(row, direct_profile_col)
        if profile_value:
            profile_value = _format_direct_profile_value_for_combine(profile_value)
        if not profile_value:
            profile_value = _build_profile_value_for_combine(row, task_col, option_col, profile_value_col)

        ess_value = _source_entry_value(row, direct_ess_col)
        if ess_value:
            ess_value = normalize_ess_cell(ess_value, "")
        if not ess_value:
            ess_value = _build_ess_value_for_combine(row, process_col, parameter_col, ess_value_col)
        if not ess_value and additional_info_col:
            ess_value = _extract_ess_from_additional_info_for_combine(
                _source_raw_text(row, additional_info_col),
                reference_patterns,
            )

        # Global fallback: parse free text from any other column in the row.
        if not profile_value or not ess_value:
            global_row_text = _row_text_for_global_fallback(row, excluded_for_global_scan)
            if global_row_text:
                if not profile_value:
                    profile_value = _extract_profile_from_any_text_for_combine(global_row_text)
                if not ess_value:
                    ess_value = _extract_ess_from_additional_info_for_combine(
                        global_row_text,
                        reference_patterns,
                    )

        action = _categorize_action_for_combine(
            source_action,
            feature_name,
            profile_value,
            ess_value,
        )

        rows.append(
            {
                "Product": product,
                "Module": module,
                "Feature Name": feature_name,
                "Profile Options Value": profile_value,
                "ESS Jobs Value": ess_value,
                "Action Required": action,
            }
        )

    return pd.DataFrame(rows, columns=TARGET_COLUMNS)


def _cell_value(row, column: str) -> str:
    return _norm(str(row.get(column, "") or ""))


def _entry_cell_value(row, column: str) -> str:
    value = row.get(column, "")
    if pd.isna(value):
        return ""

    text = str(value).replace("\r\n", "\n").replace("\r", "\n").strip()
    lines = [_norm(line) for line in text.split("\n")]
    return ", ".join(line for line in lines if line)


def _drop_truly_blank_normalized_rows(normalized: pd.DataFrame) -> pd.DataFrame:
    blank_mask = normalized.apply(
        lambda row: all(_cell_value(row, column) == "" for column in TARGET_COLUMNS),
        axis=1,
    )
    blank_indices = list(normalized.index[blank_mask])
    if blank_indices:
        excel_rows = [index + 2 for index in blank_indices]
        log.info(
            "Normalized blank-row filter: before=%d | excluded=%d | after=%d | reason='all target columns blank' | excel_rows=%s",
            len(normalized),
            len(blank_indices),
            len(normalized) - len(blank_indices),
            excel_rows,
        )
        return normalized.loc[~blank_mask].reset_index(drop=True)

    log.info(
        "Normalized blank-row filter: before=%d | excluded=0 | after=%d",
        len(normalized),
        len(normalized),
    )
    return normalized


def load_filtered_data(excel_path: str):
    normalized_excel_path = normalize_excel_file(excel_path)
    normalized = _read_normalized_feature_frame(normalized_excel_path)
    log.info("Pipeline trace: normalized rows before invalid-row filtering=%d", len(normalized))
    normalized = _drop_truly_blank_normalized_rows(normalized)

    optin_features, profile_rows, ess_rows, auto_rows, feature_rows = [], [], [], [], []
    seen_optin = set()
    raw_optin_feature_count = 0
    duplicate_optin_feature_count = 0
    raw_profile_entry_count = 0
    invalid_profile_entry_count = 0
    raw_ess_entry_count = 0
    rows_without_action_flags = 0

    for row_index, row in normalized.iterrows():
        action = _cell_value(row, "Action Required")
        action_norm = action.casefold()
        product = _cell_value(row, "Product")
        module = _cell_value(row, "Module")
        feature = _cell_value(row, "Feature Name")
        profile_value = _entry_cell_value(row, "Profile Options Value")
        ess_value = _entry_cell_value(row, "ESS Jobs Value")

        feature_rows.append(
            {
                "source_row": row_index + 2,
                "feature_name": feature,
                "opt_in_output": action,
                "profile_option_output": profile_value,
                "ess_jobs_output": ess_value,
            }
        )

        # Auto-direct rows are already enabled by default.
        # Do not route them through any execution flow; log only via run_auto().
        if action_norm == "auto":
            auto_rows.append(
                {
                    "product": product,
                    "module": module,
                    "feature": feature,
                    "action_required": action,
                    "profile_options_value": profile_value,
                    "ess_jobs_value": ess_value,
                }
            )
            continue

        action_flags = _extract_action_flags(action)
        if not action_flags:
            rows_without_action_flags += 1
            continue

        if "opt_in" in action_flags and product and feature:
            raw_optin_feature_count += 1
            key = (product.casefold(), module.casefold(), feature.casefold()) if module else None
            if key is None or key not in seen_optin:
                if key is not None:
                    seen_optin.add(key)
                optin_features.append(
                    {
                        "product": product,
                        "module": module,
                        "name": feature,
                        "source_row": row_index + 2,
                    }
                )
            else:
                duplicate_optin_feature_count += 1
        elif "opt_in" in action_flags:
            raw_optin_feature_count += 1

        if "profile_options" in action_flags:
            parsed_profile_entries = _parse_profile_entries(profile_value)
            raw_profile_entry_count += len(parsed_profile_entries)
            invalid_profile_entry_count += sum(1 for entry in parsed_profile_entries if _norm(entry.get("skip_reason", "")))
            profile_rows.extend(parsed_profile_entries)

        if "ess_jobs" in action_flags:
            parsed_ess_entries = _parse_ess_entries(ess_value)
            raw_ess_entry_count += len(parsed_ess_entries)
            ess_rows.extend(parsed_ess_entries)

        if "auto" in action_flags:
            auto_rows.append(
                {
                    "product": product,
                    "module": module,
                    "feature": feature,
                    "action_required": action,
                    "profile_options_value": profile_value,
                    "ess_jobs_value": ess_value,
                }
            )

    log.info(
        "Pipeline trace: feature_rows=%d | rows_without_action_flags=%d",
        len(feature_rows),
        rows_without_action_flags,
    )
    log.info(
        "Pipeline trace: opt-in features raw=%d | duplicates_skipped_for_execution=%d | execution_count=%d",
        raw_optin_feature_count,
        duplicate_optin_feature_count,
        len(optin_features),
    )
    log.info(
        "Pipeline trace: profile entries raw=%d | invalid_skipped=%d",
        raw_profile_entry_count,
        invalid_profile_entry_count,
    )
    profile_rows = _dedupe_profile_entries(profile_rows)
    log.info("Pipeline trace: profile entries after drop_duplicates=%d", len(profile_rows))
    log.info("Pipeline trace: ESS entries raw=%d", raw_ess_entry_count)
    ess_rows = _dedupe_ess_entries(ess_rows)
    log.info("Pipeline trace: ESS entries after drop_duplicates=%d", len(ess_rows))
    log.info("Pipeline trace: auto rows=%d", len(auto_rows))

    return optin_features, profile_rows, ess_rows, auto_rows, feature_rows


def _append_feature_log(
    log_path: Path,
    feature_name: str,
    opt_in_output: str,
    profile_option_output: str,
    ess_jobs_output: str,
    stage: str,
    source_row: object = "",
):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(
            f"[{ts}] Stage={stage} | SourceRow='{source_row}' | Feature='{feature_name}' | "
            f"OptIn='{opt_in_output}' | ProfileOption='{profile_option_output}' | ESSJobs='{ess_jobs_output}'\n"
        )


def run_auto(auto_rows, auto_log_path: Path):
    if not auto_rows:
        log.info("Auto: no filtered rows to process")
        _append_realtime(auto_log_path, "No Auto entries to process")
        return {"processed": 0}

    for idx, row in enumerate(auto_rows, start=1):
        _append_realtime(
            auto_log_path,
            (
                f"Row={idx} | Product='{row.get('product', '')}' | Module='{row.get('module', '')}' "
                f"| Feature='{row.get('feature', '')}' | ActionRequired='{row.get('action_required', '')}' "
                f"| Status=AUTO ENABLED | ProfileOptionsValue='{row.get('profile_options_value', '')}' "
                f"| ESSJobsValue='{row.get('ess_jobs_value', '')}'"
            ),
        )
    _append_realtime(auto_log_path, f"Auto rows logged as enabled={len(auto_rows)}")
    return {"processed": len(auto_rows), "enabled": len(auto_rows)}


def run_optin(page, optin_mod, features, optin_log_path: Path, screenshots_dir: Path):
    results = []
    if not features:
        log.info("Opt-in: no filtered rows to process")
        _append_realtime(optin_log_path, "No opt-in features to process")
        return results

    # Preserve uploaded Excel order. All opt-in rows are processed from the
    # Procurement offering's global Features Overview instead of navigating to
    # each row's Product tile.
    scoped_features = []
    seen_excel_scope = set()
    for f in features:
        product = _norm(f.get("product", ""))
        module = _norm(f.get("module", ""))
        feature_name = _norm(f.get("name", ""))
        source_row = f.get("source_row", "")
        if not product or not feature_name:
            _append_realtime(
                optin_log_path,
                (
                    "Status=SKIPPED | Reason='Missing Product/Feature in uploaded Excel' "
                    f"| SourceRow='{source_row}' | Product='{product}' | Module='{module}' | Feature='{feature_name}'"
                ),
            )
            continue

        # Customer uploads may contain only Product + Feature Names. When Module is
        # absent, keep each row so the input count mirrors the normalized workbook;
        # the New Features page supplies the functional area after search.
        if module:
            scope_key = (product.casefold(), module.casefold(), feature_name.casefold())
            if scope_key in seen_excel_scope:
                continue
            seen_excel_scope.add(scope_key)
        else:
            scope_key = (product.casefold(), feature_name.casefold(), str(source_row))
            seen_excel_scope.add(scope_key)

        scoped_features.append(
            {
                "product": product,
                "module": module,
                "name": feature_name,
                "source_row": source_row,
            }
        )

    if not scoped_features:
        log.info("Opt-in: no valid feature rows to process")
        _append_realtime(optin_log_path, "No valid feature rows for opt-in")
        return results

    _append_realtime(
        optin_log_path,
        (
            "Opt-in scope loaded from uploaded Excel only | "
            "Offering='Procurement' | FeaturesOverview='All' | "
            f"InputRows={len(scoped_features)}"
        ),
    )

    original_read_matching_feature_rows = optin_mod.read_matching_feature_rows

    def _module_scoped_matching_rows(pg, feature_name: str):
        rows = original_read_matching_feature_rows(pg, feature_name)
        expected_module = _norm(getattr(optin_mod, "_EXPECTED_EXCEL_MODULE", ""))
        if not expected_module:
            return rows

        expected_key = expected_module.casefold()
        scoped_rows = [
            row for row in rows
            if _norm(row.get("functional_area", "")).casefold() == expected_key
        ]
        if rows and not scoped_rows:
            log.info(
                "Opt-in row skipped by Excel module scope: feature='%s' expected_module='%s' matched_modules=%s",
                feature_name,
                expected_module,
                sorted({_norm(row.get("functional_area", "")) for row in rows if row.get("functional_area")}),
            )
        return scoped_rows

    optin_mod.read_matching_feature_rows = _module_scoped_matching_rows

    # Enforce strict feature search field targeting (no global Search fallback).
    def _strict_search_feature(pg, feature_name: str):
        search_box = None
        for sel in [
            'input[aria-label="Feature"]',
            'input[name*="qbeFeature"]',
            'input[id*="qbeFeature"]',
        ]:
            try:
                pg.wait_for_selector(sel, state="visible", timeout=4_000)
                search_box = pg.locator(sel).first
                search_box.click(timeout=1_500)
                break
            except Exception:
                continue

        if not search_box:
            raise RuntimeError("New Features search field not visible")

        search_box.fill("")
        search_box.type(feature_name, delay=30)
        if (search_box.input_value() or "").strip() != feature_name.strip():
            search_box.click()
            search_box.press("ControlOrMeta+A")
            search_box.press("Backspace")
            search_box.fill(feature_name)
        pg.keyboard.press("Enter")
        time.sleep(1.2)
        try:
            optin_mod.wait_busy(pg)
        except Exception:
            pass

    optin_mod.search_feature = _strict_search_feature

    def _safe_locator_click(pg, locator, label: str, timeout_ms: int = 3_000) -> bool:
        try:
            locator.click(timeout=timeout_ms, no_wait_after=True)
            return True
        except TypeError:
            try:
                locator.click(timeout=timeout_ms)
                return True
            except Exception:
                pass
        except Exception:
            pass

        try:
            locator.evaluate("""el => {
                el.scrollIntoView({block: 'center', inline: 'center'});
                ['mousedown', 'mouseup', 'click'].forEach(name =>
                    el.dispatchEvent(new MouseEvent(name, {
                        bubbles: true, cancelable: true, view: window, button: 0, buttons: 1
                    }))
                );
            }""")
            log.info("%s clicked via DOM event fallback", label)
            return True
        except Exception:
            return False

    def _select_features_overview_all(pg) -> None:
        """Select Features Overview = All after Available Features is opened."""
        dropdown_selectors = [
            'a[id$="soc1::drop"][aria-haspopup="true"]',
            'a[id*="AP1:soc1::drop"][aria-haspopup="true"]',
            'table[id$="AP1:soc1"] a[aria-haspopup="true"]',
        ]
        option_selectors = [
            'li[role="option"][_adfiv="-1"]',
            'ul[id$="soc1::pop"] li[role="option"]:text-is("All")',
            'li[role="option"]:text-is("All")',
        ]

        dropdown, dropdown_selector = _visible_first(pg, dropdown_selectors, timeout_ms=2_000)
        if dropdown is None:
            raise RuntimeError("Features Overview dropdown not visible after opening Available Features")

        if not _safe_locator_click(pg, dropdown, "Features Overview dropdown", timeout_ms=4_000):
            raise RuntimeError(f"Could not open Features Overview dropdown via '{dropdown_selector}'")

        option = None
        option_selector = ""
        for _ in range(3):
            option, option_selector = _visible_first(pg, option_selectors, timeout_ms=1_500)
            if option is not None:
                break
            # ADF can render the list before making it visible. Re-open the
            # dropdown and let its scrollable popup settle before retrying.
            _safe_locator_click(pg, dropdown, "Features Overview dropdown", timeout_ms=2_000)
            time.sleep(0.4)

        if option is None:
            raise RuntimeError("Features Overview 'All' option not visible")

        try:
            option.scroll_into_view_if_needed(timeout=2_000)
        except Exception:
            try:
                option.evaluate("el => el.scrollIntoView({block: 'center'})")
            except Exception:
                pass

        if not _safe_locator_click(pg, option, "Features Overview: All", timeout_ms=4_000):
            raise RuntimeError(f"Could not select Features Overview 'All' via '{option_selector}'")

        try:
            optin_mod.wait_busy(pg, timeout=30_000)
        except Exception:
            pass
        time.sleep(UI_STABILIZE_SEC)

        # Do not continue to feature searches until the closed dropdown visibly
        # shows "All". ADF can update its internal value before repainting the
        # displayed selection, especially on slower instances.
        try:
            pg.wait_for_function(
                """() => {
                    const visible = (el) => {
                        if (!el) return false;
                        const style = window.getComputedStyle(el);
                        const rect = el.getBoundingClientRect();
                        return style.visibility !== 'hidden' &&
                            style.display !== 'none' && rect.width > 0 && rect.height > 0;
                    };
                    const norm = (value) => (value || '').replace(/\\s+/g, ' ').trim().toLowerCase();
                    const controls = [
                        document.querySelector('input[id$="soc1::content"]'),
                        document.querySelector('table[id$="AP1:soc1"]'),
                        document.querySelector('[title="Features Overview : All"]')
                    ];
                    return controls.some(el => visible(el) && (
                        norm(el.value) === 'all' ||
                        norm(el.textContent) === 'all' ||
                        norm(el.getAttribute('title')) === 'features overview : all'
                    ));
                }""",
                timeout=max(NAV_TIMEOUT_MS, 30_000),
            )
        except Exception as exc:
            raise RuntimeError(
                "Features Overview dropdown did not visibly show 'All' before timeout"
            ) from exc

        log.info("Features Overview dropdown visibly shows All")

    original_go_to_new_features = optin_mod.go_to_new_features

    def _go_to_new_features_with_all(pg):
        original_go_to_new_features(pg)
        _select_features_overview_all(pg)

    # Every initial navigation and recovery path must reselect the global view.
    optin_mod.go_to_new_features = _go_to_new_features_with_all

    def _install_robust_edit_features_optin_patch():
        original_process_feature = optin_mod.process_feature

        def _mark_edit_checkbox_target(pg, feature_name: str, functional_area: str) -> dict:
            return pg.evaluate(
                """(payload) => {
                    const norm = (s) => (s || '')
                        .toLowerCase()
                        .replace(/[^a-z0-9\\s]/g, ' ')
                        .replace(/\\s+/g, ' ')
                        .trim();
                    const tokenSet = (s) => new Set(norm(s).split(' ').filter(Boolean));
                    const overlapScore = (a, b) => {
                        const A = tokenSet(a), B = tokenSet(b);
                        if (!A.size || !B.size) return 0;
                        let hit = 0;
                        for (const t of A) if (B.has(t)) hit++;
                        return hit / Math.max(A.size, B.size);
                    };
                    const compact = (s) => (s || '').replace(/\\s+/g, ' ').trim();
                    const readText = (el) => {
                        if (!el) return '';
                        const parts = [el.innerText, el.textContent, el.getAttribute?.('title'), el.getAttribute?.('aria-label')];
                        el.querySelectorAll?.('[title], [aria-label]')?.forEach(child => {
                            parts.push(child.getAttribute('title'));
                            parts.push(child.getAttribute('aria-label'));
                        });
                        const seen = new Set();
                        return parts.map(compact).filter(Boolean).filter(part => {
                            const key = part.toLowerCase();
                            if (seen.has(key)) return false;
                            seen.add(key);
                            return true;
                        }).join(' ');
                    };
                    const checkboxState = (label) => {
                        const inputId = label.getAttribute('for');
                        const input = inputId ? document.getElementById(inputId) : null;
                        const aria = label.getAttribute('aria-checked') || input?.getAttribute?.('aria-checked') || '';
                        const checkedAttr = input?.getAttribute?.('checked');
                        const value = (input?.value || '').toLowerCase();
                        const selectedClass = /selected|checked|active/i.test(`${label.className || ''} ${input?.className || ''}`);
                        const checkedIcon = !!label.querySelector?.(
                            'img[src*="check"], img[src*="accept"], img[alt*="Checked"], img[title*="Checked"], img[alt*="Selected"], img[title*="Selected"]'
                        );
                        return {
                            checked: !!(input?.checked || checkedAttr === 'checked' || aria === 'true' || value === 'true' || value === 'y' || selectedClass || checkedIcon),
                            input_id: inputId || '',
                            input_type: input?.type || '',
                            input_checked: !!input?.checked,
                            aria_checked: aria,
                            value,
                            label_class: `${label.className || ''}`,
                        };
                    };
                    const collectCandidates = (targetText) => {
                        const targetNorm = norm(targetText);
                        if (!targetNorm) return [];
                        const candidates = [];
                        for (const row of Array.from(document.querySelectorAll('tr'))) {
                            const label = row.querySelector('label.x17j');
                            if (!label) continue;
                            const textNodes = Array.from(row.querySelectorAll(
                                'span.x2ey, span[style*="white-space:normal"], td, th, a[title], span[title]'
                            ));
                            const rowText = compact((textNodes.length ? textNodes.map(readText).join(' ') : readText(row)));
                            const rowNorm = norm(rowText);
                            if (!rowNorm) continue;
                            let score = 0;
                            if (rowNorm === targetNorm) score = 1.0;
                            else if (rowNorm.includes(targetNorm) || targetNorm.includes(rowNorm)) score = 0.9;
                            else score = overlapScore(rowNorm, targetNorm);
                            if (score >= 0.45) candidates.push({ row, label, rowText, score });
                        }
                        candidates.sort((a, b) => b.score - a.score || b.rowText.length - a.rowText.length);
                        return candidates;
                    };

                    document.querySelectorAll('label[data-codex-optin-target]').forEach(label => {
                        label.removeAttribute('data-codex-optin-target');
                    });

                    let matchedBy = 'feature';
                    let candidates = collectCandidates(payload.feature);
                    if (!candidates.length) {
                        matchedBy = 'functional_area';
                        candidates = collectCandidates(payload.functional_area);
                    }
                    if (!candidates.length) {
                        return {found: false, enabled: false, status: 'not-found', matched_by: null};
                    }

                    const best = candidates[0];
                    best.label.setAttribute('data-codex-optin-target', '1');
                    try { best.label.scrollIntoView({block: 'center', inline: 'center'}); } catch (e) {}
                    const state = checkboxState(best.label);
                    return {
                        found: true,
                        enabled: state.checked,
                        status: state.checked ? 'already-checked' : 'needs-click',
                        matched_by: matchedBy,
                        matched_text: best.rowText,
                        score: best.score,
                        candidate_count: candidates.length,
                        ...state,
                    };
                }""",
                {"feature": feature_name.strip(), "functional_area": (functional_area or "").strip()},
            ) or {"found": False, "enabled": False, "status": "not-found", "matched_by": None}

        def _target_checkbox_state(pg) -> dict:
            return pg.evaluate(
                """() => {
                    const label = document.querySelector('label[data-codex-optin-target="1"]');
                    if (!label) return {found: false, checked: false};
                    const inputId = label.getAttribute('for');
                    const input = inputId ? document.getElementById(inputId) : null;
                    const aria = label.getAttribute('aria-checked') || input?.getAttribute?.('aria-checked') || '';
                    const checkedAttr = input?.getAttribute?.('checked');
                    const value = (input?.value || '').toLowerCase();
                    const selectedClass = /selected|checked|active/i.test(`${label.className || ''} ${input?.className || ''}`);
                    const checkedIcon = !!label.querySelector?.(
                        'img[src*="check"], img[src*="accept"], img[alt*="Checked"], img[title*="Checked"], img[alt*="Selected"], img[title*="Selected"]'
                    );
                    return {
                        found: true,
                        checked: !!(input?.checked || checkedAttr === 'checked' || aria === 'true' || value === 'true' || value === 'y' || selectedClass || checkedIcon),
                        input_id: inputId || '',
                        input_checked: !!input?.checked,
                        aria_checked: aria,
                        value,
                    };
                }"""
            ) or {"found": False, "checked": False}

        def _click_target_checkbox_human(pg) -> bool:
            try:
                box = pg.locator('label[data-codex-optin-target="1"]').first.bounding_box(timeout=2_000)
            except Exception:
                box = None
            if not box:
                return False

            try:
                cx = box["x"] + (box["width"] / 2)
                cy = box["y"] + (box["height"] / 2)
                pg.mouse.move(cx, cy)
                time.sleep(0.15)
                pg.mouse.down()
                time.sleep(0.12)
                pg.mouse.up()
                return True
            except Exception:
                return False

        def _set_target_checkbox_checked_via_dom(pg) -> bool:
            try:
                return bool(
                    pg.evaluate(
                        """() => {
                            const label = document.querySelector('label[data-codex-optin-target="1"]');
                            if (!label) return false;
                            const inputId = label.getAttribute('for');
                            const input = inputId ? document.getElementById(inputId) : null;
                            if (!input) return false;

                            input.focus?.();
                            if ('checked' in input) input.checked = true;
                            input.setAttribute('checked', 'checked');
                            input.setAttribute('aria-checked', 'true');
                            label.setAttribute('aria-checked', 'true');

                            for (const name of ['mousedown', 'mouseup', 'click', 'input', 'change', 'blur']) {
                                input.dispatchEvent(new Event(name, {bubbles: true, cancelable: true}));
                            }
                            for (const name of ['mousedown', 'mouseup', 'click']) {
                                label.dispatchEvent(new MouseEvent(name, {
                                    bubbles: true, cancelable: true, view: window, button: 0, buttons: 1
                                }));
                            }
                            return true;
                        }"""
                    )
                )
            except Exception:
                return False

        def _ensure_target_checkbox_checked(pg, reason: str, timeout_ms: int = 8_000) -> bool:
            state = _target_checkbox_state(pg)
            if not state.get("found"):
                log.warning("  Edit Features checkbox target missing before %s", reason)
                return False
            if state.get("checked"):
                log.info("  Edit Features checkbox already checked before %s", reason)
                return True

            log.info("  Edit Features checkbox is unchecked before %s; enabling it now", reason)
            label_locator = pg.locator('label[data-codex-optin-target="1"]').first
            input_id = str(state.get("input_id") or "")
            input_locator = pg.locator(f'[id="{input_id}"]').first if input_id else None

            attempts = [
                ("label", lambda: label_locator.click(timeout=3_000, no_wait_after=True)),
                ("human-center", lambda: _click_target_checkbox_human(pg)),
                ("input", lambda: input_locator.click(timeout=3_000, no_wait_after=True) if input_locator else False),
                ("label-force", lambda: label_locator.click(timeout=3_000, force=True, no_wait_after=True)),
                ("dom-set", lambda: _set_target_checkbox_checked_via_dom(pg)),
            ]

            for attempt_name, action in attempts:
                try:
                    result = action()
                    clicked = True if result is None else bool(result)
                except TypeError:
                    try:
                        if attempt_name == "label":
                            label_locator.click(timeout=3_000)
                            clicked = True
                        elif attempt_name == "input" and input_locator:
                            input_locator.click(timeout=3_000)
                            clicked = True
                        elif attempt_name == "label-force":
                            label_locator.click(timeout=3_000, force=True)
                            clicked = True
                        else:
                            clicked = False
                    except Exception:
                        clicked = False
                except Exception:
                    clicked = False

                if not clicked:
                    continue

                try:
                    pg.keyboard.press("Tab")
                except Exception:
                    pass
                if _wait_for_target_checkbox_checked(pg, timeout_ms=timeout_ms):
                    log.info("  Edit Features checkbox checked successfully via %s", attempt_name)
                    return True

            log.warning("  Edit Features checkbox is still unchecked before %s; Done will not be clicked", reason)
            return False

        def _wait_for_target_checkbox_checked(pg, timeout_ms: int = 6_000) -> bool:
            deadline = time.time() + (timeout_ms / 1000)
            stable_checks = 0
            while time.time() < deadline:
                try:
                    optin_mod.wait_busy(pg, timeout=1_000)
                except Exception:
                    pass
                state = _target_checkbox_state(pg)
                if state.get("checked"):
                    stable_checks += 1
                    if stable_checks >= 2:
                        return True
                else:
                    stable_checks = 0
                time.sleep(0.35)
            return False

        def _patched_enable_on_edit_features_page(pg, feature_name: str, functional_area: str = "") -> dict:
            log.info("  -> Edit Features page: enabling '%s'", feature_name[:55])
            if functional_area:
                log.info("  -> Functional Area fallback: '%s'", functional_area[:55])

            try:
                pg.wait_for_selector('label.x17j', state="attached", timeout=getattr(optin_mod, "TIMEOUT", 45_000))
            except Exception:
                log.warning("  Edit Features page did not load (label.x17j not found)")
                return {"enabled": False, "status": "edit-page-timeout", "matched_by": None}

            result = {"status": "not-found", "matched_by": None, "enabled": False}
            for scroll_attempt in range(18):
                result = _mark_edit_checkbox_target(pg, feature_name, functional_area)
                if result.get("found"):
                    if scroll_attempt:
                        log.info("  Edit Features match found after auto-scroll attempt %d", scroll_attempt + 1)
                    break

                moved = pg.evaluate(
                    """() => {
                        const isScrollable = (el) => {
                            if (!el) return false;
                            const style = window.getComputedStyle(el);
                            return /(auto|scroll)/.test(style.overflowY || '') && el.scrollHeight > el.clientHeight + 20;
                        };
                        const containers = Array.from(document.querySelectorAll(
                            'div, table, tbody, [role="grid"], [role="treegrid"], [role="rowgroup"]'
                        )).filter(isScrollable);
                        containers.sort((a, b) => (b.scrollHeight - b.clientHeight) - (a.scrollHeight - a.clientHeight));
                        const scroller = containers[0] || document.scrollingElement || document.documentElement;
                        const beforeTop = scroller.scrollTop;
                        const beforeWin = window.scrollY;
                        const step = Math.max(260, Math.floor((scroller.clientHeight || window.innerHeight || 600) * 0.75));
                        scroller.scrollTop = Math.min(scroller.scrollTop + step, scroller.scrollHeight);
                        window.scrollBy(0, step);
                        return scroller.scrollTop !== beforeTop || window.scrollY !== beforeWin;
                    }"""
                )
                if not moved:
                    try:
                        pg.mouse.wheel(0, 700)
                    except Exception:
                        pass
                time.sleep(0.45)

            log.info(
                "  Tick target: %s | candidates=%s | matched_by=%s | score=%s | match='%s'",
                result.get("status"),
                result.get("candidate_count"),
                result.get("matched_by"),
                result.get("score"),
                str(result.get("matched_text", ""))[:80],
            )

            if not result.get("found"):
                log.warning("  Feature and Functional Area not found on Edit Features page")
                return {"enabled": False, "status": "not-found", "matched_by": None}

            if result.get("enabled"):
                log.info("  Already enabled on Edit Features page (%s)", result.get("matched_by"))
                return {**result, "enabled": True, "status": "already-checked"}

            if _ensure_target_checkbox_checked(pg, "Done"):
                return {**result, "enabled": True, "status": "ticked"}

            log.warning("  Edit Features checkbox did not stay checked before Done")
            return {**result, "enabled": False, "status": "click-no-stable-state"}

        def _patched_click_done(pg):
            target_state = _target_checkbox_state(pg)
            if target_state.get("found") and not _ensure_target_checkbox_checked(pg, "Done"):
                log.warning("  Skipping Done because Edit Features checkbox is not checked")
                return

            try:
                pg.keyboard.press("Tab")
            except Exception:
                pass
            try:
                optin_mod.wait_busy(pg, timeout=10_000)
            except Exception:
                pass
            time.sleep(0.8)

            selectors = [
                'a[accesskey="o"][role="button"]',
                'a.xrg[role="button"]',
                'a[role="button"]:has-text("Done")',
                'button:has-text("Done")',
                'a.xrg',
            ]
            for sel in selectors:
                try:
                    pg.wait_for_selector(sel, state="visible", timeout=5_000)
                    btn = pg.locator(sel).first
                    if _safe_locator_click(pg, btn, "Done", timeout_ms=5_000):
                        try:
                            optin_mod.wait_busy(pg, timeout=90_000)
                        except Exception:
                            pass
                        time.sleep(6.0)
                        log.info("  Done clicked; waited for slow page transition")
                        return
                except Exception:
                    continue
            log.warning("  Done button not found")

        def _verified_process_feature(pg, feat):
            result = original_process_feature(pg, feat)
            if getattr(result, "status", "") != "Enabled Now":
                return result

            try:
                optin_mod.search_feature(pg, feat.name)
                rows = optin_mod.read_matching_feature_rows(pg, feat.name)
                if rows and any(row.get("is_enabled") for row in rows):
                    log.info("  Verified enabled state on New Features page after Done")
                    return result

                error = "Checkbox was selected on Edit Features page, but New Features still shows disabled after Done"
                log.warning("  %s", error)
                return optin_mod.FeatureResult(
                    name=result.name,
                    functional_area=result.functional_area,
                    status="Error",
                    error=error,
                )
            except Exception as exc:
                error = f"Could not verify enabled state after Done: {exc}"
                log.warning("  %s", error)
                return optin_mod.FeatureResult(
                    name=result.name,
                    functional_area=result.functional_area,
                    status="Error",
                    error=error,
                )

        optin_mod.enable_on_edit_features_page = _patched_enable_on_edit_features_page
        optin_mod.click_done = _patched_click_done
        optin_mod.process_feature = _verified_process_feature

    _install_robust_edit_features_optin_patch()

    def _visible_first(pg, selectors, timeout_ms: int = 1_500):
        for sel in selectors:
            try:
                locator = pg.locator(sel)
                count = min(locator.count(), 8)
                for idx in range(count):
                    candidate = locator.nth(idx)
                    if candidate.is_visible(timeout=timeout_ms):
                        return candidate, sel
            except Exception:
                continue
        return None, ""

    def _setup_and_maintenance_ready(pg, timeout_ms: int = 2_500) -> bool:
        ready_selectors = [
            'td.xo2:has-text("Actions")',
            'button:has-text("Actions")',
            'a:has-text("Actions")',
            ':text-is("Go to Offerings")',
        ]
        deadline = time.time() + (timeout_ms / 1000)
        while time.time() < deadline:
            try:
                optin_mod.wait_busy(pg)
            except Exception:
                pass
            locator, _ = _visible_first(pg, ready_selectors, timeout_ms=500)
            if locator is not None:
                return True
            time.sleep(0.25)
        return False

    def _safe_go_to_setup_and_maintenance(pg):
        settings_selectors = [
            '#pt1\\:_UIScmil2u',
            'img[title="Settings and Actions"]',
            'img[alt="Settings and Actions"]',
            '[aria-label="Settings and Actions"]',
            '[title="Settings and Actions"]',
        ]
        setup_selectors = [
            'a:has-text("Setup and Maintenance")',
            '[role="menuitem"]:has-text("Setup and Maintenance")',
            'td:has-text("Setup and Maintenance")',
            'text=Setup and Maintenance',
        ]

        last_err = None
        for attempt in range(1, NAV_RETRIES + 1):
            try:
                if _setup_and_maintenance_ready(pg, timeout_ms=1_000):
                    log.info("Setup and Maintenance already ready (attempt %d)", attempt)
                    return

                settings_sel = ""
                settings, settings_sel = _visible_first(pg, settings_selectors, timeout_ms=1_000)
                if settings is None:
                    clicked_settings = False
                    for sel in settings_selectors:
                        try:
                            pg.click(sel, timeout=SHORT_WAIT_MS, no_wait_after=True)
                            settings_sel = sel
                            clicked_settings = True
                            break
                        except TypeError:
                            pg.click(sel, timeout=SHORT_WAIT_MS)
                            settings_sel = sel
                            clicked_settings = True
                            break
                        except Exception as exc:
                            last_err = exc
                            continue
                    if not clicked_settings:
                        raise RuntimeError("Settings and Actions control not visible")
                else:
                    if not _safe_locator_click(pg, settings, "Settings and Actions"):
                        raise RuntimeError("Settings and Actions click failed")

                time.sleep(0.8)
                setup_item, setup_sel = _visible_first(pg, setup_selectors, timeout_ms=1_500)
                if setup_item is None:
                    raise RuntimeError("Setup and Maintenance menu item not visible")

                if not _safe_locator_click(pg, setup_item, "Setup and Maintenance"):
                    raise RuntimeError("Setup and Maintenance click failed")
                time.sleep(UI_STABILIZE_SEC)
                try:
                    optin_mod.wait_busy(pg)
                except Exception:
                    pass

                if _setup_and_maintenance_ready(pg, timeout_ms=NAV_TIMEOUT_MS):
                    log.info(
                        "Setup and Maintenance opened (attempt %d) via settings='%s', item='%s'",
                        attempt,
                        settings_sel,
                        setup_sel,
                    )
                    return
                raise RuntimeError("Setup and Maintenance page did not expose expected controls")
            except Exception as exc:
                last_err = exc
                log.warning("Setup and Maintenance navigation failed (attempt %d/%d): %s", attempt, NAV_RETRIES, exc)
                time.sleep(UI_STABILIZE_SEC)

        raise RuntimeError(f"Setup and Maintenance navigation failed after retries: {last_err}")

    optin_mod.go_to_setup_and_maintenance = _safe_go_to_setup_and_maintenance

    original_go_to_offerings = optin_mod.go_to_offerings

    def _safe_go_to_offerings(pg, reason: str = ""):
        last_err = None
        for attempt in range(1, NAV_RETRIES + 1):
            try:
                original_go_to_offerings(pg)
                try:
                    optin_mod.wait_busy(pg)
                except Exception:
                    pass
                time.sleep(UI_STABILIZE_SEC)
                log.info("Go to Offerings succeeded (attempt %d)%s", attempt, f" | {reason}" if reason else "")
                return
            except Exception as exc:
                last_err = exc
                log.warning(
                    "Go to Offerings failed (attempt %d/%d)%s: %s",
                    attempt,
                    NAV_RETRIES,
                    f" | {reason}" if reason else "",
                    exc,
                )
                try:
                    optin_mod.go_to_setup_and_maintenance(pg)
                    time.sleep(UI_STABILIZE_SEC)
                except Exception:
                    pass
        raise RuntimeError(f"Go to Offerings failed after retries: {last_err}")

    optin_mod.go_to_offerings = _safe_go_to_offerings

    optin_mod.go_to_setup_and_maintenance(page)
    optin_mod.go_to_offerings(page, reason="initial")

    offering = "Procurement"
    optin_mod.OFFERING = offering
    try:
        # go_to_new_features opens Procurement -> New Features -> Available
        # Features; the installed wrapper then selects Features Overview = All.
        optin_mod.go_to_new_features(page)
    except Exception as nav_ex:
        error = f"Procurement New Features / Features Overview All not accessible: {nav_ex}"
        log.error("Opt-in navigation failed: %s", error)
        _append_realtime(optin_log_path, f"Offering='{offering}' | Status=FAIL | Error='{error}'")
        return results

    processed_count = 0
    total_count = len(scoped_features)
    _append_realtime(
        optin_log_path,
        f"Offering='{offering}' | FeaturesOverview='All' | Status=START | Features={total_count}",
    )

    for f in scoped_features:
        processed_count += 1
        product_log_prefix = f"SourceProduct='{f['product']}' | Offering='{offering}'"
        feat = optin_mod.Feature(name=f["name"], module=f["module"], row=processed_count + 1)
        try:
            # Stay in Procurement's global Available Features context. Any
            # recovery path reopens the page and reselects Overview = All.
            if not _is_new_features_search_ready(page):
                try:
                    optin_mod.go_to_offerings(page, reason="recovery_procurement_all")
                    optin_mod.OFFERING = offering
                    optin_mod.go_to_new_features(page)
                except Exception as recovery_ex:
                    raise RuntimeError(
                        f"Could not recover Procurement Features Overview All: {recovery_ex}"
                    ) from recovery_ex

            _detect_and_dismiss_warning_popup(page)

            # In the All view, use the Excel Module as Functional Area scope so
            # identical feature names do not select a different result row.
            optin_mod._EXPECTED_EXCEL_MODULE = f["module"]
            r = optin_mod.process_feature(page, feat)

            post_warn = _detect_and_dismiss_warning_popup(page)
            if post_warn:
                err_result = optin_mod.FeatureResult(
                    name=f["name"],
                    functional_area=f["module"],
                    status="Error",
                    error=f"Warning popup after processing: {post_warn[:180]}",
                )
                results.append(err_result)
                _append_realtime(
                    optin_log_path,
                    f"{product_log_prefix} | Feature='{err_result.name}' | FunctionalArea='{err_result.functional_area}' | Status=FAIL | Error='{err_result.error}'",
                )
                _capture_workflow_screenshot(
                    page, screenshots_dir, "optin", "after", processed_count,
                    err_result.name, "error",
                )
                continue

            results.append(r)
            _append_realtime(
                optin_log_path,
                f"{product_log_prefix} | Feature='{r.name}' | FunctionalArea='{r.functional_area}' | Status={r.status} | Error='{r.error or ''}'",
            )
            _capture_workflow_screenshot(
                page, screenshots_dir, "optin", "after", processed_count,
                r.name, r.status,
            )
        except Exception as exc:
            warn_text = _detect_and_dismiss_warning_popup(page)
            final_err = str(exc)
            if warn_text:
                final_err = f"{final_err} | Warning popup: {warn_text[:180]}"
            err_result = optin_mod.FeatureResult(
                name=f["name"], functional_area=f["module"], status="Error", error=final_err
            )
            results.append(err_result)
            _append_realtime(
                optin_log_path,
                f"{product_log_prefix} | Feature='{err_result.name}' | FunctionalArea='{err_result.functional_area}' | Status=FAIL | Error='{err_result.error}'",
            )
            _capture_workflow_screenshot(
                page, screenshots_dir, "optin", "after", processed_count,
                err_result.name, "error",
            )

    _append_realtime(
        optin_log_path,
        f"Offering='{offering}' | FeaturesOverview='All' | Status=COMPLETED | ProcessedFeatures={processed_count}/{total_count}",
    )
    return results


def run_profile(page, profile_mod, profile_entries, profile_log_path: Path, screenshots_dir: Path):
    if not profile_entries:
        log.info("Profile Options: no filtered rows to process")
        _append_realtime(profile_log_path, "No profile option entries to process")
        return 0, 0

    total_pass, total_fail = 0, 0

    # Remember every Profile Options Value entry already seen in the Excel input,
    # then batch unique entries by task/value so each setup task opens once per batch.
    seen = set()
    grouped_entries: dict[tuple[str, str], dict[str, object]] = {}
    group_order: list[tuple[str, str]] = []
    skipped_duplicates = 0
    skipped_invalid = 0
    for e in profile_entries:
        skip_reason = _norm(e.get("skip_reason", ""))
        source_value = _norm(e.get("source_value", ""))
        if skip_reason:
            skipped_invalid += 1
            continue
        task = e.get("task_name") or "Manage Administrator Profile Values"
        value = e.get("value") or "Yes"
        code = e.get("code")
        if not code:
            skipped_invalid += 1
            continue
        source_value = source_value or f"{task} > {code} = {value}"
        dedupe_key = (*_execution_key(task, code), _profile_value_key(value))
        if dedupe_key in seen:
            skipped_duplicates += 1
            continue
        seen.add(dedupe_key)

        group_key = (_norm(task).casefold(), _profile_value_key(value))
        if group_key not in grouped_entries:
            grouped_entries[group_key] = {
                "task": _norm(task),
                "value": _norm(value) or "Yes",
                "codes": [],
            }
            group_order.append(group_key)
        grouped_entries[group_key]["codes"].append(_norm(code))

    if skipped_duplicates:
        log.info("Profile Options: skipped duplicate Excel entries: %d", skipped_duplicates)
    if skipped_invalid:
        log.info("Profile Options: skipped invalid Excel entries: %d", skipped_invalid)

    processed_index = 0
    for group_key in group_order:
        group = grouped_entries.get(group_key, {})
        task = str(group.get("task") or "Manage Administrator Profile Values")
        value = str(group.get("value") or "Yes")
        codes = list(group.get("codes") or [])
        if not codes:
            continue

        try:
            log.info(
                "Profile Options task batch: Task='%s' | Value='%s' | Codes=%d",
                task,
                value,
                len(codes),
            )
            _append_realtime(
                profile_log_path,
                f"Task='{task}' | Value='{value}' | Status=START_BATCH | Codes={len(codes)}",
            )
            p, f = profile_mod.set_profile_options(
                page,
                codes,
                profile_value=value,
                task_name=task,
                log_path=profile_log_path,
            )
            total_pass += p
            total_fail += f
            status = "PASS" if p == len(codes) and f == 0 else "PARTIAL_FAIL" if p else "FAIL"
            _append_realtime(
                profile_log_path,
                f"Task='{task}' | Value='{value}' | Status={status} | Passed={p} | Failed={f} | Codes='{', '.join(codes)}'",
            )
            for code in codes:
                processed_index += 1
                _capture_workflow_screenshot(
                    page,
                    screenshots_dir,
                    "profile",
                    "after",
                    processed_index,
                    f"{code}_{value}",
                    "set" if status == "PASS" else "error",
                )
        except Exception as ex:
            total_fail += len(codes)
            _append_realtime(
                profile_log_path,
                f"Task='{task}' | Value='{value}' | Status=FAIL | Error='{str(ex)}' | Codes='{', '.join(codes)}'",
            )
            for code in codes:
                processed_index += 1
                _capture_workflow_screenshot(
                    page,
                    screenshots_dir,
                    "profile",
                    "after",
                    processed_index,
                    f"{code}_{value}",
                    "error",
                )
    return total_pass, total_fail


def _click_ess_locator(page, locator, description: str, js_mousedown_func=None) -> bool:
    """Click an ADF control using the same cascade as the ESS module."""
    try:
        locator.click(timeout=2_000)
        return True
    except Exception:
        pass

    try:
        if callable(js_mousedown_func):
            js_mousedown_func(page, locator, description)
            return True
    except Exception:
        pass

    try:
        locator.focus(timeout=1_000)
        page.keyboard.press("Enter")
        return True
    except Exception:
        return False


def _find_ess_submit_button(page):
    process_details = page.locator('[role="dialog"]:has-text("Process Details")').first
    submit_candidates = [
        process_details.locator('xpath=.//button[normalize-space()="Submit"]'),
        process_details.locator('xpath=.//*[contains(normalize-space(),"Process Options")]/following::*[normalize-space()="Submit"][1]'),
        process_details.get_by_role("button", name="Submit"),
        page.locator('xpath=//*[contains(normalize-space(),"Process Options")]/following::*[normalize-space()="Submit"][1]'),
        page.get_by_role("button", name="Submit"),
    ]
    for group in submit_candidates:
        try:
            if group.count() <= 0:
                continue
            for pos in range(min(group.count(), 5)):
                candidate = group.nth(pos)
                if candidate.is_visible(timeout=1_000):
                    return candidate
        except Exception:
            continue
    return None


def _find_ess_confirmation_ok_button(page):
    popup_candidates = [
        page.locator('[role="dialog"]:has-text("Confirmation")').first,
        page.locator('[role="dialog"]:has-text("Process Details")').first,
        page.locator('[role="dialog"]').last,
    ]
    for popup in popup_candidates:
        try:
            popup.wait_for(state="visible", timeout=2_000)
            ok_button = popup.get_by_role("button", name="OK").first
            if ok_button.count() > 0 and ok_button.is_visible(timeout=1_000):
                return ok_button
            ok_text = popup.locator('xpath=.//*[normalize-space()="OK"]').first
            if ok_text.count() > 0 and ok_text.is_visible(timeout=1_000):
                return ok_text
        except Exception:
            continue

    try:
        global_ok = page.get_by_role("button", name="OK")
        if global_ok.count() > 0:
            for pos in range(global_ok.count() - 1, -1, -1):
                candidate = global_ok.nth(pos)
                if candidate.is_visible(timeout=800):
                    return candidate
    except Exception:
        pass
    return None


def _submit_ess_job_and_capture_confirmation(
    page,
    ess_mod,
    screenshots_dir: Path,
    idx: int,
    screenshot_name: str,
) -> tuple[str, str]:
    log.info("  ➤ Clicking Process Details header Submit")
    process_details = page.locator('[role="dialog"]:has-text("Process Details")').first
    submit_button = _find_ess_submit_button(page)
    if submit_button is None:
        return "FAILED", "Header Submit button not found in Process Details"

    if not _click_ess_locator(page, submit_button, "Process Details header Submit", ess_mod.js_mousedown):
        return "FAILED", "Unable to click Process Details header Submit"

    try:
        ess_mod.wait_busy(page, timeout=30_000)
    except Exception:
        pass
    time.sleep(1.0)

    submit_confirmed = False
    for selector in [
        'text="queued up for submission"',
        'text="This process will be queued up for submission"',
        'text="Request ID"',
        'text="Confirmation"',
    ]:
        try:
            page.locator(selector).first.wait_for(state="visible", timeout=4_000)
            submit_confirmed = True
            break
        except Exception:
            continue

    if not submit_confirmed:
        try:
            process_details.wait_for(state="hidden", timeout=4_000)
            submit_confirmed = True
        except Exception:
            pass

    ok_button = _find_ess_confirmation_ok_button(page)
    if ok_button is not None:
        log.info("  ℹ️ Submit confirmation popup detected")
        _capture_workflow_screenshot(
            page,
            screenshots_dir,
            "ess",
            "s5_process_confirmation_popup_detected_after",
            idx,
            screenshot_name,
            "submitted",
        )
        if _click_ess_locator(page, ok_button, "Submit confirmation OK", ess_mod.js_mousedown):
            try:
                ess_mod.wait_busy(page, timeout=15_000)
            except Exception:
                pass
            time.sleep(0.6)
            log.info("  ✅ Confirmation OK clicked")
        else:
            log.warning("  ⚠️ Confirmation popup found but OK click failed")
            return "FAILED", "Confirmation popup found but OK click failed"
    else:
        log.info("  ℹ️ No post-submit confirmation popup detected")

    if submit_confirmed:
        return "SUCCESS", "Submit confirmed"
    return "SUCCESS", "Submit clicked (no explicit confirmation selector)"


def run_ess(page, ess_mod, entries, ess_log_path: Path, screenshots_dir: Path):
    if not entries:
        log.info("ESS Jobs: no filtered rows to process")
        _append_realtime(ess_log_path, "No ESS job entries to process")
        return {"success": 0, "failed": 0}

    unique_entries = []
    seen_jobs = set()
    skipped_duplicates = 0
    for entry in entries:
        process_name = entry.get("process_name", "")
        param_label = entry.get("param_label", "")
        param_value = entry.get("param_value", "")
        job_key = _execution_key(process_name, param_label, param_value)
        if job_key in seen_jobs:
            skipped_duplicates += 1
            continue
        seen_jobs.add(job_key)
        unique_entries.append(entry)

    if skipped_duplicates:
        log.info("ESS Jobs: skipped duplicate Excel entries: %d", skipped_duplicates)

    if not unique_entries:
        log.info("ESS Jobs: all parsed entries were duplicate rows")
        return {"success": 0, "failed": 0}

    ess_mod.click_navigator(page)
    initial_work_area = ess_mod.capture_work_area_from_breadcrumb(page)
    ess_mod.expand_tools_section(page)
    ess_mod.click_scheduled_processes(page)
    work_area = ess_mod.capture_work_area_from_breadcrumb(page)
    if work_area.lower() == ess_mod._normalize_work_area_text("Tools").lower() and initial_work_area:
        work_area = initial_work_area

    success = failed = 0
    for idx, entry in enumerate(unique_entries, start=1):
        process_name = entry.get("process_name", "")
        param_label = entry.get("param_label", "")
        param_value = entry.get("param_value", "")
        screenshot_name = f"{process_name}_{param_label}_{param_value}"
        try:
            ess_mod.click_schedule_new_process(page)
            _capture_workflow_screenshot(
                page,
                screenshots_dir,
                "ess",
                "s1_schedule_new_process_blank_after",
                idx,
                screenshot_name,
                "scheduled",
            )
            selected = ess_mod.select_process_in_schedule_dialog_direct(page, process_name, direct=True)
            _capture_workflow_screenshot(
                page,
                screenshots_dir,
                "ess",
                "s2_process_selected_in_schedule_dialog_after",
                idx,
                screenshot_name,
                "scheduled" if selected else "error",
            )
            if not selected:
                failed += 1
                _append_realtime(
                    ess_log_path,
                    f"Process='{process_name}' | Param='{param_label}' | Value='{param_value}' | Status=FAIL | Error='Process not selectable'",
                )
                continue
            ess_mod.open_process_details_from_schedule_dialog(page)
            _capture_workflow_screenshot(
                page,
                screenshots_dir,
                "ess",
                "s3_process_details_opened_after",
                idx,
                screenshot_name,
                "scheduled",
            )
            sanitized_param_value = ess_mod.sanitize_parameter_value(param_value)
            result = ess_mod.run_parameter_handling(
                page=page,
                process_name=process_name,
                job_request_id="AUTO",
                work_area=work_area,
                page_name="Scheduled Processes",
                input_parameters=[{"label": param_label, "value": sanitized_param_value}],
                strict_mode=False,
                validation_mode="WARN_AND_HOLD",
            )
            _capture_workflow_screenshot(
                page,
                screenshots_dir,
                "ess",
                "s4_value_filled_after",
                idx,
                screenshot_name,
                "filled",
            )
            submit_status, submit_message = _submit_ess_job_and_capture_confirmation(
                page,
                ess_mod,
                screenshots_dir,
                idx,
                screenshot_name,
            )
            result.update(
                {
                    "submit_status": submit_status,
                    "submit_message": submit_message,
                    "used_param_label": param_label,
                    "used_param_value_original": param_value,
                    "used_param_value_sanitized": sanitized_param_value,
                }
            )
            missing_labels = result.get("extra_input_labels", []) or []
            if missing_labels:
                failed += 1
                missing_text = ", ".join(str(label) for label in missing_labels)
                discovered_text = ", ".join(str(label) for label in result.get("discovered_labels", []) or [])
                _append_realtime(
                    ess_log_path,
                    (
                        f"Process='{process_name}' | Param='{param_label}' | Value='{param_value}' "
                        f"| Status=FAIL | Error='Parameter label not present in UI: {missing_text}' "
                        f"| DiscoveredLabels='{discovered_text}'"
                    ),
                )
            elif str(result.get("submit_status", "FAILED")).upper() == "SUCCESS":
                success += 1
                _append_realtime(
                    ess_log_path,
                    f"Process='{process_name}' | Param='{param_label}' | Value='{param_value}' | Status=PASS",
                )
            else:
                failed += 1
                _append_realtime(
                    ess_log_path,
                    f"Process='{process_name}' | Param='{param_label}' | Value='{param_value}' | Status=FAIL | Error='Submit failed'",
                )
        except Exception as ex:
            failed += 1
            _append_realtime(
                ess_log_path,
                f"Process='{process_name}' | Param='{param_label}' | Value='{param_value}' | Status=FAIL | Error='{str(ex)}'",
            )
    return {"success": success, "failed": failed}


def main():
    optin_mod, profile_mod, ess_mod = _load_inline_modules()
    run_dir, log_files = _init_run_logs()
    screenshots_dir = run_dir / "screenshots"
    screenshots_dir.mkdir(parents=True, exist_ok=True)
    log.info("Log folder created: %s", run_dir)
    log.info("Screenshots folder created: %s", screenshots_dir)
    log.info("Run mode: %s", "All" if RUN_MODE == "all" else RUN_MODE)

    optin_features, profile_entries, ess_entries, auto_entries, feature_rows = load_filtered_data(EXCEL_PATH)
    for fr in feature_rows:
        _append_feature_log(
            log_files["feature"],
            feature_name=fr.get("feature_name", "Unknown Feature"),
            opt_in_output=fr.get("opt_in_output", ""),
            profile_option_output=fr.get("profile_option_output", ""),
            ess_jobs_output=fr.get("ess_jobs_output", ""),
            stage="INPUT_VALUES",
            source_row=fr.get("source_row", ""),
        )
    log.info("Feature rows written to feature-wise log: %d", len(feature_rows))
    log.info("Input rows available for selected run mode:")
    if _should_run("optin"):
        log.info("  Opt-in features   : %d", len(optin_features))
    if _should_run("profile"):
        log.info("  Profile entries   : %d", len(profile_entries))
    if _should_run("ess"):
        log.info("  ESS job entries   : %d", len(ess_entries))
    if _should_run("auto"):
        log.info("  Auto entries      : %d", len(auto_entries))

    optin_results = []
    profile_pass = profile_fail = 0
    ess_summary = {"success": 0, "failed": 0}
    auto_summary = {"processed": 0, "enabled": 0}

    should_use_browser = any(_should_run(section) for section in ("optin", "profile", "ess"))
    with sync_playwright() as pw:
        browser = None
        page = None
        if should_use_browser:
            browser = pw.chromium.launch(
                headless=BROWSER_HEADLESS,
                slow_mo=BROWSER_SLOW_MO_MS,
                args=["--disable-dev-shm-usage"],
            )
            context = browser.new_context(ignore_https_errors=True)
            page = context.new_page()
            page.set_default_timeout(60_000)

        try:
            if should_use_browser:
                login(page)

            if _should_run("optin"):
                optin_results = run_optin(page, optin_mod, optin_features, log_files["optin"], screenshots_dir)

            if _should_run("profile"):
                profile_pass, profile_fail = run_profile(page, profile_mod, profile_entries, log_files["profile"], screenshots_dir)

            if _should_run("ess"):
                ess_summary = run_ess(page, ess_mod, ess_entries, log_files["ess"], screenshots_dir)

            if _should_run("auto"):
                auto_summary = run_auto(auto_entries, log_files["auto"])

            optin_pass = sum(1 for r in optin_results if getattr(r, "status", "") in {"Already Enabled", "Enabled Now"})
            optin_fail = len(optin_results) - optin_pass
            ess_pass = int(ess_summary.get("success", 0))
            ess_fail = int(ess_summary.get("failed", 0))
            auto_processed = int(auto_summary.get("processed", 0))
            auto_enabled = int(auto_summary.get("enabled", auto_processed))
            total_pass = optin_pass + profile_pass + ess_pass + auto_enabled
            total_fail = optin_fail + profile_fail + ess_fail

            log.info("=" * 80)
            log.info("COMBINED SUMMARY")
            log.info("Run mode: %s", "All" if RUN_MODE == "all" else RUN_MODE)
            if _should_run("optin"):
                log.info("Opt-in processed: %d", len(optin_results))
                log.info("Opt-in pass/fail: %d/%d", optin_pass, optin_fail)
            if _should_run("profile"):
                log.info("Profile pass/fail: %d/%d", profile_pass, profile_fail)
            if _should_run("ess"):
                log.info("ESS success/fail: %d/%d", ess_summary["success"], ess_summary["failed"])
            if _should_run("auto"):
                log.info("Auto enabled/logged: %d", auto_enabled)
            log.info("TOTAL pass/fail: %d/%d", total_pass, total_fail)
            log.info("=" * 80)

            with open(log_files["summary"], "a", encoding="utf-8") as f:
                f.write("COMBINED SUMMARY\n")
                f.write("=" * 100 + "\n")
                f.write(f"Run Mode      : {'All' if RUN_MODE == 'all' else RUN_MODE}\n")
                f.write(f"Feature Rows Loaded : {len(feature_rows)}\n")
                f.write(f"Feature Rows Logged : {len(feature_rows)}\n")
                f.write("-" * 100 + "\n")
                if _should_run("optin"):
                    f.write(f"Opt-in Total  : {len(optin_results)}\n")
                    f.write(f"Opt-in Passed : {optin_pass}\n")
                    f.write(f"Opt-in Failed : {optin_fail}\n")
                    f.write("-" * 100 + "\n")
                if _should_run("profile"):
                    f.write(f"Profile Passed: {profile_pass}\n")
                    f.write(f"Profile Failed: {profile_fail}\n")
                    f.write("-" * 100 + "\n")
                if _should_run("ess"):
                    f.write(f"ESS Passed    : {ess_pass}\n")
                    f.write(f"ESS Failed    : {ess_fail}\n")
                    f.write("-" * 100 + "\n")
                if _should_run("auto"):
                    f.write(f"Auto Enabled/Logged: {auto_enabled}\n")
                    f.write("-" * 100 + "\n")
                f.write(f"TOTAL Passed  : {total_pass}\n")
                f.write(f"TOTAL Failed  : {total_fail}\n")
                f.write(f"Completed At  : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            log.info("Summary log generated: %s", log_files["summary"])
        finally:
            try:
                if os.isatty(0):
                    input("Press Enter to close browser...")
            except Exception:
                time.sleep(1)
            if browser:
                browser.close()


if __name__ == "__main__":
    main()
