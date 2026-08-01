from __future__ import annotations

import datetime as dt
import math
import re
import unicodedata
from collections.abc import Iterable, Mapping
from typing import Any

from openpyxl.utils.datetime import from_excel


_CONCEPT_GROUPS: tuple[tuple[str, ...], ...] = (
    ("recommend", "recommendation", "recommended", "recommends"),
    ("construction gross profit", "construction gp"),
    ("interest rate cap", "rate cap", "purchase a cap", "purchasing a cap"),
    ("latest forecast", "latest outlook", "latest approved outlook", "latest reforecast", "current outlook", "latest rf", "fy26 rf"),
    ("model status", "overall control status", "control status", "overall status", "control summary", "all formula driven controls"),
    ("controller tied", "controller tie", "controller tie complete", "controller reviewed"),
    ("minimum balance", "minimum operating balance", "bank minimum balance", "bank minimum funding"),
    (
        "management correspondence", "management direction", "management clarification",
        "final management clarification", "post committee clarification", "controlling correspondence",
        "review correspondence", "management request thread", "assignment and source tie direction",
    ),
    ("contractor accounting mcp", "accounting mcp", "mcp get cash balances", "mcp cash balances", "mcp source"),
    ("period", "reporting period", "as of date", "as of", "planning cutoff", "reporting cutoff", "fiscal year"),
    ("scenario", "case", "risk case"),
    ("unit", "units"),
    ("policy", "treasury policy", "accounting policy", "executed policy"),
    ("trailing four quarter", "trailing 4 quarter", "rolling 4q", "rolling t4q", "rolling twelve month", "t4q"),
    ("depreciation and amortization", "d and a", "d a"),
    ("working capital", "net working capital", "nwc", "change in nwc"),
    ("purchase price", "purchase enterprise value", "transaction value"),
    ("executive summary", "executive overview", "what matters now"),
    ("cash roll forward", "cash rollforward", "cash bridge", "fcf cash nwc"),
    ("equipment line", "equipment lines"),
    ("do not approve as structured", "reject as submitted", "reject the structure", "do not approve", "not approved"),
    ("stock acquisition", "stock purchase", "stock purchase without section 338 election", "stock deal"),
)


def normalize_text(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode()
    text = text.casefold().replace("&", " and ")
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def _singular(token: str) -> str:
    if len(token) <= 3 or token.endswith(("ss", "us", "is")):
        return token
    if token.endswith("ies") and len(token) > 4:
        return token[:-3] + "y"
    if token.endswith(("ches", "shes", "xes", "zes")):
        return token[:-2]
    if token.endswith("s"):
        return token[:-1]
    return token


def semantic_tokens(value: Any) -> list[str]:
    return [_singular(token) for token in normalize_text(value).split()]


def _concept_aliases(value: Any) -> set[str]:
    normalized = normalize_text(value)
    aliases = {normalized}
    for group in _CONCEPT_GROUPS:
        normalized_group = {normalize_text(item) for item in group}
        if normalized in normalized_group:
            aliases.update(normalized_group)
    return aliases


def contains_concept(text: Any, concept: Any) -> bool:
    haystack_tokens = semantic_tokens(text)
    if not haystack_tokens:
        return False
    for alias in _concept_aliases(concept):
        wanted = semantic_tokens(alias)
        if wanted and any(haystack_tokens[index:index + len(wanted)] == wanted for index in range(len(haystack_tokens))):
            return True
    return False


def _contains_nonnegated_concept(text: Any, concept: Any) -> bool:
    tokens = semantic_tokens(text)
    for alias in _concept_aliases(concept):
        wanted = semantic_tokens(alias)
        for index in range(len(tokens) - len(wanted) + 1):
            if tokens[index:index + len(wanted)] != wanted:
                continue
            prefix = tokens[max(0, index - 3):index]
            if not any(token in {"not", "no", "avoid", "reject", "exclude"} for token in prefix):
                return True
    return False


def semantic_equal(actual: Any, expected: Any) -> bool:
    if normalize_text(actual) == normalize_text(expected):
        return True
    if semantic_tokens(actual) == semantic_tokens(expected):
        return True
    if _concept_aliases(actual) & _concept_aliases(expected):
        return True
    if _contains_nonnegated_concept(actual, expected):
        return True
    return False


def iter_numeric_candidates(value: Any) -> Iterable[Any]:
    if isinstance(value, bool):
        return
    if isinstance(value, (int, float)) and math.isfinite(float(value)):
        yield value
        return
    if isinstance(value, str):
        yield value
        return
    if isinstance(value, Mapping):
        preferred = ("value", "amount", "ratio", "percentage", "percent", "total", "result")
        for key in preferred:
            if key in value:
                yield from iter_numeric_candidates(value[key])
        for key, child in value.items():
            if key not in preferred:
                yield from iter_numeric_candidates(child)
        return
    if isinstance(value, (list, tuple, set)):
        for child in value:
            yield from iter_numeric_candidates(child)


def semantic_segments(value: Any, *, include_keys: bool = False) -> list[str]:
    if value is None:
        return []
    if isinstance(value, Mapping):
        children: list[str] = []
        for key, child in value.items():
            if include_keys:
                children.append(str(key))
            children.extend(semantic_segments(child, include_keys=include_keys))
        return children + ([" ".join(children)] if children else [])
    if isinstance(value, (list, tuple, set)):
        segments: list[str] = []
        for child in value:
            child_segments = semantic_segments(child, include_keys=include_keys)
            segments.extend(child_segments)
            if child_segments:
                segments.append(" ".join(child_segments))
        return segments
    return [str(value)]


def _meaningful_tokens(value: Any) -> list[str]:
    stop = {"the", "a", "an", "of", "and", "for", "recommended", "recommendation"}
    return [token for token in semantic_tokens(value) if token not in stop]


def _numeric_literals(value: Any) -> list[float]:
    values: list[float] = []
    for token in re.findall(r"\(?-?\$?\d[\d,]*(?:\.\d+)?\s*(?:%|[kmb])?\)?", str(value or ""), flags=re.I):
        text = token.strip().replace("$", "").replace(",", "")
        negative = text.startswith("(") and text.endswith(")")
        text = text.strip("() ")
        percent = text.endswith("%")
        suffix = text[-1:].casefold()
        multiplier = {"k": 1_000.0, "m": 1_000_000.0, "b": 1_000_000_000.0}.get(suffix, 1.0)
        if suffix in {"k", "m", "b"}:
            text = text[:-1]
        text = text.rstrip("%").strip()
        try:
            number = float(text) * multiplier
        except ValueError:
            continue
        values.append((-number if negative else number) / (100 if percent else 1))
    return values


def semantic_value_matches(actual: Any, expected: str) -> bool:
    segments = semantic_segments(actual, include_keys=True)
    if any(semantic_equal(segment, expected) for segment in segments):
        return True
    joined = " ".join(segments)
    if _contains_nonnegated_concept(joined, expected):
        return True
    expected_numbers = _numeric_literals(expected)
    actual_numbers = [number for candidate in iter_numeric_candidates(actual) for number in _numeric_literals(candidate)]
    numbers_match = not expected_numbers or all(
        any(abs(actual_number - expected_number) <= max(0.02, abs(expected_number) * 1e-6) for actual_number in actual_numbers)
        for expected_number in expected_numbers
    )
    wanted = [token for token in _meaningful_tokens(expected) if not token.isdigit()]
    present = set(_meaningful_tokens(joined))
    return (
        numbers_match
        and bool(wanted)
        and all(token in present for token in wanted)
        and not contains_concept(joined, expected)
    )


def ordered_semantic_list_matches(actual: Any, expected: list[str]) -> bool:
    if not expected:
        return not semantic_segments(actual)
    if isinstance(actual, (list, tuple)):
        position = 0
        for child in actual:
            if semantic_value_matches(child, expected[position]):
                position += 1
                if position == len(expected):
                    return True
        for segment in semantic_segments(actual, include_keys=True):
            if ordered_semantic_list_matches(segment, expected):
                return True
        return False
    normalized = semantic_tokens(" ".join(semantic_segments(actual, include_keys=True)))
    cursor = 0
    for item in expected:
        matches: list[int] = []
        for alias in _concept_aliases(item):
            wanted = semantic_tokens(alias)
            matches.extend(index for index in range(cursor, len(normalized) - len(wanted) + 1) if normalized[index:index + len(wanted)] == wanted)
        if not matches:
            return False
        cursor = min(matches) + 1
    return True


def unordered_semantic_list_matches(actual: Any, expected: list[str]) -> bool:
    """Match a set-like list without requiring presentation order.

    Rich objects remain acceptable because each top-level child is evaluated
    through ``semantic_value_matches``.  A one-to-one assignment prevents one
    vague child from satisfying multiple expected concepts, and extra list
    members remain a failure for an exact output contract.
    """
    if not expected:
        return isinstance(actual, (list, tuple, set)) and not actual
    if not isinstance(actual, (list, tuple, set)) or len(actual) != len(expected):
        return False
    children = list(actual)

    def assign(index: int, remaining: tuple[int, ...]) -> bool:
        if index == len(expected):
            return True
        return any(
            semantic_value_matches(children[candidate], expected[index])
            and assign(index + 1, tuple(value for value in remaining if value != candidate))
            for candidate in remaining
        )

    return assign(0, tuple(range(len(children))))


def normalize_date(value: Any) -> dt.date | None:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)) and 1 <= float(value) <= 100_000:
        try:
            converted = from_excel(float(value))
            return converted.date() if isinstance(converted, dt.datetime) else converted
        except (OverflowError, ValueError, TypeError):
            return None
    text = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%b %d, %Y", "%B %d, %Y"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def date_matches(actual: Any, expected: Any) -> bool:
    expected_date = normalize_date(expected)
    candidates: list[Any] = [actual]
    candidates.extend(semantic_segments(actual))
    return expected_date is not None and any(normalize_date(candidate) == expected_date for candidate in candidates)
