"""Strict structured-output contract for the Client Deck generator."""

from __future__ import annotations

from generators.client_deck import AllocationRow, ClientDeckData


class ClientDeckContentError(RuntimeError):
    pass


def client_deck_schema() -> dict[str, object]:
    metric_object = {"type": "object", "additionalProperties": {"type": "number"}}
    return {
        "type": "object",
        "required": [
            "allocation", "risk_metrics", "sector_performance", "sector_portfolio",
            "sector_benchmark", "contributors", "detractors", "earnings_years",
            "earnings_values", "earnings_notes", "optional_sections", "sources",
        ],
        "properties": {
            "allocation": {
                "type": "array", "minItems": 1, "maxItems": 12,
                "items": {
                    "type": "object", "required": ["label", "value"],
                    "properties": {"label": {"type": "string"}, "value": {"type": "number"}},
                },
            },
            "risk_metrics": {"type": "object", "additionalProperties": {"type": "string"}},
            "sector_performance": metric_object,
            "sector_portfolio": metric_object,
            "sector_benchmark": metric_object,
            "contributors": {"type": "array", "items": {"type": "object"}},
            "detractors": {"type": "array", "items": {"type": "object"}},
            "earnings_years": {"type": "array", "items": {"type": "string"}},
            "earnings_values": {"type": "array", "items": {"type": "number"}},
            "earnings_notes": {"type": "array", "items": {"type": "string"}},
            "optional_sections": {"type": "object"},
            "sources": {"type": "object", "additionalProperties": {"type": "string"}},
        },
    }


def _text(value: object, field: str, *, limit: int = 180) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ClientDeckContentError(f"{field} is required.")
    text = value.strip()
    if len(text) > limit:
        raise ClientDeckContentError(f"{field} is too long.")
    return text


def _number(value: object, field: str, *, minimum: float | None = None,
            maximum: float | None = None) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ClientDeckContentError(f"{field} must be numeric.")
    result = float(value)
    if minimum is not None and result < minimum:
        raise ClientDeckContentError(f"{field} is below the allowed range.")
    if maximum is not None and result > maximum:
        raise ClientDeckContentError(f"{field} is above the allowed range.")
    return result


def _mapping(value: object, field: str) -> dict[str, object]:
    if not isinstance(value, dict) or not value:
        raise ClientDeckContentError(f"{field} must be a non-empty object.")
    return value


def _numeric_mapping(value: object, field: str, *, minimum: float, maximum: float,
                     max_items: int = 15) -> dict[str, float]:
    mapping = _mapping(value, field)
    if len(mapping) > max_items:
        raise ClientDeckContentError(f"{field} contains too many rows.")
    return {
        _text(label, f"{field} label", limit=50): _number(number, f"{field} {label}", minimum=minimum, maximum=maximum)
        for label, number in mapping.items()
    }


def _attribution_rows(value: object, field: str) -> tuple[tuple[str, str, str, str], ...]:
    if not isinstance(value, list) or not 1 <= len(value) <= 10:
        raise ClientDeckContentError(f"{field} requires 1–10 rows.")
    rows = []
    for index, item in enumerate(value, start=1):
        if not isinstance(item, dict):
            raise ClientDeckContentError(f"{field} row {index} must be an object.")
        total_return = _number(item.get("return"), f"{field} row {index} return", minimum=-10, maximum=10)
        contribution = _number(item.get("contribution"), f"{field} row {index} contribution", minimum=-10, maximum=10)
        rows.append((
            _text(item.get("symbol"), f"{field} row {index} symbol", limit=15),
            _text(item.get("holding"), f"{field} row {index} holding", limit=60),
            f"{total_return:+.2%}",
            f"{contribution:+.2%}",
        ))
    return tuple(rows)


def normalize_client_deck_payload(payload: dict[str, object], *, client_name: str,
                                  period: str, as_of: str) -> ClientDeckData:
    if not isinstance(payload, dict):
        raise ClientDeckContentError("Client Deck output must be one JSON object.")
    allocation_value = payload.get("allocation")
    if not isinstance(allocation_value, list) or not 1 <= len(allocation_value) <= 12:
        raise ClientDeckContentError("allocation requires 1–12 rows.")
    allocation = []
    for index, item in enumerate(allocation_value, start=1):
        if not isinstance(item, dict):
            raise ClientDeckContentError(f"allocation row {index} must be an object.")
        allocation.append(AllocationRow(
            _text(item.get("label"), f"allocation row {index} label", limit=50),
            _number(item.get("value"), f"allocation row {index} value", minimum=0),
        ))
    if sum(row.value for row in allocation) <= 0:
        raise ClientDeckContentError("allocation total must be greater than zero.")

    risk_value = _mapping(payload.get("risk_metrics"), "risk_metrics")
    required_risk = {"Portfolio total", "Risk", "Annual dividend", "Max drawdown"}
    if not required_risk.issubset(risk_value):
        raise ClientDeckContentError(
            "risk_metrics requires Portfolio total, Risk, Annual dividend, and Max drawdown."
        )
    if len(risk_value) > 8:
        raise ClientDeckContentError("risk_metrics contains too many values.")
    risk_metrics = {
        _text(key, "risk metric label", limit=40): _text(value, f"risk metric {key}", limit=40)
        for key, value in risk_value.items()
    }

    sector_performance = _numeric_mapping(
        payload.get("sector_performance"), "sector_performance", minimum=-2, maximum=2
    )
    sector_portfolio = _numeric_mapping(
        payload.get("sector_portfolio"), "sector_portfolio", minimum=0, maximum=100
    )
    sector_benchmark = _numeric_mapping(
        payload.get("sector_benchmark"), "sector_benchmark", minimum=0, maximum=100
    )
    if set(sector_portfolio) != set(sector_benchmark):
        raise ClientDeckContentError("sector_portfolio and sector_benchmark must use identical sector labels.")

    years_value = payload.get("earnings_years")
    values_value = payload.get("earnings_values")
    notes_value = payload.get("earnings_notes")
    if not isinstance(years_value, list) or not isinstance(values_value, list) or not 2 <= len(years_value) <= 8:
        raise ClientDeckContentError("earnings_years and earnings_values require 2–8 values.")
    if len(years_value) != len(values_value):
        raise ClientDeckContentError("earnings_years and earnings_values must have equal length.")
    if not isinstance(notes_value, list) or not 1 <= len(notes_value) <= 6:
        raise ClientDeckContentError("earnings_notes requires 1–6 notes.")

    optional_value = payload.get("optional_sections")
    if not isinstance(optional_value, dict):
        raise ClientDeckContentError("optional_sections must be an object, even when empty.")
    if set(optional_value) - {"rmd", "529", "annuity"}:
        raise ClientDeckContentError("optional_sections contains an unsupported section.")
    optional: dict[str, tuple[str, ...]] = {}
    for key, lines in optional_value.items():
        if not isinstance(lines, list) or not 1 <= len(lines) <= 8:
            raise ClientDeckContentError(f"optional section {key} requires 1–8 lines.")
        optional[key] = tuple(_text(line, f"optional section {key}", limit=180) for line in lines)

    sources_value = _mapping(payload.get("sources"), "sources")
    required_sources = {
        "allocation", "risk", "sector_performance", "sector_exposure", "attribution", "earnings"
    }
    if not required_sources.issubset(sources_value):
        raise ClientDeckContentError("sources is missing one or more required section citations.")
    sources = {
        _text(key, "source section", limit=40): _text(value, f"source {key}", limit=140)
        for key, value in sources_value.items()
    }

    return ClientDeckData(
        client_name=_text(client_name, "client name", limit=80),
        period=_text(period, "review period", limit=60),
        as_of=_text(as_of, "as-of label", limit=80),
        allocation=tuple(allocation),
        risk_metrics=risk_metrics,
        sector_performance=sector_performance,
        sector_portfolio=sector_portfolio,
        sector_benchmark=sector_benchmark,
        contributors=_attribution_rows(payload.get("contributors"), "contributors"),
        detractors=_attribution_rows(payload.get("detractors"), "detractors"),
        earnings_years=tuple(_text(item, "earnings year", limit=12) for item in years_value),
        earnings_values=tuple(_number(item, "earnings value", minimum=-1_000_000, maximum=1_000_000) for item in values_value),
        earnings_notes=tuple(_text(item, "earnings note", limit=180) for item in notes_value),
        optional_sections=optional,
        sources=sources,
    )
