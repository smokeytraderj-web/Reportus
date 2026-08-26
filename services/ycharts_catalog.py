"""Search and validate a locally supplied YCharts Excel-function reference."""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from xml.etree import ElementTree as ET


_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS = {"m": _MAIN, "r": _REL}
_TOKEN = re.compile(r"[a-z0-9]+")
_STOPWORDS = frozenset({
    "and", "are", "build", "chart", "charts", "create", "data", "excel", "for",
    "from", "have", "into", "make", "report", "sheet", "show", "that", "the",
    "this", "use", "with", "workbook", "ycharts",
})
_QUERY_EXPANSIONS = {
    "performance": ("return", "total return", "price return"),
    "returns": ("return", "total return"),
    "yield": ("dividend yield", "distribution yield"),
    "valuation": ("price earnings", "enterprise value", "ebitda"),
    "risk": ("volatility", "drawdown", "beta", "sharpe"),
    "price": ("price", "close"),
    "aum": ("assets under management",),
    "expense": ("expense ratio",),
    "sector": ("sector", "sector exposure"),
}


class YChartsCatalogError(RuntimeError):
    pass


@dataclass(frozen=True, slots=True)
class YChartsMetric:
    name: str
    code: str
    function: str
    category: str
    entities: tuple[str, ...]
    value_type: str = ""
    description: str = ""
    financial_statement: str = ""
    forward_estimate: str = ""

    def prompt_line(self) -> str:
        entity = "/".join(self.entities)
        details = f"; {self.value_type}" if self.value_type else ""
        return f"{self.function} | {self.code or '(security value)'} | {self.name} | {entity}{details}"


def _tokens(value: str) -> set[str]:
    return {
        token for token in _TOKEN.findall(value.casefold())
        if len(token) > 1 and token not in _STOPWORDS
    }


def _column(reference: str) -> int:
    match = re.match(r"[A-Z]+", reference)
    if match is None:
        raise YChartsCatalogError("The YCharts reference contains an invalid cell address.")
    result = 0
    for letter in match.group(0):
        result = result * 26 + ord(letter) - 64
    return result


def _rows(archive: zipfile.ZipFile, target: str, shared: list[str]):
    root = ET.fromstring(archive.read(target))
    for row in root.findall(".//m:sheetData/m:row", _NS):
        values: dict[int, str] = {}
        for cell in row.findall("m:c", _NS):
            value_node = cell.find("m:v", _NS)
            inline = cell.find("m:is", _NS)
            if cell.attrib.get("t") == "s" and value_node is not None:
                value = shared[int(value_node.text or "0")]
            elif cell.attrib.get("t") == "inlineStr" and inline is not None:
                value = "".join(node.text or "" for node in inline.iterfind(".//m:t", _NS))
            else:
                value = value_node.text if value_node is not None else ""
            values[_column(cell.attrib["r"])] = value or ""
        if values:
            yield [values.get(index, "") for index in range(1, max(values) + 1)]


def _read_reference(path) -> tuple[YChartsMetric, ...]:
    if path.suffix.lower() not in {".xlsx", ".xlsm"} or not path.is_file():
        raise YChartsCatalogError("Upload the YCharts Complete Excel Reference as .xlsx or .xlsm.")
    try:
        with zipfile.ZipFile(path) as archive:
            if "xl/sharedStrings.xml" in archive.namelist():
                strings_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                shared = [
                    "".join(node.text or "" for node in item.iterfind(".//m:t", _NS))
                    for item in strings_root.findall("m:si", _NS)
                ]
            else:
                shared = []
            rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            relationships = {item.attrib["Id"]: item.attrib["Target"] for item in rels_root}
            workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
            sheets_root = workbook_root.find("m:sheets", _NS)
            raw_entries: list[YChartsMetric] = []
            for sheet in sheets_root if sheets_root is not None else ():
                name = sheet.attrib["name"]
                if " | " not in name:
                    continue
                target = relationships[sheet.attrib[f"{{{_REL}}}id"]].lstrip("/")
                if not target.startswith("xl/"):
                    target = "xl/" + target
                rows = list(_rows(archive, target, shared))
                if len(rows) < 3:
                    continue
                entity, category = name.split(" | ", 1)
                headers = {
                    value.strip().casefold(): index
                    for index, value in enumerate(rows[1]) if value
                }
                for row in rows[2:]:
                    def value(header: str) -> str:
                        index = headers.get(header.casefold())
                        return row[index].strip() if index is not None and index < len(row) else ""

                    metric_name = value("Metric Name")
                    syntax = value("Syntax")
                    function_match = re.match(r"\s*([A-Z]+)\(", syntax)
                    if not metric_name or function_match is None:
                        continue
                    raw_entries.append(YChartsMetric(
                        name=metric_name,
                        code=value("Metric Code"),
                        function=function_match.group(1),
                        category=category.casefold(),
                        entities=(entity,),
                        value_type=value("Type"),
                        description=value("Description"),
                        financial_statement=value("Financial Statement"),
                        forward_estimate=value("Forward Estimate Available"),
                    ))
    except (OSError, KeyError, ValueError, zipfile.BadZipFile, ET.ParseError) as exc:
        raise YChartsCatalogError("The selected YCharts reference workbook could not be read.") from exc

    combined: dict[tuple[object, ...], YChartsMetric] = {}
    for metric in raw_entries:
        key = (
            metric.category, metric.function, metric.code, metric.name, metric.value_type,
            metric.description, metric.financial_statement, metric.forward_estimate,
        )
        previous = combined.get(key)
        if previous is None:
            combined[key] = metric
        elif metric.entities[0] not in previous.entities:
            combined[key] = YChartsMetric(
                previous.name, previous.code, previous.function, previous.category,
                previous.entities + metric.entities, previous.value_type, previous.description,
                previous.financial_statement, previous.forward_estimate,
            )
    metrics = tuple(combined.values())
    if not any(metric.function == "YCP" for metric in metrics) or not any(
        metric.function == "YCI" for metric in metrics
    ):
        raise YChartsCatalogError("This is not the YCharts Complete Excel Reference workbook.")
    return metrics


class YChartsCatalog:
    def __init__(self, metrics: tuple[YChartsMetric, ...] = ()):
        self.metrics = metrics
        self._valid = {
            (metric.function, metric.code.casefold()) for metric in metrics if metric.code
        }

    @classmethod
    def from_reference_workbook(cls, path) -> "YChartsCatalog":
        return cls(_read_reference(path))

    def contains(self, function: str, code: str) -> bool:
        normalized = function.upper()
        if normalized in {"YCS", "YCD", "YCDS"}:
            normalized = "YCP"
        return (normalized, code.casefold()) in self._valid

    def search(self, query: str, *, limit: int = 100) -> tuple[YChartsMetric, ...]:
        if not 1 <= limit <= 250:
            raise ValueError("YCharts search limit must be between 1 and 250.")
        query_tokens = _tokens(query)
        expanded = set(query_tokens)
        for token in tuple(query_tokens):
            for phrase in _QUERY_EXPANSIONS.get(token, ()):
                expanded.update(_tokens(phrase))
        normalized_query = " ".join(sorted(expanded))
        scored: list[tuple[float, str, str, YChartsMetric, set[str]]] = []
        for metric in self.metrics:
            haystack = " ".join((
                metric.name, metric.code.replace("_", " "), metric.description,
                metric.financial_statement, " ".join(metric.entities),
            )).casefold()
            metric_tokens = _tokens(haystack)
            code_tokens = _tokens(metric.code.replace("_", " "))
            overlap = expanded & metric_tokens
            if not overlap:
                continue
            score = sum(4 if token in metric.code.casefold() else 2 for token in overlap)
            if code_tokens and code_tokens.issubset(expanded):
                score += 10
            score -= .5 * len(code_tokens - expanded)
            if normalized_query and normalized_query in haystack:
                score += 8
            if metric.category == "metrics":
                score += .25
            scored.append((-score, metric.name.casefold(), metric.code.casefold(), metric, metric_tokens))
        scored.sort(key=lambda item: item[:3])
        result: list[YChartsMetric] = []
        seen: set[tuple[str, str, str]] = set()

        def add(metric: YChartsMetric) -> None:
            key = (metric.function, metric.code.casefold(), metric.name.casefold())
            if key not in seen and len(result) < limit:
                result.append(metric)
                seen.add(key)

        candidates_by_token = {
            token: [metric for _, _, _, metric, metric_tokens in scored if token in metric_tokens]
            for token in sorted(expanded)
        }
        for round_number in range(4):
            for candidates in candidates_by_token.values():
                if round_number < len(candidates):
                    add(candidates[round_number])
                if len(result) == limit:
                    return tuple(result)
        for _, _, _, metric, _ in scored:
            add(metric)
            if len(result) == limit:
                break
        return tuple(result)

    def prompt_reference(self, query: str, *, limit: int = 100) -> str:
        matches = self.search(query, limit=limit)
        if not matches:
            return "No YCharts metric candidates matched the request. Do not invent a metric code."
        return (
            "Approved YCharts metric candidates from the locally supplied Complete Excel Reference:\n"
            + "\n".join(metric.prompt_line() for metric in matches)
        )
