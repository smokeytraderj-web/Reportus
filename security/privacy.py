"""Fail-closed local privacy inspection for Reporticles uploads."""

from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path
from xml.etree import ElementTree


class ProhibitedCategory(StrEnum):
    ACCOUNT_IDENTIFIER = "account identifier"
    GOVERNMENT_IDENTIFIER = "government or tax identifier"
    DATE_OF_BIRTH = "date of birth"
    STREET_ADDRESS = "street address"
    PHONE_NUMBER = "phone number"
    EMAIL_ADDRESS = "email address"
    CREDENTIAL = "credential or secret"
    ROUTING_INFORMATION = "bank-routing information"


@dataclass(frozen=True, slots=True)
class PrivacyFinding:
    """A safe finding that never contains the matched sensitive value."""

    category: ProhibitedCategory
    source: str
    location: str


@dataclass(frozen=True, slots=True)
class PrivacyScanResult:
    """Complete local privacy decision for an upload batch."""

    approved: bool
    findings: tuple[PrivacyFinding, ...] = ()
    errors: tuple[str, ...] = ()


class PrivacyScanError(RuntimeError):
    """Raised internally when content cannot be inspected safely."""


_PATTERNS: tuple[tuple[ProhibitedCategory, re.Pattern[str]], ...] = (
    (
        ProhibitedCategory.EMAIL_ADDRESS,
        re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.IGNORECASE),
    ),
    (
        ProhibitedCategory.GOVERNMENT_IDENTIFIER,
        re.compile(r"\b(?:\d{3}[- ]?\d{2}[- ]?\d{4}|(?:SSN|TIN|EIN)\s*[:#-]?\s*[A-Z0-9-]{4,})\b", re.IGNORECASE),
    ),
    (
        ProhibitedCategory.PHONE_NUMBER,
        re.compile(r"(?<!\d)(?:\+?1[-. ()]*)?(?:\d{3}[-. ()]*)\d{3}[-. ]*\d{4}(?!\d)"),
    ),
    (
        ProhibitedCategory.DATE_OF_BIRTH,
        re.compile(r"\b(?:DOB|date\s+of\s+birth|birth\s*date)\b", re.IGNORECASE),
    ),
    (
        ProhibitedCategory.ROUTING_INFORMATION,
        re.compile(r"\b(?:routing|ABA)\s*(?:number|no\.?|#)?\s*[:#-]?\s*\d{9}\b", re.IGNORECASE),
    ),
    (
        ProhibitedCategory.ACCOUNT_IDENTIFIER,
        re.compile(
            r"(?:\b(?:account|acct|brokerage)\s*(?:(?:number|no\.?|#|ending(?:\s+in)?)\s*[:#-]?\s*)?(?:[X*•….-]*\s*)?[A-Z-]*\d[A-Z0-9-]{2,}\b|(?:[Xx*•…]{1,}\s*)\d{3,}\b)",
            re.IGNORECASE,
        ),
    ),
    (
        ProhibitedCategory.CREDENTIAL,
        re.compile(
            r"\b(?:password|passcode|api[_ -]?key|secret|access[_ -]?token|private[_ -]?key)\s*[:=]\s*\S+",
            re.IGNORECASE,
        ),
    ),
    (
        ProhibitedCategory.STREET_ADDRESS,
        re.compile(
            r"\b\d{1,6}\s+[A-Za-z0-9.' -]{2,40}\s+(?:Street|St\.?|Road|Rd\.?|Avenue|Ave\.?|Boulevard|Blvd\.?|Lane|Ln\.?|Drive|Dr\.?|Court|Ct\.?|Way|Parkway|Pkwy\.?)\b",
            re.IGNORECASE,
        ),
    ),
)

_TEXT_EXTENSIONS = {".txt", ".md", ".csv", ".tsv", ".json", ".xml", ".html", ".htm"}
_OFFICE_EXTENSIONS = {".docx", ".xlsx", ".xlsm", ".pptx"}
_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
_LEGACY_OFFICE_EXTENSIONS = {".doc", ".xls", ".ppt"}


def _safe_location(text: str, offset: int) -> str:
    """Return a line reference without echoing matched content."""

    return f"line {text.count(chr(10), 0, offset) + 1}"


def _visible_xml_text(raw_xml: bytes) -> str:
    try:
        root = ElementTree.fromstring(raw_xml)
    except ElementTree.ParseError as exc:
        raise PrivacyScanError("contains malformed XML") from exc
    return "\n".join(part.strip() for part in root.itertext() if part.strip())


class PrivacyScanner:
    """Inspect supported uploads locally and reject a batch on any uncertainty."""

    def scan_text(self, text: str, *, source: str = "typed request") -> PrivacyScanResult:
        """Apply the same fail-closed policy to text entered inside Reporticles."""

        if not text.strip():
            return PrivacyScanResult(False, errors=("No text was provided.",))
        findings = tuple(self._scan_text(source, source, text))
        return PrivacyScanResult(approved=not findings, findings=findings)

    def scan_files(self, paths: list[Path] | tuple[Path, ...]) -> PrivacyScanResult:
        findings: list[PrivacyFinding] = []
        errors: list[str] = []
        if not paths:
            return PrivacyScanResult(False, errors=("No files were provided.",))

        for path in paths:
            try:
                findings.extend(self._scan_text(path.name, "filename", path.name))
                chunks = self._extract_text(path)
                for source, text in chunks:
                    findings.extend(self._scan_text(path.name, source, text))
            except (OSError, PrivacyScanError, zipfile.BadZipFile) as exc:
                errors.append(f"{path.name}: {exc}")

        return PrivacyScanResult(
            approved=not findings and not errors,
            findings=tuple(findings),
            errors=tuple(errors),
        )

    def _scan_text(self, filename: str, source: str, text: str) -> list[PrivacyFinding]:
        findings: list[PrivacyFinding] = []
        seen: set[tuple[ProhibitedCategory, str]] = set()
        for category, pattern in _PATTERNS:
            for match in pattern.finditer(text):
                location = f"{source}, {_safe_location(text, match.start())}"
                key = (category, location)
                if key not in seen:
                    findings.append(PrivacyFinding(category, filename, location))
                    seen.add(key)
        return findings

    def _extract_text(self, path: Path) -> list[tuple[str, str]]:
        if not path.is_file():
            raise PrivacyScanError("file is missing")
        extension = path.suffix.lower()
        if extension in _TEXT_EXTENSIONS:
            return [("document", path.read_text(encoding="utf-8", errors="replace"))]
        if extension in _OFFICE_EXTENSIONS:
            return self._extract_open_xml(path)
        if extension == ".pdf":
            return self._extract_pdf(path)
        if extension in _IMAGE_EXTENSIONS:
            return [("image OCR", self._extract_image_text(path))]
        if extension in _LEGACY_OFFICE_EXTENSIONS:
            raise PrivacyScanError("legacy Office format is not safely inspectable; save as a modern Office file")
        raise PrivacyScanError(f"unsupported file type {extension or '(none)'}")

    def _extract_open_xml(self, path: Path) -> list[tuple[str, str]]:
        chunks: list[tuple[str, str]] = []
        with zipfile.ZipFile(path) as archive:
            encrypted = any(info.flag_bits & 0x1 for info in archive.infolist())
            if encrypted:
                raise PrivacyScanError("encrypted Office files are not allowed")
            for info in archive.infolist():
                name = info.filename
                if not name.endswith(".xml"):
                    continue
                if not name.startswith(("word/", "xl/", "ppt/", "docProps/")):
                    continue
                chunks.append((name, _visible_xml_text(archive.read(info))))
        if not chunks:
            raise PrivacyScanError("Office document contained no inspectable text")
        return chunks

    def _extract_pdf(self, path: Path) -> list[tuple[str, str]]:
        try:
            from pypdf import PdfReader
        except ImportError as exc:
            raise PrivacyScanError("PDF inspection dependency is unavailable") from exc
        reader = PdfReader(path)
        if reader.is_encrypted:
            raise PrivacyScanError("encrypted PDFs are not allowed")
        chunks = []
        for page_number, page in enumerate(reader.pages, start=1):
            text = page.extract_text() or ""
            chunks.append((f"page {page_number}", text))
        if not any(text.strip() for _, text in chunks):
            raise PrivacyScanError("PDF has no extractable text; provide an inspectable copy")
        return chunks

    def _extract_image_text(self, path: Path) -> str:
        executable = shutil.which("tesseract")
        if executable is None:
            raise PrivacyScanError("image OCR is unavailable, so the screenshot cannot be approved")
        with tempfile.TemporaryDirectory(prefix="reporticles-ocr-") as temp_dir:
            output_base = Path(temp_dir) / "ocr"
            completed = subprocess.run(
                [executable, str(path), str(output_base)],
                check=False,
                capture_output=True,
                text=True,
                timeout=90,
            )
            if completed.returncode != 0:
                raise PrivacyScanError("image OCR failed")
            output_file = output_base.with_suffix(".txt")
            if not output_file.is_file():
                raise PrivacyScanError("image OCR produced no inspectable output")
            return output_file.read_text(encoding="utf-8", errors="replace")
