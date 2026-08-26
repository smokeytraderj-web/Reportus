"""Local-only preparation of portal captures before privacy review or synthesis."""

from __future__ import annotations

import tempfile
from dataclasses import dataclass
from pathlib import Path

from PIL import Image, ImageOps, UnidentifiedImageError


_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}


class PortalCaptureError(RuntimeError):
    """A portal screenshot could not be reduced to a safe report input."""


@dataclass(frozen=True, slots=True)
class PreparedPortalSelections:
    selections: dict[str, tuple[Path, ...]]
    temporary_directory: tempfile.TemporaryDirectory[str] | None = None

    def cleanup(self) -> None:
        if self.temporary_directory is not None:
            self.temporary_directory.cleanup()


def prepare_client_deck_portal_captures(
    selections: dict[str, tuple[Path, ...]],
) -> PreparedPortalSelections:
    """Keep only the Riskalyze analytics panel from risk-snapshot images.

    Riskalyze's current-portfolio page places household and account-level details
    in the left two-thirds of a widescreen capture. The report needs only the
    portfolio analytics panel on the right. Cropping happens locally before the
    privacy scanner, staging, OCR, or any provider request can see the image.
    """

    risk_paths = selections.get("risk_snapshot", ())
    images = [path for path in risk_paths if path.suffix.lower() in _IMAGE_EXTENSIONS]
    if not images:
        return PreparedPortalSelections(dict(selections))

    temporary = tempfile.TemporaryDirectory(prefix="reportus-riskalyze-safe-")
    root = Path(temporary.name)
    replacements: dict[Path, Path] = {}
    try:
        for index, source in enumerate(images, start=1):
            try:
                with Image.open(source) as raw:
                    image = ImageOps.exif_transpose(raw).convert("RGB")
            except (OSError, UnidentifiedImageError) as exc:
                raise PortalCaptureError(
                    "The Riskalyze screenshot could not be opened. Capture it again as PNG or JPG."
                ) from exc
            if image.width < 900 or image.height < 500 or image.width / image.height < 1.45:
                raise PortalCaptureError(
                    "Use a full-width Riskalyze Current Portfolio screenshot (at least 900×500)."
                )
            safe_panel = image.crop((round(image.width * .66), 0, image.width, image.height))
            safe_panel = safe_panel.resize(
                (safe_panel.width * 2, safe_panel.height * 2), Image.Resampling.LANCZOS
            )
            destination = root / f"riskalyze_analytics_{index}.png"
            safe_panel.save(destination, format="PNG", optimize=True)
            replacements[source] = destination
    except Exception:
        temporary.cleanup()
        raise

    prepared = dict(selections)
    prepared["risk_snapshot"] = tuple(replacements.get(path, path) for path in risk_paths)
    return PreparedPortalSelections(prepared, temporary)
