from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class CheckPolicy:
    level: str = "structural"
    export_pdf: bool = False

    def __post_init__(self) -> None:
        if self.level not in {"structural", "layout", "visual"}:
            raise ValueError(f"Unknown check level: {self.level}")

    @property
    def render(self) -> bool:
        return self.export_pdf or self.level != "structural"

    def flags(self) -> list[str]:
        flags = ["--check", self.level]
        if self.export_pdf:
            flags.append("--export-pdf")
        return flags

    def summary(self) -> dict[str, object]:
        return {
            "level": self.level,
            "exportPdf": self.export_pdf,
            "layoutRequested": self.level in {"layout", "visual"},
            "previewsRequested": self.level == "visual",
            "visualReviewed": False,
        }
