from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class FactCard:
    id: str
    type: str
    keywords: list[str] = field(default_factory=list)
    page: int | None = None
