"""Output path checks shared by PDF and JSON writers."""

from __future__ import annotations

from pathlib import Path


class OutputPathGuard:
    """Stateless checks that avoid writes through symbolic-link directories."""

    @staticmethod
    def symlinked_parent(path: Path) -> Path | None:
        parent = path.expanduser().absolute().parent
        for candidate in (parent, *parent.parents):
            if candidate.is_symlink():
                return candidate
        return None

    @staticmethod
    def comparison_key(path: Path) -> Path:
        """Return a normalized key for detecting colliding output names."""

        return path.expanduser().resolve(strict=False)
