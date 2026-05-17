"""Safe cleanup for application-managed runtime files."""

from __future__ import annotations

import logging
import shutil
from pathlib import Path
from typing import Iterable

logger = logging.getLogger(__name__)


class CleanupService:
    """Delete only files stored inside an explicit runtime-directory allowlist."""

    def __init__(self, allowed_runtime_dirs: Iterable[Path]) -> None:
        self.allowed_runtime_dirs = tuple(Path(path) for path in allowed_runtime_dirs)

    def cleanup_runtime_files(self, *, reason: str) -> None:
        """Remove contents from all allowlisted runtime directories.

        Missing directories are skipped. The directories themselves are kept so
        services can continue to write into them after startup cleanup.
        """
        for directory in self.allowed_runtime_dirs:
            try:
                self._cleanup_directory(directory, reason=reason)
            except Exception:
                logger.exception(
                    "runtime_cleanup_failed",
                    extra={
                        "event": "runtime_cleanup_failed",
                        "cleanup_reason": reason,
                        "runtime_dir": str(directory),
                    },
                )

    def _cleanup_directory(self, directory: Path, *, reason: str) -> None:
        allowed_root = directory.resolve(strict=False)

        if not directory.exists():
            logger.info(
                "runtime_cleanup_skipped_missing_dir",
                extra={
                    "event": "runtime_cleanup_skipped_missing_dir",
                    "cleanup_reason": reason,
                    "runtime_dir": str(allowed_root),
                },
            )
            return

        if not directory.is_dir():
            logger.warning(
                "runtime_cleanup_skipped_not_dir",
                extra={
                    "event": "runtime_cleanup_skipped_not_dir",
                    "cleanup_reason": reason,
                    "runtime_dir": str(allowed_root),
                },
            )
            return

        if self._is_protected_directory(allowed_root):
            logger.error(
                "runtime_cleanup_refused_protected_dir",
                extra={
                    "event": "runtime_cleanup_refused_protected_dir",
                    "cleanup_reason": reason,
                    "runtime_dir": str(allowed_root),
                },
            )
            return

        deleted = 0
        failed = 0
        for child in directory.iterdir():
            try:
                if self._delete_child(child, allowed_root):
                    deleted += 1
            except Exception:
                failed += 1
                logger.exception(
                    "runtime_cleanup_delete_failed",
                    extra={
                        "event": "runtime_cleanup_delete_failed",
                        "cleanup_reason": reason,
                        "runtime_dir": str(allowed_root),
                        "target": str(child),
                    },
                )

        logger.info(
            "runtime_cleanup_complete",
            extra={
                "event": "runtime_cleanup_complete",
                "cleanup_reason": reason,
                "runtime_dir": str(allowed_root),
                "deleted_entries": deleted,
                "failed_entries": failed,
            },
        )

    def _delete_child(self, child: Path, allowed_root: Path) -> bool:
        resolved_child = child.resolve(strict=False)
        if not self._is_inside(resolved_child, allowed_root):
            logger.error(
                "runtime_cleanup_refused_outside_allowlist",
                extra={
                    "event": "runtime_cleanup_refused_outside_allowlist",
                    "runtime_dir": str(allowed_root),
                    "target": str(resolved_child),
                },
            )
            return False

        if child.is_symlink() or child.is_file():
            child.unlink()
            return True

        if child.is_dir():
            shutil.rmtree(child)
            return True

        logger.warning(
            "runtime_cleanup_skipped_unknown_path_type",
            extra={
                "event": "runtime_cleanup_skipped_unknown_path_type",
                "runtime_dir": str(allowed_root),
                "target": str(resolved_child),
            },
        )
        return False

    @staticmethod
    def _is_inside(path: Path, root: Path) -> bool:
        try:
            path.relative_to(root)
            return path != root
        except ValueError:
            return False

    @staticmethod
    def _is_protected_directory(path: Path) -> bool:
        resolved = path.resolve(strict=False)
        home = Path.home().resolve(strict=False)
        protected_user_dirs = (
            home / "Downloads",
            home / "Desktop",
            home / "Documents",
        )

        if resolved.anchor and str(resolved) == resolved.anchor:
            return True

        if resolved == home:
            return True

        return any(
            resolved == folder or folder in resolved.parents
            for folder in protected_user_dirs
        )
