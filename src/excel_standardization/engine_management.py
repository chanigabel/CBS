"""Dynamic engine registration, configuration, execution, and audit support."""

from __future__ import annotations

import importlib
import json
import logging
import traceback
import uuid
from abc import ABC, abstractmethod
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from time import perf_counter
from typing import Any, Callable, Dict, Iterable, List, Optional


logger = logging.getLogger(__name__)


ROLE_VIEWER = "viewer"
ROLE_OPERATOR = "operator"
ROLE_ENGINE_ADMIN = "engine_admin"
ROLE_SYSTEM_ADMIN = "system_admin"

VIEW_ROLES = {ROLE_VIEWER, ROLE_OPERATOR, ROLE_ENGINE_ADMIN, ROLE_SYSTEM_ADMIN}
OPERATE_ROLES = {ROLE_OPERATOR, ROLE_ENGINE_ADMIN, ROLE_SYSTEM_ADMIN}
MODIFY_ROLES = {ROLE_ENGINE_ADMIN, ROLE_SYSTEM_ADMIN}
REGISTER_ROLES = {ROLE_SYSTEM_ADMIN}

VALID_RUN_MODES = {"sequential", "parallel"}
VALID_ON_ERROR = {"stop", "continue", "skip_record", "manual_review"}
PASSTHROUGH_ENGINE_CLASS = "src.excel_standardization.engine_management.PassthroughEngine"


class EngineAccessError(PermissionError):
    """Raised when a role is not allowed to perform an engine operation."""


class EngineConfigurationError(ValueError):
    """Raised when an engine configuration is invalid."""


class BaseEngine(ABC):
    """Shared interface for dynamically executable engines."""

    engine_key: str
    display_name: str
    version: str = "1.0.0"
    description: str = ""
    supported_fields: List[str] = []

    @abstractmethod
    def run(self, payload: Dict[str, Any], context: Dict[str, Any]) -> Dict[str, Any]:
        """Process a payload and return the updated payload."""


class PassthroughEngine(BaseEngine):
    """Safe default engine used when the UI registers a key without plugin code."""

    engine_key = "passthrough"
    display_name = "Passthrough Engine"
    version = "1.0.0"
    description = "Leaves the payload unchanged."
    supported_fields: List[str] = []

    def run(self, payload: Dict[str, Any], context: Dict[str, Any]) -> Dict[str, Any]:
        return payload


@dataclass(frozen=True)
class EngineConfig:
    engine_key: str
    class_path: str
    enabled: bool
    priority: int
    run_mode: str = "sequential"
    display_name: str = ""
    description: str = ""
    version: str = "1.0.0"
    on_error: str = "stop"
    settings: Dict[str, Any] = field(default_factory=dict)
    permissions: Dict[str, List[str]] = field(default_factory=dict)
    deprecated: bool = False

    @classmethod
    def from_mapping(cls, engine_key: str, raw: Dict[str, Any]) -> "EngineConfig":
        if not isinstance(raw, dict):
            raise EngineConfigurationError(f"Engine '{engine_key}' configuration must be an object")
        class_path = raw.get("class") or raw.get("class_path")
        if not class_path:
            raise EngineConfigurationError(f"Engine '{engine_key}' is missing required class path")
        if "enabled" not in raw:
            raise EngineConfigurationError(f"Engine '{engine_key}' is missing required enabled flag")
        if "priority" not in raw:
            raise EngineConfigurationError(f"Engine '{engine_key}' is missing required priority")

        run_mode = raw.get("run_mode", "sequential")
        on_error = raw.get("on_error", "stop")
        if run_mode not in VALID_RUN_MODES:
            raise EngineConfigurationError(f"Engine '{engine_key}' has invalid run_mode '{run_mode}'")
        if on_error not in VALID_ON_ERROR:
            raise EngineConfigurationError(f"Engine '{engine_key}' has invalid on_error '{on_error}'")

        return cls(
            engine_key=str(raw.get("engine_key") or engine_key),
            class_path=str(class_path),
            enabled=bool(raw["enabled"]),
            priority=int(raw["priority"]),
            run_mode=run_mode,
            display_name=str(raw.get("display_name") or ""),
            description=str(raw.get("description") or ""),
            version=str(raw.get("version") or "1.0.0"),
            on_error=on_error,
            settings=dict(raw.get("settings") or {}),
            permissions=dict(raw.get("permissions") or {}),
            deprecated=bool(raw.get("deprecated", False)),
        )

    def to_mapping(self) -> Dict[str, Any]:
        return {
            "engine_key": self.engine_key,
            "display_name": self.display_name,
            "description": self.description,
            "class": self.class_path,
            "enabled": self.enabled,
            "priority": self.priority,
            "run_mode": self.run_mode,
            "version": self.version,
            "on_error": self.on_error,
            "settings": self.settings,
            "permissions": self.permissions,
            "deprecated": self.deprecated,
        }


class EngineRegistry:
    """Registry for engine classes and constructor factories."""

    def __init__(self) -> None:
        self._factories: Dict[str, Callable[[], Any]] = {}
        self._classes: Dict[str, type] = {}

    def register(
        self,
        engine_key: str,
        engine_class: type,
        factory: Optional[Callable[[], Any]] = None,
        *,
        replace: bool = False,
    ) -> None:
        if not replace and engine_key in self._classes:
            raise EngineConfigurationError(f"Engine '{engine_key}' is already registered")
        self._classes[engine_key] = engine_class
        self._factories[engine_key] = factory or engine_class

    def remove(self, engine_key: str) -> None:
        self._classes.pop(engine_key, None)
        self._factories.pop(engine_key, None)

    def has(self, engine_key: str) -> bool:
        return engine_key in self._classes

    def create(self, engine_key: str) -> Any:
        if engine_key not in self._factories:
            raise EngineConfigurationError(f"Engine '{engine_key}' is not registered")
        return self._factories[engine_key]()

    def keys(self) -> List[str]:
        return sorted(self._classes)

    def class_for(self, engine_key: str) -> Optional[type]:
        return self._classes.get(engine_key)


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def import_class(class_path: str) -> type:
    module_name, _, class_name = class_path.rpartition(".")
    if not module_name or not class_name:
        raise EngineConfigurationError(f"Invalid class path '{class_path}'")
    module = importlib.import_module(module_name)
    obj = getattr(module, class_name)
    if not isinstance(obj, type):
        raise EngineConfigurationError(f"Class path '{class_path}' did not resolve to a class")
    return obj


def default_engine_config() -> Dict[str, Any]:
    permissions = {
        "can_view": sorted(VIEW_ROLES),
        "can_modify": sorted(MODIFY_ROLES),
    }
    return {
        "config_version": "2026.05.14-001",
        "engines": {
            "name": {
                "engine_key": "name",
                "display_name": "Name Engine",
                "description": "Standardizes first, last, and father name fields.",
                "class": "src.excel_standardization.engines.name_engine.NameEngine",
                "enabled": True,
                "priority": 10,
                "run_mode": "sequential",
                "version": "1.0.0",
                "on_error": "continue",
                "settings": {},
                "permissions": permissions,
            },
            "gender": {
                "engine_key": "gender",
                "display_name": "Gender Engine",
                "description": "Normalizes gender values into canonical export codes.",
                "class": "src.excel_standardization.engines.gender_engine.GenderEngine",
                "enabled": True,
                "priority": 20,
                "run_mode": "sequential",
                "version": "1.0.0",
                "on_error": "continue",
                "settings": {},
                "permissions": permissions,
            },
            "date": {
                "engine_key": "date",
                "display_name": "Date Engine",
                "description": "Parses, validates, and standardizes date fields.",
                "class": "src.excel_standardization.engines.date_engine.DateEngine",
                "enabled": True,
                "priority": 30,
                "run_mode": "sequential",
                "version": "1.0.0",
                "on_error": "continue",
                "settings": {},
                "permissions": permissions,
            },
            "identifier": {
                "engine_key": "identifier",
                "display_name": "Identifier Engine",
                "description": "Validates ID values and normalizes passport fields.",
                "class": "src.excel_standardization.engines.identifier_engine.IdentifierEngine",
                "enabled": True,
                "priority": 40,
                "run_mode": "sequential",
                "version": "1.0.0",
                "on_error": "continue",
                "settings": {},
                "permissions": permissions,
            },
        },
    }


class JsonEngineConfigStore:
    """JSON-backed engine configuration with atomic in-process reloads."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self._raw_config: Dict[str, Any] = {}
        self._previous_config: Optional[Dict[str, Any]] = None
        self.ensure_exists()
        self.reload()

    def ensure_exists(self) -> None:
        if self.path.exists():
            return
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(
            json.dumps(default_engine_config(), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    def reload(self) -> Dict[str, Any]:
        raw = json.loads(self.path.read_text(encoding="utf-8"))
        self.validate(raw)
        self._previous_config = deepcopy(self._raw_config) if self._raw_config else None
        self._raw_config = raw
        return self.snapshot()

    def snapshot(self) -> Dict[str, Any]:
        return deepcopy(self._raw_config)

    def config_version(self) -> str:
        return str(self._raw_config.get("config_version") or "unversioned")

    def engine_configs(self) -> List[EngineConfig]:
        engines = self._raw_config.get("engines") or {}
        return [EngineConfig.from_mapping(key, value) for key, value in engines.items()]

    def get_engine_config(self, engine_key: str) -> EngineConfig:
        engines = self._raw_config.get("engines") or {}
        if engine_key not in engines:
            raise KeyError(engine_key)
        return EngineConfig.from_mapping(engine_key, engines[engine_key])

    def save(self, raw: Dict[str, Any]) -> None:
        self.validate(raw)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(
            json.dumps(raw, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        self._previous_config = deepcopy(self._raw_config)
        self._raw_config = deepcopy(raw)

    def update_engine(self, engine_key: str, updates: Dict[str, Any]) -> EngineConfig:
        raw = self.snapshot()
        engines = raw.setdefault("engines", {})
        if engine_key not in engines:
            raise KeyError(engine_key)
        merged = dict(engines[engine_key])
        merged.update(updates)
        merged["engine_key"] = engine_key
        engines[engine_key] = merged
        raw["config_version"] = utc_now_iso()
        self.save(raw)
        return self.get_engine_config(engine_key)

    def add_engine(self, payload: Dict[str, Any]) -> EngineConfig:
        engine_key = str(payload.get("engine_key") or "").strip()
        if not engine_key:
            raise EngineConfigurationError("engine_key is required")
        raw = self.snapshot()
        engines = raw.setdefault("engines", {})
        if engine_key in engines:
            raise EngineConfigurationError(f"Engine '{engine_key}' already exists")
        payload = dict(payload)
        payload["engine_key"] = engine_key
        payload.setdefault("class", PASSTHROUGH_ENGINE_CLASS)
        payload.setdefault("enabled", False)
        payload.setdefault("priority", 1000)
        payload.setdefault("run_mode", "sequential")
        engines[engine_key] = payload
        raw["config_version"] = utc_now_iso()
        self.save(raw)
        return self.get_engine_config(engine_key)

    def remove_engine(self, engine_key: str) -> EngineConfig:
        raw = self.snapshot()
        engines = raw.setdefault("engines", {})
        if engine_key not in engines:
            raise KeyError(engine_key)
        removed = EngineConfig.from_mapping(engine_key, engines.pop(engine_key))
        raw["config_version"] = utc_now_iso()
        self.save(raw)
        return removed

    def validate(self, raw: Dict[str, Any]) -> None:
        if not isinstance(raw, dict):
            raise EngineConfigurationError("Engine configuration must be an object")
        engines = raw.get("engines")
        if not isinstance(engines, dict):
            raise EngineConfigurationError("Engine configuration must contain an engines object")
        for engine_key, engine_raw in engines.items():
            cfg = EngineConfig.from_mapping(engine_key, engine_raw)
            if cfg.engine_key != engine_key:
                raise EngineConfigurationError(
                    f"Engine key mismatch: section '{engine_key}' has '{cfg.engine_key}'"
                )


class EngineManager:
    """Coordinates registry, JSON configuration, authorization, and audits."""

    def __init__(self, config_path: Path) -> None:
        self.registry = EngineRegistry()
        self.store = JsonEngineConfigStore(config_path)
        self.execution_log: List[Dict[str, Any]] = []
        self.audit_log: List[Dict[str, Any]] = []
        self._register_builtin_engines()
        self.resolve_configured_engines()

    def _register_builtin_engines(self) -> None:
        from .engines.date_engine import DateEngine
        from .engines.gender_engine import GenderEngine
        from .engines.identifier_engine import IdentifierEngine
        from .engines.name_engine import NameEngine
        from .engines.text_processor import TextProcessor

        self.registry.register("name", NameEngine, lambda: NameEngine(TextProcessor()), replace=True)
        self.registry.register("gender", GenderEngine, GenderEngine, replace=True)
        self.registry.register("date", DateEngine, DateEngine, replace=True)
        self.registry.register("identifier", IdentifierEngine, IdentifierEngine, replace=True)

    def authorize(self, role: str, allowed_roles: Iterable[str]) -> None:
        if role not in set(allowed_roles):
            raise EngineAccessError(f"Role '{role}' is not allowed to perform this operation")

    def resolve_configured_engines(self) -> None:
        for cfg in self.store.engine_configs():
            if self.registry.has(cfg.engine_key):
                continue
            engine_class = import_class(cfg.class_path)
            self.registry.register(cfg.engine_key, engine_class)

    def reload(self, *, role: str = ROLE_OPERATOR, user: str = "system") -> Dict[str, Any]:
        self.authorize(role, OPERATE_ROLES)
        raw = self.store.reload()
        self.registry = EngineRegistry()
        self._register_builtin_engines()
        self.resolve_configured_engines()
        self.audit("reload", user, role, None, {"config_version": self.config_version})
        return raw

    @property
    def config_version(self) -> str:
        return self.store.config_version()

    def list_engines(
        self,
        *,
        enabled: Optional[bool] = None,
        role: str = ROLE_VIEWER,
    ) -> List[Dict[str, Any]]:
        self.authorize(role, VIEW_ROLES)
        configs = self.store.engine_configs()
        if enabled is not None:
            configs = [cfg for cfg in configs if cfg.enabled is enabled]
        return [self.engine_summary(cfg) for cfg in sorted(configs, key=lambda item: item.priority)]

    def get_engine(self, engine_key: str, *, role: str = ROLE_VIEWER) -> Dict[str, Any]:
        self.authorize(role, VIEW_ROLES)
        return self.engine_summary(self.store.get_engine_config(engine_key))

    def enabled_engine_configs(self) -> List[EngineConfig]:
        self.resolve_configured_engines()
        return sorted(
            [cfg for cfg in self.store.engine_configs() if cfg.enabled and not cfg.deprecated],
            key=lambda item: item.priority,
        )

    def engine_summary(self, cfg: EngineConfig) -> Dict[str, Any]:
        engine_class = self.registry.class_for(cfg.engine_key)
        return {
            **cfg.to_mapping(),
            "registered": engine_class is not None,
            "class_metadata": self.class_metadata(engine_class) if engine_class else {},
        }

    def class_metadata(self, engine_class: Optional[type]) -> Dict[str, Any]:
        if engine_class is None:
            return {}
        return {
            "engine_key": getattr(engine_class, "engine_key", ""),
            "display_name": getattr(engine_class, "display_name", engine_class.__name__),
            "version": getattr(engine_class, "version", "1.0.0"),
            "description": getattr(engine_class, "description", ""),
            "supported_fields": list(getattr(engine_class, "supported_fields", []) or []),
        }

    def enable(self, engine_key: str, *, role: str, user: str) -> Dict[str, Any]:
        self.authorize(role, MODIFY_ROLES)
        cfg = self.store.update_engine(engine_key, {"enabled": True})
        self.audit("enable", user, role, engine_key, {"enabled": True})
        return self.engine_summary(cfg)

    def disable(self, engine_key: str, *, role: str, user: str) -> Dict[str, Any]:
        self.authorize(role, MODIFY_ROLES)
        cfg = self.store.update_engine(engine_key, {"enabled": False})
        self.audit("disable", user, role, engine_key, {"enabled": False})
        return self.engine_summary(cfg)

    def update_engine(
        self,
        engine_key: str,
        updates: Dict[str, Any],
        *,
        role: str,
        user: str,
    ) -> Dict[str, Any]:
        self.authorize(role, MODIFY_ROLES)
        if "engine_key" in updates and updates["engine_key"] != engine_key:
            raise EngineConfigurationError("engine_key cannot be changed")
        cfg = self.store.update_engine(engine_key, updates)
        self.resolve_configured_engines()
        self.audit("update", user, role, engine_key, updates)
        return self.engine_summary(cfg)

    def add_engine(self, payload: Dict[str, Any], *, role: str, user: str) -> Dict[str, Any]:
        self.authorize(role, REGISTER_ROLES)
        payload = dict(payload)
        payload.setdefault("class", PASSTHROUGH_ENGINE_CLASS)
        engine_class = import_class(str(payload.get("class") or payload.get("class_path") or ""))
        cfg = self.store.add_engine(payload)
        self.registry.register(cfg.engine_key, engine_class)
        self.audit("register", user, role, cfg.engine_key, payload)
        return self.engine_summary(cfg)

    def remove_engine(self, engine_key: str, *, role: str, user: str) -> Dict[str, Any]:
        self.authorize(role, REGISTER_ROLES)
        cfg = self.store.remove_engine(engine_key)
        self.registry.remove(engine_key)
        self.audit("remove", user, role, engine_key, {})
        return cfg.to_mapping()

    def audit(self, action: str, user: str, role: str, engine_key: Optional[str], details: Dict[str, Any]) -> None:
        event = {
            "event": "engine_configuration_changed",
            "action": action,
            "engine_key": engine_key,
            "updated_by": user,
            "role": role,
            "updated_at": utc_now_iso(),
            "details": details,
            "config_version": self.config_version,
        }
        self.audit_log.append(event)
        logger.info("engine_configuration_changed", extra=event)

    def record_execution(self, event: Dict[str, Any]) -> None:
        self.execution_log.append(event)
        if event.get("status") == "success":
            logger.info("engine_execution_completed", extra=event)
        elif event.get("status") == "skipped":
            logger.info("engine_execution_skipped", extra=event)
        else:
            logger.error("engine_execution_failed", extra=event)


class DynamicEngineRunner:
    """Runs enabled engines in configured priority order."""

    def __init__(self, manager: EngineManager) -> None:
        self.manager = manager

    def run(
        self,
        payload: Dict[str, Any],
        context: Dict[str, Any],
        built_in_handlers: Optional[Dict[str, Callable[[Dict[str, Any], Dict[str, Any]], Dict[str, Any]]]] = None,
    ) -> Dict[str, Any]:
        built_in_handlers = built_in_handlers or {}
        execution_id = context.setdefault("execution_id", f"exec-{uuid.uuid4().hex}")
        executed: List[str] = []
        skipped: List[str] = []

        for cfg in self.manager.enabled_engine_configs():
            start = perf_counter()
            started_at = utc_now_iso()
            event = {
                "event": "engine_execution_started",
                "execution_id": execution_id,
                "config_version": self.manager.config_version,
                "engine_key": cfg.engine_key,
                "engine_version": cfg.version,
                "status": "started",
                "start_time": started_at,
            }
            logger.info("engine_execution_started", extra=event)
            try:
                if cfg.engine_key in built_in_handlers:
                    payload = built_in_handlers[cfg.engine_key](payload, {"engine_config": cfg, **context})
                else:
                    engine = self.manager.registry.create(cfg.engine_key)
                    if not hasattr(engine, "run"):
                        raise EngineConfigurationError(
                            f"Engine '{cfg.engine_key}' must expose run(payload, context)"
                        )
                    payload = engine.run(payload, {"engine_config": cfg, **context})
                executed.append(cfg.engine_key)
                completed = self._execution_event(
                    execution_id, cfg, "success", started_at, start, payload, None
                )
                self.manager.record_execution(completed)
            except Exception as exc:
                failed = self._execution_event(
                    execution_id, cfg, "failed", started_at, start, payload, exc
                )
                self.manager.record_execution(failed)
                if cfg.on_error == "stop":
                    raise
                skipped.append(cfg.engine_key)

        all_configured = [cfg.engine_key for cfg in self.manager.store.engine_configs()]
        disabled = [key for key in all_configured if key not in executed and key not in skipped]
        context["engine_run_summary"] = {
            "config_version": self.manager.config_version,
            "executed_engines": executed,
            "skipped_engines": skipped + disabled,
            "status": "completed",
        }
        return payload

    def _execution_event(
        self,
        execution_id: str,
        cfg: EngineConfig,
        status: str,
        started_at: str,
        start: float,
        payload: Dict[str, Any],
        exc: Optional[Exception],
    ) -> Dict[str, Any]:
        event = {
            "event": "engine_execution_completed" if status == "success" else "engine_execution_failed",
            "execution_id": execution_id,
            "config_version": self.manager.config_version,
            "engine_key": cfg.engine_key,
            "engine_version": cfg.version,
            "status": status,
            "start_time": started_at,
            "end_time": utc_now_iso(),
            "duration_ms": round((perf_counter() - start) * 1000, 3),
            "output_record_count": 1 if isinstance(payload, dict) else None,
        }
        if exc is not None:
            event["error"] = str(exc)
            event["traceback"] = traceback.format_exc()
        return event


_default_manager: Optional[EngineManager] = None


def get_default_engine_config_path() -> Path:
    return Path.cwd() / "config" / "engine_config.json"


def get_default_engine_manager() -> EngineManager:
    global _default_manager
    if _default_manager is None:
        _default_manager = EngineManager(get_default_engine_config_path())
    return _default_manager
