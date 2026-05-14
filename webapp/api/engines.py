"""Administrative API for dynamic engine management."""

from typing import Any, Dict, Optional

from fastapi import APIRouter, Depends, Header, HTTPException, Query
from pydantic import BaseModel, Field

from src.excel_standardization.engine_management import (
    EngineAccessError,
    EngineConfigurationError,
    EngineManager,
    PASSTHROUGH_ENGINE_CLASS,
    ROLE_VIEWER,
)
from webapp.dependencies import get_column_mapping_schema_service, get_engine_manager
from webapp.models.requests import ColumnSchemaMappingRequest
from webapp.services.column_mapping_schema import ColumnMappingSchemaService


router = APIRouter(tags=["engines"])


class EngineRegistrationRequest(BaseModel):
    engine_key: str
    class_path: str = Field(
        default=PASSTHROUGH_ENGINE_CLASS,
        alias="class",
    )
    enabled: bool = False
    priority: int = 1000
    run_mode: str = "sequential"
    display_name: str = ""
    description: str = ""
    version: str = "1.0.0"
    on_error: str = "stop"
    settings: Dict[str, Any] = Field(default_factory=dict)
    permissions: Dict[str, Any] = Field(default_factory=dict)

    model_config = {"populate_by_name": True}

    def to_engine_payload(self) -> Dict[str, Any]:
        data = self.model_dump(by_alias=True)
        if not data.get("display_name"):
            data["display_name"] = f"{self.engine_key} Engine"
        if not data.get("description"):
            data["description"] = "Dynamically registered passthrough engine."
        return data


class EngineUpdateRequest(BaseModel):
    enabled: Optional[bool] = None
    priority: Optional[int] = None
    run_mode: Optional[str] = None
    display_name: Optional[str] = None
    description: Optional[str] = None
    version: Optional[str] = None
    on_error: Optional[str] = None
    settings: Optional[Dict[str, Any]] = None
    permissions: Optional[Dict[str, Any]] = None
    deprecated: Optional[bool] = None

    def updates(self) -> Dict[str, Any]:
        return {key: value for key, value in self.model_dump().items() if value is not None}


def _role(x_engine_role: Optional[str] = Header(default=ROLE_VIEWER)) -> str:
    return x_engine_role or ROLE_VIEWER


def _user(x_engine_user: Optional[str] = Header(default="api-user")) -> str:
    return x_engine_user or "api-user"


def _handle_error(exc: Exception) -> None:
    if isinstance(exc, EngineAccessError):
        raise HTTPException(status_code=403, detail=str(exc))
    if isinstance(exc, KeyError):
        raise HTTPException(status_code=404, detail=f"Engine '{exc.args[0]}' not found")
    if isinstance(exc, EngineConfigurationError):
        raise HTTPException(status_code=400, detail=str(exc))
    raise exc


@router.get("/engines")
def list_engines(
    enabled: Optional[bool] = Query(default=None),
    role: str = Depends(_role),
    manager: EngineManager = Depends(get_engine_manager),
) -> Dict[str, Any]:
    try:
        return {
            "config_version": manager.config_version,
            "engines": manager.list_engines(enabled=enabled, role=role),
        }
    except Exception as exc:
        _handle_error(exc)


@router.get("/engines/mappings")
def list_column_schema_mappings(
    role: str = Depends(_role),
    schema_service: ColumnMappingSchemaService = Depends(get_column_mapping_schema_service),
) -> Dict[str, Any]:
    try:
        if role not in {"viewer", "operator", "engine_admin", "system_admin"}:
            raise EngineAccessError(f"Role '{role}' is not allowed to perform this operation")
        return {
            "fields": schema_service.fields(),
            "mappings": schema_service.mappings(),
            "suggestions": schema_service.suggestions(),
        }
    except Exception as exc:
        _handle_error(exc)


@router.get("/engines/{engine_key}")
def get_engine(
    engine_key: str,
    role: str = Depends(_role),
    manager: EngineManager = Depends(get_engine_manager),
) -> Dict[str, Any]:
    try:
        return manager.get_engine(engine_key, role=role)
    except Exception as exc:
        _handle_error(exc)


@router.post("/engines")
def register_engine(
    request: EngineRegistrationRequest,
    role: str = Depends(_role),
    user: str = Depends(_user),
    manager: EngineManager = Depends(get_engine_manager),
) -> Dict[str, Any]:
    try:
        return manager.add_engine(request.to_engine_payload(), role=role, user=user)
    except Exception as exc:
        _handle_error(exc)


@router.patch("/engines/{engine_key}")
def update_engine(
    engine_key: str,
    request: EngineUpdateRequest,
    role: str = Depends(_role),
    user: str = Depends(_user),
    manager: EngineManager = Depends(get_engine_manager),
) -> Dict[str, Any]:
    try:
        return manager.update_engine(engine_key, request.updates(), role=role, user=user)
    except Exception as exc:
        _handle_error(exc)


@router.post("/engines/{engine_key}/enable")
def enable_engine(
    engine_key: str,
    role: str = Depends(_role),
    user: str = Depends(_user),
    manager: EngineManager = Depends(get_engine_manager),
) -> Dict[str, Any]:
    try:
        return manager.enable(engine_key, role=role, user=user)
    except Exception as exc:
        _handle_error(exc)


@router.post("/engines/{engine_key}/disable")
def disable_engine(
    engine_key: str,
    role: str = Depends(_role),
    user: str = Depends(_user),
    manager: EngineManager = Depends(get_engine_manager),
) -> Dict[str, Any]:
    try:
        return manager.disable(engine_key, role=role, user=user)
    except Exception as exc:
        _handle_error(exc)


@router.delete("/engines/remove-mapping")
def delete_column_schema_mapping(
    request: ColumnSchemaMappingRequest,
    role: str = Depends(_role),
    user: str = Depends(_user),
    schema_service: ColumnMappingSchemaService = Depends(get_column_mapping_schema_service),
) -> Dict[str, Any]:
    try:
        if role not in {"engine_admin", "system_admin"}:
            raise EngineAccessError(f"Role '{role}' is not allowed to perform this operation")
        mappings = schema_service.remove_mapping(request.standard_name, request.synonym)
        return {"updated_by": user, "mappings": mappings}
    except Exception as exc:
        _handle_error(exc)


@router.delete("/engines/{engine_key}")
def remove_engine(
    engine_key: str,
    role: str = Depends(_role),
    user: str = Depends(_user),
    manager: EngineManager = Depends(get_engine_manager),
) -> Dict[str, Any]:
    try:
        return manager.remove_engine(engine_key, role=role, user=user)
    except Exception as exc:
        _handle_error(exc)


@router.post("/engines/reload")
def reload_engines(
    role: str = Depends(_role),
    user: str = Depends(_user),
    manager: EngineManager = Depends(get_engine_manager),
) -> Dict[str, Any]:
    try:
        return manager.reload(role=role, user=user)
    except Exception as exc:
        _handle_error(exc)


@router.get("/engine-audit/history")
def engine_audit_history(
    role: str = Depends(_role),
    manager: EngineManager = Depends(get_engine_manager),
) -> Dict[str, Any]:
    try:
        manager.authorize(role, {"operator", "engine_admin", "system_admin"})
        return {"audit": manager.audit_log, "executions": manager.execution_log}
    except Exception as exc:
        _handle_error(exc)


@router.post("/engines/add-mapping")
def add_column_schema_mapping(
    request: ColumnSchemaMappingRequest,
    role: str = Depends(_role),
    user: str = Depends(_user),
    schema_service: ColumnMappingSchemaService = Depends(get_column_mapping_schema_service),
) -> Dict[str, Any]:
    try:
        if role not in {"engine_admin", "system_admin"}:
            raise EngineAccessError(f"Role '{role}' is not allowed to perform this operation")
        mappings = schema_service.add_mapping(request.standard_name, request.synonym)
        return {"updated_by": user, "mappings": mappings}
    except Exception as exc:
        _handle_error(exc)


@router.post("/engines/remove-mapping")
def remove_column_schema_mapping(
    request: ColumnSchemaMappingRequest,
    role: str = Depends(_role),
    user: str = Depends(_user),
    schema_service: ColumnMappingSchemaService = Depends(get_column_mapping_schema_service),
) -> Dict[str, Any]:
    try:
        if role not in {"engine_admin", "system_admin"}:
            raise EngineAccessError(f"Role '{role}' is not allowed to perform this operation")
        mappings = schema_service.remove_mapping(request.standard_name, request.synonym)
        return {"updated_by": user, "mappings": mappings}
    except Exception as exc:
        _handle_error(exc)


@router.post("/engines/reload-mapping")
def reload_column_schema_mapping(
    role: str = Depends(_role),
    user: str = Depends(_user),
    schema_service: ColumnMappingSchemaService = Depends(get_column_mapping_schema_service),
) -> Dict[str, Any]:
    try:
        if role not in {"operator", "engine_admin", "system_admin"}:
            raise EngineAccessError(f"Role '{role}' is not allowed to perform this operation")
        schema_service.reload()
        return {
            "updated_by": user,
            "fields": schema_service.fields(),
            "mappings": schema_service.mappings(),
            "suggestions": schema_service.suggestions(),
        }
    except Exception as exc:
        _handle_error(exc)
