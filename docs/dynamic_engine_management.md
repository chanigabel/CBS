# Dynamic Engine Management Guide

This guide defines the rules, configuration patterns, and operational guidelines for a dynamic engine management system. The system is designed to run standard engines such as `NameEngine`, `GenderEngine`, and `DateEngine`, while allowing future engines such as `LocationEngine` or customer-specific engines to be added without changing core orchestration logic.

## 1. Engine Structure in the System

Each engine must be implemented as its own class and must perform one focused action or validation domain.

Examples:

- `NameEngine`: standardizes first names, last names, and related name fields.
- `GenderEngine`: normalizes gender values into the expected internal representation.
- `DateEngine`: validates and standardizes date fields.
- `LocationEngine`: standardizes city, country, address, or geographic fields.
- `CustomEngine`: handles customer-specific business rules.

All engines should follow a shared interface so the system can run them uniformly.

```python
from abc import ABC, abstractmethod
from typing import Any, Dict


class BaseEngine(ABC):
    engine_key: str
    display_name: str

    @abstractmethod
    def run(self, payload: Dict[str, Any], context: Dict[str, Any]) -> Dict[str, Any]:
        """Process payload data and return the updated payload."""
        raise NotImplementedError
```

Example engine:

```python
class NameEngine(BaseEngine):
    engine_key = "name"
    display_name = "Name Engine"

    def run(self, payload, context):
        payload["first_name"] = normalize_name(payload.get("first_name"))
        payload["last_name"] = normalize_name(payload.get("last_name"))
        return payload
```

Engines should be registered in a dictionary, registry, dependency container, plugin loader, or similar structure.

```python
ENGINE_REGISTRY = {
    "name": NameEngine,
    "gender": GenderEngine,
    "date": DateEngine,
}
```

Recommended engine rules:

- Each engine has one stable `engine_key`.
- Each engine exposes metadata such as display name, version, description, and supported fields.
- Engines must not directly enable or disable other engines.
- Engines should avoid hidden global state.
- Engines should be deterministic where possible.
- Engines should receive shared dependencies through `context`, not by importing orchestration internals.

## 2. Flags for Engine Activation

Each engine must have an associated boolean flag that determines whether it is enabled.

Example configuration:

```json
{
  "engines": {
    "name": {
      "enabled": true,
      "class": "engines.name.NameEngine",
      "run_mode": "sequential",
      "priority": 10
    },
    "gender": {
      "enabled": true,
      "class": "engines.gender.GenderEngine",
      "run_mode": "sequential",
      "priority": 20
    },
    "date": {
      "enabled": false,
      "class": "engines.date.DateEngine",
      "run_mode": "sequential",
      "priority": 30
    }
  }
}
```

Flags may be configured through:

- JSON, YAML, TOML, or another versioned configuration file.
- A database table.
- An admin dashboard.
- A REST or GraphQL API.
- Environment-specific configuration loaded during deployment.

Recommended flag rules:

- `enabled: true` means the engine may run.
- `enabled: false` means the engine must not run.
- Missing engine flags should default to disabled unless the system has a clear compatibility requirement.
- Configuration changes should be validated before being applied.
- Production systems should audit every flag change.

## 3. Adding and Removing Engines Dynamically

The system should allow new engines to be added without modifying the central runner. This is done by registering the engine and updating configuration.

Example: adding `LocationEngine`.

Engine implementation:

```python
class LocationEngine(BaseEngine):
    engine_key = "location"
    display_name = "Location Engine"

    def run(self, payload, context):
        payload["city"] = normalize_city(payload.get("city"))
        payload["country"] = normalize_country(payload.get("country"))
        return payload
```

Registry update:

```python
ENGINE_REGISTRY["location"] = LocationEngine
```

Configuration update:

```json
{
  "engines": {
    "location": {
      "enabled": true,
      "class": "engines.location.LocationEngine",
      "run_mode": "sequential",
      "priority": 40,
      "settings": {
        "default_country": "Israel",
        "strict_validation": false
      }
    }
  }
}
```

For plugin-based systems, registration can be loaded from plugin metadata rather than editing a static dictionary.

```json
{
  "plugin": "customer_location_plugin",
  "engine_key": "location",
  "class": "customer_plugins.location.LocationEngine",
  "enabled": true
}
```

Removing an engine should usually mean disabling it first. Full removal should happen only after confirming no workflows, saved configurations, or audits still depend on it.

Recommended removal lifecycle:

1. Set `enabled` to `false`.
2. Mark the engine as deprecated if it should no longer be used.
3. Remove it from configuration after dependent workflows are migrated.
4. Remove the engine code or plugin only after no active configuration references it.

## 4. API for Managing Engines

The system should expose an administrative API or interface for managing engines. This may be implemented as a REST API, GraphQL API, CLI, or admin dashboard.

Required administrator actions:

- View all registered engines.
- View only active engines.
- Add or register a new engine.
- Enable or disable an engine.
- Update engine settings.
- Remove or deprecate an engine.
- View recent execution and configuration audit history.

Example REST endpoints:

| Method | Endpoint | Purpose |
| --- | --- | --- |
| `GET` | `/api/engines` | List all registered engines |
| `GET` | `/api/engines?enabled=true` | List enabled engines |
| `GET` | `/api/engines/{engine_key}` | Get one engine configuration |
| `POST` | `/api/engines` | Register a new engine |
| `PATCH` | `/api/engines/{engine_key}` | Update engine metadata or settings |
| `POST` | `/api/engines/{engine_key}/enable` | Enable an engine |
| `POST` | `/api/engines/{engine_key}/disable` | Disable an engine |
| `DELETE` | `/api/engines/{engine_key}` | Remove or deprecate an engine |
| `POST` | `/api/engines/reload` | Reload engine configuration at runtime |

Example request to register an engine:

```http
POST /api/engines
Content-Type: application/json
Authorization: Bearer <admin-token>

{
  "engine_key": "location",
  "display_name": "Location Engine",
  "class": "engines.location.LocationEngine",
  "enabled": true,
  "run_mode": "sequential",
  "priority": 40,
  "settings": {
    "default_country": "Israel",
    "strict_validation": false
  }
}
```

Example request to disable an engine:

```http
POST /api/engines/date/disable
Authorization: Bearer <admin-token>
```

Example response:

```json
{
  "engine_key": "date",
  "enabled": false,
  "updated_by": "admin@example.com",
  "updated_at": "2026-05-14T09:30:00Z"
}
```

API best practices:

- Require authentication for all write operations.
- Validate engine keys and class paths.
- Reject duplicate engine keys.
- Keep an audit log of every configuration change.
- Support dry-run validation before applying a new engine configuration.
- Return clear errors for unknown engines, invalid settings, and dependency conflicts.

## 5. How Engines Run

The engine runner should load configuration, filter enabled engines, order them, and execute them according to their configured mode.

Basic sequential execution:

```python
def run_enabled_engines(payload, config, context):
    engines = config["engines"]

    enabled_engines = [
        (key, engine_config)
        for key, engine_config in engines.items()
        if engine_config.get("enabled") is True
    ]

    enabled_engines.sort(key=lambda item: item[1].get("priority", 1000))

    for engine_key, engine_config in enabled_engines:
        engine_class = resolve_engine_class(engine_config["class"])
        engine = engine_class()
        payload = engine.run(payload, context)

    return payload
```

Sequential mode should be used when:

- Engines depend on previous engine output.
- Output order matters.
- The same fields may be modified by multiple engines.
- Error handling must stop the pipeline immediately.

Parallel mode may be used when:

- Engines are independent.
- Engines do not write to the same fields.
- The system has merge rules for combining results.
- Performance is more important than strict ordering.

Parallel execution requires a merge strategy:

```json
{
  "parallel_merge_strategy": "fail_on_conflict"
}
```

Recommended merge strategies:

- `fail_on_conflict`: reject conflicting updates to the same field.
- `last_write_wins`: apply priority order after parallel execution.
- `field_ownership`: each engine may only write fields it owns.
- `manual_review`: conflicting outputs are sent to review.

## 6. Updating the Engine Configuration

Engine configuration should be stored centrally and should be easy to update.

Supported storage options:

- Local JSON/YAML file for simple deployments.
- Database table for multi-user systems.
- Configuration service for distributed deployments.
- Admin dashboard backed by database storage.

Example database shape:

| Column | Description |
| --- | --- |
| `engine_key` | Stable unique key, such as `name` |
| `display_name` | Human-readable engine name |
| `class_path` | Import path for the engine class |
| `enabled` | Boolean activation flag |
| `priority` | Execution order |
| `run_mode` | `sequential` or `parallel` |
| `settings_json` | Engine-specific settings |
| `version` | Engine version |
| `updated_by` | User who changed the config |
| `updated_at` | Last update timestamp |

Runtime reload options:

- Manual reload endpoint: `POST /api/engines/reload`.
- File watcher for local config files.
- Database polling with version numbers.
- Event-based reload through a message bus.

Recommended reload rules:

- Validate new configuration before swapping it into runtime.
- Apply configuration atomically.
- Keep the previous working configuration available for rollback.
- Do not interrupt an already-running pipeline unless explicitly supported.
- Log the configuration version used for each execution.

## 7. System Workflow Example

Example end-to-end workflow:

1. System starts.
2. Configuration is loaded from the central source.
3. Engine definitions are validated.
4. Engine classes are resolved and registered.
5. Request or batch job starts.
6. Runner checks each engine flag.
7. Enabled engines are ordered by priority.
8. Engines execute sequentially or in parallel.
9. Results are merged and returned.
10. Execution logs and audit records are stored.

Example configuration:

```json
{
  "config_version": "2026.05.14-001",
  "engines": {
    "name": {
      "enabled": true,
      "class": "engines.name.NameEngine",
      "priority": 10,
      "run_mode": "sequential"
    },
    "gender": {
      "enabled": true,
      "class": "engines.gender.GenderEngine",
      "priority": 20,
      "run_mode": "sequential"
    },
    "date": {
      "enabled": false,
      "class": "engines.date.DateEngine",
      "priority": 30,
      "run_mode": "sequential"
    },
    "location": {
      "enabled": true,
      "class": "engines.location.LocationEngine",
      "priority": 40,
      "run_mode": "sequential"
    }
  }
}
```

Example runner output:

```json
{
  "config_version": "2026.05.14-001",
  "executed_engines": ["name", "gender", "location"],
  "skipped_engines": ["date"],
  "status": "completed"
}
```

## 8. Error Handling and Logging

The system must log engine activation, execution status, warnings, and failures.

Log each execution with:

- Execution ID.
- Configuration version.
- Engine key.
- Engine version.
- Enabled or skipped status.
- Start time and end time.
- Runtime duration.
- Input record count, if applicable.
- Output record count, if applicable.
- Error message and stack trace for failures.

Example log event:

```json
{
  "event": "engine_execution_completed",
  "execution_id": "exec-20260514-001",
  "config_version": "2026.05.14-001",
  "engine_key": "name",
  "status": "success",
  "duration_ms": 42
}
```

Failure handling should be configurable per engine.

```json
{
  "engines": {
    "location": {
      "enabled": true,
      "class": "engines.location.LocationEngine",
      "on_error": "continue"
    }
  }
}
```

Recommended `on_error` options:

- `stop`: stop the pipeline and return an error.
- `continue`: log the failure and continue with the next engine.
- `skip_record`: skip the current record but continue the batch.
- `manual_review`: mark the result for review.

Error handling best practices:

- Do not hide engine failures.
- Include enough context to diagnose the issue.
- Avoid logging sensitive personal data unless explicitly allowed.
- Use structured logs where possible.
- Record both technical errors and business validation failures.

## 9. Security and Access Control

Engine management can change production behavior, so write access must be restricted.

Recommended roles:

| Role | Permissions |
| --- | --- |
| Viewer | View engines and execution status |
| Operator | Reload configuration and view logs |
| Engine Admin | Enable, disable, and update engines |
| System Admin | Register, remove, or approve third-party engines |

Security requirements:

- Require authentication for admin APIs.
- Require authorization checks for every write operation.
- Audit all configuration changes.
- Store who changed what, when, and from where.
- Use approval workflows for high-risk production changes.
- Validate class paths and plugin sources before loading engines.
- Restrict third-party engines to approved packages or signed plugins.
- Protect configuration files and database records from unauthorized edits.

For user-defined or third-party engines:

- Run code in a restricted environment where possible.
- Limit filesystem, network, and secret access.
- Review code before installation.
- Pin package versions.
- Scan dependencies for vulnerabilities.
- Define data privacy requirements before passing sensitive records to the engine.

## 10. Future Enhancements

The system should be designed so new engines can be added with minimal refactoring.

Recommended extension points:

- Engine registry.
- Plugin discovery.
- Engine metadata schema.
- Engine-specific settings.
- Configuration validation.
- Execution hooks.
- Audit logging.
- Runtime reload.

Potential future features:

- Admin dashboard for engine management.
- Engine marketplace or plugin catalog.
- Per-tenant engine configuration.
- Per-workflow engine configuration.
- Engine dependency graph.
- Engine health checks.
- Dry-run mode before enabling an engine.
- Canary enablement for a subset of traffic.
- Versioned engine rollbacks.
- Field ownership rules to prevent conflicting writes.
- Metrics dashboard for execution time, failure rate, and usage.

Guidelines for third-party or user-defined engines:

- Must implement the shared `BaseEngine` interface.
- Must declare `engine_key`, display name, version, and supported fields.
- Must include configuration schema for custom settings.
- Must include tests for expected inputs and failure cases.
- Must define whether it supports sequential execution, parallel execution, or both.
- Must document whether it reads or writes sensitive fields.
- Must be approved before production use.

## Recommended Configuration Contract

Use a consistent schema for every engine.

```json
{
  "engine_key": "name",
  "display_name": "Name Engine",
  "description": "Standardizes first name and last name fields.",
  "class": "engines.name.NameEngine",
  "enabled": true,
  "priority": 10,
  "run_mode": "sequential",
  "version": "1.0.0",
  "on_error": "stop",
  "settings": {},
  "permissions": {
    "can_view": ["viewer", "operator", "engine_admin", "system_admin"],
    "can_modify": ["engine_admin", "system_admin"]
  }
}
```

Required fields:

- `engine_key`
- `class`
- `enabled`
- `priority`
- `run_mode`

Optional but recommended fields:

- `display_name`
- `description`
- `version`
- `on_error`
- `settings`
- `permissions`
- `dependencies`
- `field_ownership`

## Best Practices Summary

- Keep engine orchestration separate from engine implementation.
- Use stable engine keys and a shared engine interface.
- Store flags and settings in a central configuration source.
- Validate configuration before applying it.
- Default unknown or invalid engines to disabled.
- Log every execution and configuration change.
- Use roles and permissions for all administrative actions.
- Prefer disabling engines before deleting them.
- Support runtime reload with atomic configuration swaps.
- Design third-party engine support around explicit contracts, approval, and isolation.
