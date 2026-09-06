"""Typed contracts shared by the loader, solver and runtime."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Annotated, Any, Literal, TypeAlias

from pydantic import BaseModel, ConfigDict, Field, model_validator

JsonScalar: TypeAlias = str | int | float | bool | None
MAX_PLAN_OPERATIONS = 128


@dataclass(frozen=True, slots=True)
class QualifiedRange:
    sheet: str
    cells: str


@dataclass(frozen=True, slots=True)
class TaskSpec:
    id: str
    instruction_type: str
    instruction: str
    spreadsheet_path: str
    init_xlsx: Path
    answer_ranges: tuple[QualifiedRange, ...]
    data_position: str | None = None

    @property
    def is_cell_level(self) -> bool:
        return self.instruction_type.casefold().startswith("cell")


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)


class SetValue(StrictModel):
    op: Literal["set_value"]
    sheet: str
    cell: str
    value: JsonScalar


class SetFormula(StrictModel):
    op: Literal["set_formula"]
    sheet: str
    cell: str
    formula: str


class FillFormula(StrictModel):
    op: Literal["fill_formula"]
    sheet: str
    range: str
    formula: str = Field(description="Formula for the top-left cell of range")


class SetArrayFormula(StrictModel):
    op: Literal["set_array_formula"]
    sheet: str
    cell: str
    formula: str = Field(description="Single-cell array formula")


class FillArrayFormula(StrictModel):
    op: Literal["fill_array_formula"]
    sheet: str
    range: str
    formula: str = Field(
        description="One array formula anchored at the top-left with this full ref"
    )


class ClearRange(StrictModel):
    op: Literal["clear_range"]
    sheet: str
    range: str


class CopyRange(StrictModel):
    op: Literal["copy_range"]
    source_sheet: str
    source_range: str
    destination_sheet: str
    destination_cell: str
    include_style: bool = True


Operation: TypeAlias = Annotated[
    SetValue
    | SetFormula
    | FillFormula
    | SetArrayFormula
    | FillArrayFormula
    | ClearRange
    | CopyRange,
    Field(discriminator="op"),
]


class SolvePlan(StrictModel):
    route: Literal["operations", "python"]
    summary: str = Field(min_length=1, max_length=500)
    operations: list[Operation] = Field(
        default_factory=list,
        max_length=MAX_PLAN_OPERATIONS,
    )
    python_code: str | None = None

    @model_validator(mode="after")
    def validate_route_payload(self) -> SolvePlan:
        if self.route == "operations":
            if not self.operations:
                raise ValueError("operations route requires at least one operation")
            if self.python_code is not None:
                raise ValueError("operations route cannot contain python_code")
        else:
            if self.operations:
                raise ValueError("python route cannot contain operations")
            if not self.python_code or not self.python_code.strip():
                raise ValueError("python route requires python_code")
        return self


def _provider_strict_solve_plan_json_schema() -> dict[str, Any]:
    """Return the shared provider-strict shape for a solve plan.

    Several structured-output backends accept optional values only when the key is
    required and its type includes ``null``. Pydantic defaults are useful to local
    callers, so the stricter representation is derived solely for model requests.
    """

    schema = SolvePlan.model_json_schema()

    def visit(node: Any) -> None:
        if isinstance(node, dict):
            properties = node.get("properties")
            if isinstance(properties, dict):
                node["required"] = list(properties)
            node.pop("default", None)
            for value in node.values():
                visit(value)
        elif isinstance(node, list):
            for value in node:
                visit(value)

    visit(schema)
    return schema


def cell_solve_plan_json_schema() -> dict[str, Any]:
    """Return the operations-only schema advertised to cell-level tasks.

    ``python_code`` is omitted entirely rather than exposed as an unusable nullable
    field. The production parser remains :class:`SolvePlan`; this schema deliberately
    describes a strict subset of plans that parser accepts.
    """

    schema = _provider_strict_solve_plan_json_schema()
    properties = schema["properties"]
    properties["route"] = {"const": "operations", "type": "string"}
    properties["operations"]["minItems"] = 1
    properties.pop("python_code")
    schema["required"] = ["route", "summary", "operations"]
    return schema


def sheet_solve_plan_json_schema() -> dict[str, Any]:
    """Return the mutually exclusive operations-or-Python schema for sheet tasks.

    The top-level properties retain their complete types and resource bounds. The
    two ``oneOf`` branches then encode the same route invariants enforced by
    :class:`SolvePlan`: a non-empty operation list with null Python, or an empty
    operation list with non-blank Python source.
    """

    schema = _provider_strict_solve_plan_json_schema()
    schema["oneOf"] = [
        {
            "properties": {
                "route": {"const": "operations"},
                "operations": {"minItems": 1},
                "python_code": {"type": "null"},
            },
            "required": ["route", "operations", "python_code"],
        },
        {
            "properties": {
                "route": {"const": "python"},
                "operations": {"maxItems": 0},
                "python_code": {"type": "string", "pattern": r"\S"},
            },
            "required": ["route", "operations", "python_code"],
        },
    ]
    return schema


def solve_plan_json_schema() -> dict[str, Any]:
    """Return the full provider-strict schema accepted for sheet-level tasks.

    This compatibility entry point now includes the route invariants instead of
    advertising combinations that the production :class:`SolvePlan` rejects.
    """

    return sheet_solve_plan_json_schema()


@dataclass(frozen=True, slots=True)
class ContextPack:
    text: str
    original_chars: int
    truncated: bool
    sha256: str


@dataclass(frozen=True, slots=True)
class ModelReply:
    text: str
    input_tokens: int | None
    output_tokens: int | None
    latency_ms: int


@dataclass(frozen=True, slots=True)
class SolveResult:
    status: str
    plan: SolvePlan | None
    error: str | None = None
