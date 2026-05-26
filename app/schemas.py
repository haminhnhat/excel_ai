from __future__ import annotations

from typing import Any, Dict, List, Literal, Optional
from pydantic import BaseModel, Field


class Change(BaseModel):
    parameter: str = Field(..., description="Parameter key from model_map.yaml inputs")
    value: float | int | str | bool | None = Field(..., description="Value to write to the mapped input cell")
    operation: Literal["set", "add", "multiply"] = "set"
    reason: Optional[str] = None


class ActionPlan(BaseModel):
    scenario_name: str = "AI Scenario"
    changes: List[Change]
    requested_outputs: List[str] = Field(default_factory=list)
    raw_command: Optional[str] = None


class AppliedChange(BaseModel):
    parameter: str
    sheet: str
    cell: str
    old_value: Any = None
    new_value: Any = None
    reason: Optional[str] = None


class OutputValue(BaseModel):
    key: str
    sheet: str
    cell: str
    value: Any = None
    type: Optional[str] = None
    unit: Optional[str] = None
    description: Optional[str] = None


class ScenarioResult(BaseModel):
    scenario_name: str
    source_file: str
    scenario_file: str
    engine: str
    recalculated: bool
    applied_changes: List[AppliedChange]
    outputs: Dict[str, OutputValue]
    warnings: List[str] = Field(default_factory=list)
