"""Domain models for Monday.com board data."""

from __future__ import annotations

from datetime import datetime
from typing import Any, List, Optional

from pydantic import BaseModel, Field, computed_field


class BoardColumn(BaseModel):
    """Represents a Monday.com board column."""

    id: str
    title: str
    type: str


class ColumnValue(BaseModel):
    """Represents a column value for a Monday.com item."""

    id: str
    text: str = ""
    title: str
    type: str
    value: Optional[str] = None

    @computed_field  # type: ignore[misc]
    @property
    def parsed_value(self) -> Any:
        """Return JSON-decoded `value` when possible, otherwise the raw text."""
        if not self.value:
            return self.text

        try:
            import json

            return json.loads(self.value)
        except (ValueError, TypeError):
            return self.text


class Group(BaseModel):
    """Represents a Monday.com group of items."""

    id: str
    title: str
    color: Optional[str] = None
    position: Optional[str] = None


class Person(BaseModel):
    """Represents a Monday.com user."""

    id: Optional[int] = None
    name: Optional[str] = None


class Item(BaseModel):
    """Represents an item (row) on a Monday.com board."""

    id: str
    name: str
    group: Optional[Group] = None
    creator: Optional[Person] = None
    created_at: Optional[datetime] = Field(None, alias="created_at")
    updated_at: Optional[datetime] = Field(None, alias="updated_at")
    column_values: List[ColumnValue] = Field(default_factory=list, alias="column_values")
    subitems: Optional[List["Item"]] = Field(default=None, alias="subitems")

    class Config:
        populate_by_name = True

    def column_value_by_id(self, column_id: str) -> Optional[ColumnValue]:
        """Return the column value for the given column id, if present."""
        for column_value in self.column_values:
            if column_value.id == column_id:
                return column_value
        return None


class Board(BaseModel):
    """Represents a Monday.com board."""

    id: str
    name: str
    description: Optional[str] = None
    state: Optional[str] = None
    columns: List[BoardColumn] = Field(default_factory=list)
    groups: List[Group] = Field(default_factory=list)
    items: List[Item] = Field(default_factory=list)

    def ordered_columns(self) -> List[BoardColumn]:
        """Return columns sorted in the order returned by the API."""
        return self.columns


__all__ = [
    "Board",
    "BoardColumn",
    "ColumnValue",
    "Group",
    "Item",
    "Person",
]


Item.model_rebuild()
