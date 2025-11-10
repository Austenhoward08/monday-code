"""Configuration utilities for Monday.com exporter."""

from __future__ import annotations

import os
from typing import Any, Optional

from pydantic import BaseModel, Field, HttpUrl, SecretStr, ValidationError


class Settings(BaseModel):
    """Validated settings for the exporter runtime."""

    api_token: SecretStr = Field(..., description="Monday.com API token.")
    api_url: HttpUrl = Field(
        "https://api.monday.com/v2",
        description="Base URL for the Monday.com GraphQL API.",
    )
    timeout_seconds: float = Field(
        30.0,
        ge=1.0,
        description="Network timeout to use for Monday.com API requests.",
    )
    page_size: int = Field(
        500,
        ge=1,
        le=1000,
        description="Number of items to request per page from Monday.com API.",
    )

    @classmethod
    def from_env(
        cls,
        api_token: Optional[str] = None,
        overrides: Optional[dict[str, Any]] = None,
    ) -> "Settings":
        """
        Build a Settings instance using environment variables, with optional overrides.

        Args:
            api_token: Explicit API token. Falls back to MONDAY_API_TOKEN.
            overrides: Additional keyword arguments to pass to the model.

        Raises:
            RuntimeError: If no API token is provided.
            ValidationError: If settings validation fails.
        """

        token = api_token or os.getenv("MONDAY_API_TOKEN")
        if not token:
            raise RuntimeError(
                "Monday.com API token is required. "
                "Provide it via the --api-token option or the MONDAY_API_TOKEN environment variable."
            )

        data: dict[str, Any] = {"api_token": token}
        if overrides:
            data.update(overrides)

        try:
            return cls.model_validate(data)
        except ValidationError as exc:  # pragma: no cover - validation errors are informative enough
            raise RuntimeError(f"Invalid configuration: {exc}") from exc


__all__ = ["Settings"]
