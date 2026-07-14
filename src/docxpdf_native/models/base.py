from __future__ import annotations

from pydantic import BaseModel, ConfigDict


class FrozenModel(BaseModel):
    """Common configuration for immutable conversion data."""

    model_config = ConfigDict(
        frozen=True,
        extra="forbid",
        ser_json_bytes="base64",
        val_json_bytes="base64",
    )
