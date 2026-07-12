"""
Central configuration for capiq-excel.

Feature flags and runtime settings that control dialect selection,
add-in mode, refresh behavior, and formula option defaults.

Config resolution order:
  1. Explicit arguments passed to functions
  2. Environment variables (CAPIQ_*)
  3. Defaults (legacy-compatible)
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


class FormulaDialect(Enum):
    """Which formula family to emit."""
    AUTO = "auto"
    CIQ = "ciq"
    SPG = "spg"
    SNL = "snl"


class AddinMode(Enum):
    """Which add-in to target."""
    AUTO = "auto"
    LEGACY = "legacy"
    PRO = "pro"


class RefreshScope(Enum):
    """Granularity of refresh operations."""
    WORKSHEET = "worksheet"
    WORKBOOK = "workbook"
    SELECTION = "selection"


class MetricType(Enum):
    """Category of data being queried — determines formula structure."""
    FINANCIAL = "financial"
    MARKET = "market"
    OWNERSHIP = "ownership"
    ESTIMATES = "estimates"
    ID_LOOKUP = "id_lookup"


@dataclass
class FormulaOptions:
    """Default options applied to generated formulas.

    For SPG/SNL functions, these are serialized into an options string like:
        "Curr=USD,Mag=Millions,ConvMethod=Recommended"

    CIQ formulas don't use this options string format.
    """
    currency: Optional[str] = None          # ISO code e.g. "USD", or "RC" for reported
    magnitude: Optional[str] = None         # "Standard"|"Actuals"|"Thousands"|"Millions"|"Billions"|"Trillions" or "0"|"3"|"6"|"9"|"12"
    conversion_method: Optional[str] = None # "Recommended" | "MRSpot"
    null_display: Optional[str] = None      # e.g. "NA", "0", "#NA", "NAZ"
    terminology: Optional[str] = None       # language code e.g. "en-US", "ja-JP", "zh-CN"

    def to_spg_options_string(self) -> Optional[str]:
        """Serialize non-None options into SPG/SNL options string format.

        Returns None if no options are set.
        Example output: "Curr=USD,Mag=Millions,ConvMethod=Recommended"
        """
        parts: list[str] = []
        if self.currency is not None:
            parts.append(f"Curr={self.currency}")
        if self.magnitude is not None:
            parts.append(f"Mag={self.magnitude}")
        if self.conversion_method is not None:
            parts.append(f"ConvMethod={self.conversion_method}")
        if self.null_display is not None:
            parts.append(f"NullValue={self.null_display}")
        if self.terminology is not None:
            parts.append(f"Terminology={self.terminology}")
        return ",".join(parts) if parts else None


def _safe_enum(enum_cls, env_var: str, default):
    """Parse an environment variable as an enum value, falling back to default."""
    raw = os.environ.get(env_var, default.value).lower()
    try:
        return enum_cls(raw)
    except ValueError:
        print(f'Warning: Invalid {env_var} "{raw}", defaulting to "{default.value}"')
        return default


def _safe_int(env_var: str, default: int) -> int:
    """Parse an environment variable as an integer, falling back to default."""
    raw = os.environ.get(env_var, str(default))
    try:
        return int(raw)
    except ValueError:
        print(f'Warning: Invalid {env_var} "{raw}", defaulting to {default}')
        return default


@dataclass
class RetryConfig:
    """Retry and timeout parameters."""
    max_retries: int = 3
    timeout_seconds: int = 240
    restart_interval: int = 500  # restart Excel every N workbooks
    retry_delay_seconds: int = 30


@dataclass
class CapiqConfig:
    """Top-level configuration for a capiq-excel session."""

    # Feature flags (Phase 0)
    formula_dialect: FormulaDialect = FormulaDialect.AUTO
    addin_mode: AddinMode = AddinMode.AUTO
    refresh_scope: RefreshScope = RefreshScope.WORKSHEET

    # Formula defaults
    freq: str = "Q"
    num_periods: int = 80
    formula_options: FormulaOptions = field(default_factory=FormulaOptions)

    # Retry / reliability
    retry: RetryConfig = field(default_factory=RetryConfig)

    @classmethod
    def from_env(cls) -> CapiqConfig:
        """Build config from environment variables, falling back to defaults."""
        return cls(
            formula_dialect=_safe_enum(FormulaDialect, "CAPIQ_FORMULA_DIALECT", FormulaDialect.AUTO),
            addin_mode=_safe_enum(AddinMode, "CAPIQ_ADDIN_MODE", AddinMode.AUTO),
            refresh_scope=_safe_enum(RefreshScope, "CAPIQ_REFRESH_SCOPE", RefreshScope.WORKSHEET),
            freq=os.environ.get("CAPIQ_FREQ", "Q"),
            num_periods=_safe_int("CAPIQ_NUM_PERIODS", 80),
            formula_options=FormulaOptions(
                currency=os.environ.get("CAPIQ_CURRENCY"),
                magnitude=os.environ.get("CAPIQ_MAGNITUDE"),
                conversion_method=os.environ.get("CAPIQ_CONVERSION_METHOD"),
                null_display=os.environ.get("CAPIQ_NULL_DISPLAY"),
                terminology=os.environ.get("CAPIQ_TERMINOLOGY"),
            ),
            retry=RetryConfig(
                max_retries=_safe_int("CAPIQ_MAX_RETRIES", 3),
                timeout_seconds=_safe_int("CAPIQ_TIMEOUT", 240),
                restart_interval=_safe_int("CAPIQ_RESTART_INTERVAL", 500),
                retry_delay_seconds=_safe_int("CAPIQ_RETRY_DELAY", 30),
            ),
        )

    def resolve_dialect(self, ciq_compat_available: bool = True) -> FormulaDialect:
        """Resolve AUTO dialect based on runtime detection."""
        if self.formula_dialect != FormulaDialect.AUTO:
            return self.formula_dialect
        # AUTO: prefer CIQ when compatibility mode is available (safe default)
        return FormulaDialect.CIQ if ciq_compat_available else FormulaDialect.SPG
