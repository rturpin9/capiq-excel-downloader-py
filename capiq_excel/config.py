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
        dialect_str = os.environ.get("CAPIQ_FORMULA_DIALECT", "auto").lower()
        mode_str = os.environ.get("CAPIQ_ADDIN_MODE", "auto").lower()
        scope_str = os.environ.get("CAPIQ_REFRESH_SCOPE", "worksheet").lower()

        try:
            dialect = FormulaDialect(dialect_str)
        except ValueError:
            print(f'Warning: Invalid CAPIQ_FORMULA_DIALECT "{dialect_str}", defaulting to "auto"')
            dialect = FormulaDialect.AUTO
        try:
            mode = AddinMode(mode_str)
        except ValueError:
            print(f'Warning: Invalid CAPIQ_ADDIN_MODE "{mode_str}", defaulting to "auto"')
            mode = AddinMode.AUTO
        try:
            scope = RefreshScope(scope_str)
        except ValueError:
            print(f'Warning: Invalid CAPIQ_REFRESH_SCOPE "{scope_str}", defaulting to "worksheet"')
            scope = RefreshScope.WORKSHEET

        def _safe_int(env_var: str, default: int) -> int:
            raw = os.environ.get(env_var, str(default))
            try:
                return int(raw)
            except ValueError:
                print(f'Warning: Invalid {env_var} "{raw}", defaulting to {default}')
                return default

        return cls(
            formula_dialect=dialect,
            addin_mode=mode,
            refresh_scope=scope,
            freq=os.environ.get("CAPIQ_FREQ", "Q"),
            num_periods=_safe_int("CAPIQ_NUM_PERIODS", 80),
            retry=RetryConfig(
                max_retries=_safe_int("CAPIQ_MAX_RETRIES", 3),
                timeout_seconds=_safe_int("CAPIQ_TIMEOUT", 240),
                restart_interval=_safe_int("CAPIQ_RESTART_INTERVAL", 500),
            ),
        )

    def resolve_dialect(self, ciq_compat_available: bool = True) -> FormulaDialect:
        """Resolve AUTO dialect based on runtime detection."""
        if self.formula_dialect != FormulaDialect.AUTO:
            return self.formula_dialect
        # AUTO: prefer CIQ when compatibility mode is available (safe default)
        return FormulaDialect.CIQ if ciq_compat_available else FormulaDialect.SPG
