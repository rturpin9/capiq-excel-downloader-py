"""
Add-in loading — unified entry point for legacy and Pro modes.

The legacy load_capiq() is preserved for backward compatibility.
New code should use load_capiq_addin() which respects the runtime
profile and config.
"""
from __future__ import annotations

import logging
from typing import Optional

from exceldriver.addin import load_addin

from capiq_excel.config import AddinMode, CapiqConfig
from capiq_excel.runtime.addin_detection import (
    RuntimeProfile,
    detect_runtime,
    load_best_addin,
    log_startup_diagnostics,
)

logger = logging.getLogger(__name__)


def load_capiq(excel):
    """Legacy entry point — loads the hardcoded CIQ add-in name.

    Preserved for backward compatibility when running in legacy mode.
    """
    return load_addin(excel, 'S&P Capital IQ Excel Plug-in')


def load_capiq_addin(excel, config: Optional[CapiqConfig] = None) -> tuple[RuntimeProfile, str]:
    """Detect the runtime environment and load the best available add-in.

    Args:
        excel: COM Excel.Application instance.
        config: Optional config; defaults to CapiqConfig.from_env().

    Returns:
        (profile, addin_name) tuple.
    """
    if config is None:
        config = CapiqConfig.from_env()

    profile = detect_runtime(excel)
    log_startup_diagnostics(profile, config)

    # In legacy mode, use the original hardcoded loader for maximum compat
    if config.addin_mode == AddinMode.LEGACY:
        logger.info("Using legacy add-in loader")
        load_capiq(excel)
        return profile, "S&P Capital IQ Excel Plug-in"

    addin_name = load_best_addin(excel, config.addin_mode, profile)
    return profile, addin_name
