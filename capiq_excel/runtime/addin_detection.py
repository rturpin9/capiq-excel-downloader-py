"""
Add-in detection and runtime profiling for Capital IQ Excel plugins.

Enumerates Excel COM add-ins and Windows registry keys to detect whether
the legacy CIQ plugin, Capital IQ Pro, or both are installed and active.
Returns a normalized RuntimeProfile used by the rest of the pipeline to
select dialect and loading strategy.

Registry keys (from Tech Guide p.6):
  - Settings:      HKCU\\Software\\SNL Financial\\SNL Office
  - Load behavior: HKCU\\Software\\Microsoft\\Office\\Excel\\Addins\\SNL.Clients.Office.Excel.ExcelAddIn
  - OPEN keys:     HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\\Options
  - Install path:  %ProgramFiles%\\SNL Financial\\SNLxL  or  %ProgramFiles%\\SP Global Market Intelligence
  - Logs:          %LocalAppData%\\SPGMI
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Optional

from capiq_excel.config import AddinMode

logger = logging.getLogger(__name__)

# Known add-in identifiers (display names / ProgIDs)
LEGACY_CIQ_NAMES = frozenset({
    "S&P Capital IQ Excel Plug-in",
    "Capital IQ Excel Plug-in",
})
PRO_ADDIN_NAMES = frozenset({
    "S&P Capital IQ Pro",
    "Capital IQ Pro",
    "S&P Capital IQ Pro Office",
    "S&P Cap IQ Pro Excel Add-In",
})
PRO_PROGIDS = frozenset({
    "SNL.Clients.Office.Excel.ExcelAddIn",
    "SPGMI.ExcelShell",
})

# Registry paths
_REG_SNL_OFFICE = r"Software\SNL Financial\SNL Office"
_REG_SNL_OFFICE_WOW64 = r"Software\Wow6432Node\SNL Financial\SNL Office"
_REG_EXCEL_ADDINS = r"Software\Microsoft\Office\Excel\Addins\SNL.Clients.Office.Excel.ExcelAddIn"

# Pro COM add-in ProgID (from OPEN1 registry key)
PRO_PROGID = "SNL.Clients.Office.Excel.Functions"
# Legacy CIQ UDF name (from OPEN2 registry key)
CIQ_UDF_NAME = "ciqfunctions.udf"


@dataclass
class RuntimeProfile:
    """Normalized snapshot of the detected Capital IQ environment."""
    addin_mode: str = "unknown"  # "legacy" | "pro" | "mixed" | "unknown"
    ciq_compat_enabled: Optional[bool] = None
    available_dialects: set[str] = field(default_factory=set)
    legacy_addin_name: Optional[str] = None
    pro_addin_name: Optional[str] = None
    excel_version: Optional[str] = None
    excel_bitness: Optional[str] = None
    # Registry-sourced info
    pro_installed: bool = False
    pro_load_behavior: Optional[int] = None  # 3 = load at startup
    install_path: Optional[str] = None
    log_path: Optional[str] = None


def detect_runtime(excel) -> RuntimeProfile:
    """
    Inspect a running Excel COM instance and Windows registry,
    then return a RuntimeProfile.

    Args:
        excel: A win32com Excel.Application object.

    Returns:
        RuntimeProfile describing the detected add-in environment.
    """
    profile = RuntimeProfile()

    # Excel version info
    try:
        profile.excel_version = str(excel.Version)
        profile.excel_bitness = _detect_bitness(excel)
    except Exception:
        logger.warning("Could not read Excel version info")

    # Phase 1: Check Windows registry for installation state
    _check_registry(profile)

    # Phase 2: Scan COM add-ins on the running Excel instance
    found_legacy = False
    found_pro = False

    try:
        for addin in excel.COMAddIns:
            name = addin.Description or ""
            progid = getattr(addin, "ProgId", "") or ""
            connected = addin.Connect

            logger.debug("COM add-in: %s (ProgID=%s, connected=%s)", name, progid, connected)

            if _matches_any(name, LEGACY_CIQ_NAMES):
                found_legacy = True
                profile.legacy_addin_name = name
                if connected:
                    profile.available_dialects.add("ciq")

            if _matches_any(name, PRO_ADDIN_NAMES) or progid in PRO_PROGIDS or progid == PRO_PROGID:
                found_pro = True
                profile.pro_addin_name = name
                if connected:
                    profile.available_dialects.update({"spg", "snl"})
    except Exception:
        logger.warning("Could not enumerate COM add-ins", exc_info=True)

    # Phase 3: Scan worksheet-function add-ins (AddIns collection)
    try:
        for addin in excel.AddIns:
            name = getattr(addin, "Name", "") or ""
            installed = getattr(addin, "Installed", False)

            if _matches_any(name, LEGACY_CIQ_NAMES) and installed:
                found_legacy = True
                profile.legacy_addin_name = profile.legacy_addin_name or name
                profile.available_dialects.add("ciq")

            if _matches_any(name, PRO_ADDIN_NAMES) and installed:
                found_pro = True
                profile.pro_addin_name = profile.pro_addin_name or name
                profile.available_dialects.update({"spg", "snl"})
    except Exception:
        logger.warning("Could not enumerate worksheet add-ins", exc_info=True)

    # Phase 4: Determine CIQ compat from registry if Pro is present
    if found_pro or profile.pro_installed:
        found_pro = True
        ciq_compat = _read_ciq_compat_from_registry()
        if ciq_compat is not None:
            profile.ciq_compat_enabled = ciq_compat
            if ciq_compat:
                profile.available_dialects.add("ciq")

    # Phase 5: Determine overall mode
    if found_legacy and found_pro:
        profile.addin_mode = "mixed"
        if profile.ciq_compat_enabled is None:
            profile.ciq_compat_enabled = "ciq" in profile.available_dialects
    elif found_pro:
        profile.addin_mode = "pro"
        if profile.ciq_compat_enabled is None:
            profile.ciq_compat_enabled = "ciq" in profile.available_dialects
    elif found_legacy:
        profile.addin_mode = "legacy"
        profile.ciq_compat_enabled = True
    else:
        profile.addin_mode = "unknown"

    return profile


def load_best_addin(excel, preferred_mode: AddinMode, profile: RuntimeProfile) -> str:
    """
    Load the most appropriate add-in based on preference and availability.

    Returns the name of the add-in that was loaded/connected.
    Raises AddinNotFoundError if no suitable add-in is found.
    """
    from capiq_excel.exceptions import AddinNotFoundError

    # Build priority order based on preference
    if preferred_mode == AddinMode.LEGACY:
        candidates = [profile.legacy_addin_name]
    elif preferred_mode == AddinMode.PRO:
        candidates = [profile.pro_addin_name, profile.legacy_addin_name]
    else:
        # AUTO: prefer whatever is available, Pro first if both present
        if profile.addin_mode == "mixed":
            candidates = [profile.pro_addin_name, profile.legacy_addin_name]
        elif profile.addin_mode == "pro":
            candidates = [profile.pro_addin_name]
        elif profile.addin_mode == "legacy":
            candidates = [profile.legacy_addin_name]
        else:
            candidates = []

    # Try each candidate
    for name in candidates:
        if name is None:
            continue
        try:
            _connect_addin(excel, name)
            logger.info("Loaded add-in: %s", name)
            return name
        except Exception:
            logger.warning("Failed to load add-in: %s", name, exc_info=True)
            continue

    raise AddinNotFoundError(
        f"No suitable Capital IQ add-in found. "
        f"Mode={preferred_mode.value}, profile={profile.addin_mode}, "
        f"legacy={profile.legacy_addin_name}, pro={profile.pro_addin_name}"
    )


def log_startup_diagnostics(profile: RuntimeProfile, config) -> None:
    """Log a summary of the runtime environment at startup."""
    logger.info("=== Capital IQ Runtime Diagnostics ===")
    logger.info("Excel version: %s (%s)", profile.excel_version, profile.excel_bitness)
    logger.info("Add-in mode: %s", profile.addin_mode)
    logger.info("Legacy add-in: %s", profile.legacy_addin_name or "(not found)")
    logger.info("Pro add-in: %s", profile.pro_addin_name or "(not found)")
    logger.info("Pro installed (registry): %s", profile.pro_installed)
    logger.info("Pro load behavior: %s", profile.pro_load_behavior)
    logger.info("CIQ compat enabled: %s", profile.ciq_compat_enabled)
    logger.info("Available dialects: %s", profile.available_dialects or "(none)")
    logger.info("Install path: %s", profile.install_path or "(unknown)")
    logger.info("Config dialect: %s", config.formula_dialect.value)
    logger.info("Config addin_mode: %s", config.addin_mode.value)
    logger.info("Config refresh_scope: %s", config.refresh_scope.value)
    logger.info("======================================")


# --- registry helpers ---

def _check_registry(profile: RuntimeProfile) -> None:
    """Read Windows registry to detect Pro installation state."""
    try:
        import winreg
    except ImportError:
        logger.debug("winreg not available (non-Windows platform)")
        return

    # Check Pro add-in load behavior
    for reg_path in (_REG_EXCEL_ADDINS,):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path) as key:
                load_behavior, _ = winreg.QueryValueEx(key, "LoadBehavior")
                profile.pro_installed = True
                profile.pro_load_behavior = load_behavior
                logger.debug("Pro LoadBehavior=%s from %s", load_behavior, reg_path)
        except FileNotFoundError:
            pass
        except Exception:
            logger.debug("Could not read registry key %s", reg_path, exc_info=True)

    # Check SNL Office settings for install path and other info
    for reg_path in (_REG_SNL_OFFICE, _REG_SNL_OFFICE_WOW64):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path) as key:
                profile.pro_installed = True
                logger.debug("Found SNL Office registry key at %s", reg_path)
                # Try to read proxy and other settings (informational)
                try:
                    proxy, _ = winreg.QueryValueEx(key, "ProxyServer")
                    logger.debug("Proxy configured: %s", proxy)
                except FileNotFoundError:
                    pass
        except FileNotFoundError:
            pass
        except Exception:
            logger.debug("Could not read registry key %s", reg_path, exc_info=True)


def _read_ciq_compat_from_registry() -> Optional[bool]:
    """Check registry for the CIQ backward compatibility toggle.

    The CIQ compat setting can be toggled in the Pro Settings UI (v21.08+)
    and is stored under the SNL Office registry key.

    Returns True/False if found, None if not detectable.
    """
    try:
        import winreg
    except ImportError:
        return None

    for reg_path in (_REG_SNL_OFFICE, _REG_SNL_OFFICE_WOW64):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path) as key:
                # The exact value name may vary; common patterns:
                # Check for DisableCIQUDF first (inverted: 0 = enabled, 1 = disabled)
                try:
                    val, _ = winreg.QueryValueEx(key, "DisableCIQUDF")
                    if isinstance(val, int):
                        return val == 0  # 0 means NOT disabled = enabled
                    if isinstance(val, str):
                        return val.lower() in ("0", "false", "no")
                except FileNotFoundError:
                    pass

                # Fall back to other possible key names
                for value_name in ("CIQCompatibility", "EnableCIQFunctions", "CIQBackwardsCompatibility"):
                    try:
                        val, _ = winreg.QueryValueEx(key, value_name)
                        # Interpret: 1/True/"Yes" = enabled
                        if isinstance(val, int):
                            return val != 0
                        if isinstance(val, str):
                            return val.lower() in ("1", "true", "yes")
                    except FileNotFoundError:
                        continue
        except FileNotFoundError:
            continue
        except Exception:
            logger.debug("Could not read CIQ compat from %s", reg_path, exc_info=True)

    return None


# --- COM helpers ---

def _matches_any(name: str, known_names: frozenset[str]) -> bool:
    """Case-insensitive substring match against known add-in names."""
    lower = name.lower()
    return any(known.lower() in lower for known in known_names)


def _detect_bitness(excel) -> str:
    """Attempt to determine Excel bitness (32/64-bit)."""
    try:
        if hasattr(excel, "Hinstance"):
            import struct
            return f"{struct.calcsize('P') * 8}-bit"
    except Exception:
        pass
    return "unknown"


def _connect_addin(excel, name: str) -> None:
    """Connect a COM add-in by display name."""
    for addin in excel.COMAddIns:
        desc = addin.Description or ""
        if desc.lower() == name.lower():
            if not addin.Connect:
                addin.Connect = True
            return
    raise RuntimeError(f"Add-in '{name}' not found in COM add-ins collection")
