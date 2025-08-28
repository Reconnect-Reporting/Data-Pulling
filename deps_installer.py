# deps_installer.py
# -*- coding: utf-8 -*-
import sys
import subprocess
import importlib.util
from typing import Dict, List, Tuple

class DependencyInstaller:
    """Checks & installs third-party deps used by your app."""

    REQUIRED: Dict[str, str] = {
        "pandas": "pandas",
        "numpy": "numpy",
        "openpyxl": "openpyxl",
        "win32com.client": "pywin32",
        "selenium": "selenium",
    }

    def __init__(self, quiet: bool = True):
        self.quiet = quiet
        self._log: List[str] = []

    def ensure(self) -> Dict[str, str]:
        status: Dict[str, str] = {}
        missing = self._missing()

        if not missing:
            for pip_name in set(self.REQUIRED.values()):
                status[pip_name] = "ok"
            self._add("All dependencies already satisfied ✅")
            return status

        self._add("Missing packages detected:")
        for mod, pip_name in missing:
            self._add(f"  - {mod}  (pip: {pip_name})")

        targets = sorted({pip for _, pip in missing})
        ok, failed = self._pip_install(targets)

        # Re-verify
        for import_path, pip_name in self.REQUIRED.items():
            if pip_name in ok:
                status[pip_name] = "installed" if self._can_import(import_path) else "failed: import still missing"
            elif pip_name in failed:
                status[pip_name] = f"failed: {failed[pip_name]}"
            else:
                status[pip_name] = "ok"

        if status.get("pywin32", "").startswith("installed"):
            self._add("Note: If Outlook automation fails, restart Python/Windows to finalize pywin32.")
        return status

    def get_log(self) -> str:
        return "\n".join(self._log)

    # ---- internals
    def _missing(self) -> List[Tuple[str, str]]:
        out: List[Tuple[str, str]] = []
        for import_path, pip_name in self.REQUIRED.items():
            if not self._can_import(import_path):
                out.append((import_path, pip_name))
        return out

    def _can_import(self, import_path: str) -> bool:
        return importlib.util.find_spec(import_path) is not None

    def _pip_install(self, pkgs: List[str]) -> Tuple[List[str], Dict[str, str]]:
        ok: List[str] = []
        failed: Dict[str, str] = {}
        in_venv = (hasattr(sys, "base_prefix") and sys.base_prefix != sys.prefix) or \
                  (hasattr(sys, "real_prefix") and sys.real_prefix != sys.prefix)

        cmd_base = [sys.executable, "-m", "pip", "install", "--upgrade"]
        if not in_venv:
            cmd_base.append("--user")

        for pkg in pkgs:
            cmd = cmd_base + [pkg]
            try:
                self._add(f"Installing {pkg} ...")
                res = subprocess.run(
                    cmd,
                    capture_output=self.quiet,
                    text=True,
                    check=False
                )
                if res.returncode == 0:
                    ok.append(pkg)
                    self._add(f"✓ {pkg} installed/updated")
                else:
                    reason = (res.stderr or res.stdout or "unknown error").strip().splitlines()[-1]
                    failed[pkg] = reason
                    self._add(f"✗ {pkg} failed: {reason}")
            except Exception as e:
                failed[pkg] = str(e)
                self._add(f"✗ {pkg} failed: {e}")
        return ok, failed

    def _add(self, msg: str):
        self._log.append(msg)
