import os
import shutil
import subprocess
import threading
import time
from typing import List, Optional

import json
import locale
import psutil
import tkinter as tk
from tkinter import messagebox
import ctypes

SERVICE_NAME = "WpcMonSvc"
PROCESS_NAME = "WpcMon.exe"
_PROGRAM_FILES = os.environ.get("ProgramFiles", "")
_PROGRAM_FILES_X86 = os.environ.get("ProgramFiles(x86)", "")
PROGRAM_CANDIDATES: List[str] = [
    os.path.join(_PROGRAM_FILES, "Windows Defender", PROCESS_NAME)
    if _PROGRAM_FILES
    else "",
    os.path.join(_PROGRAM_FILES_X86, "Windows Defender", PROCESS_NAME)
    if _PROGRAM_FILES_X86
    else "",
    PROCESS_NAME,
]

POLL_INTERVAL_MS = 2000
MANIFEST_PATH = os.path.join(os.path.dirname(__file__), "manifest.json")
I18N_PATH = os.path.join(os.path.dirname(__file__), "i18n.json")


def load_manifest() -> dict:
    """Load metadata from manifest.json with defaults."""
    defaults = {
        "name": "RipWPC",
        "version": "0.0.0",
        "author": "Unknown",
        "description": {
            "en": "",
            "fr": "",
        },
    }
    try:
        with open(MANIFEST_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                merged = {**defaults, **data}
                # ensure description is a dict with en/fr keys
                desc = merged.get("description", {})
                if isinstance(desc, str):
                    desc = {"en": desc, "fr": ""}
                if not isinstance(desc, dict):
                    desc = {"en": "", "fr": ""}
                merged["description"] = {
                    "en": desc.get("en", ""),
                    "fr": desc.get("fr", ""),
                }
                return merged
    except Exception:
        pass
    return defaults


def load_translations() -> dict:
    """Load translations from i18n.json with a minimal built-in fallback."""
    fallback = {
        "en": {
            "title": "RipWPC Control",
            "header": "Version: {version} · Author: {author}",
            "status_initializing": "Initializing...",
            "status_checking": "Checking...",
            "status_service_running": "Service: running",
            "status_service_stopped": "Service: stopped",
            "status_program_running": "Program: running",
            "status_program_stopped": "Program: stopped",
            "button_stop": "Stop service + program",
            "button_start": "Start service + program",
            "description_default": "No description provided.",
        },
        "fr": {
            "title": "Contrôle RipWPC",
            "header": "Version : {version} · Auteur : {author}",
            "status_initializing": "Initialisation...",
            "status_checking": "Vérification...",
            "status_service_running": "Service : en cours",
            "status_service_stopped": "Service : arrêté",
            "status_program_running": "Programme : en cours",
            "status_program_stopped": "Programme : arrêté",
            "button_stop": "Arrêter le service + programme",
            "button_start": "Démarrer le service + programme",
            "description_default": "Aucune description renseignée.",
        },
    }
    try:
        with open(I18N_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                return {**fallback, **data}
    except Exception:
        pass
    return fallback


def detect_language(translations: dict) -> str:
    """Select language: FR if OS locale starts with fr, otherwise EN."""
    try:
        lang, _ = locale.getdefaultlocale()
        if lang and str(lang).lower().startswith("fr"):
            return "fr" if "fr" in translations else "en"
    except Exception:
        pass
    return "en"


def _trouver_commande_programme() -> Optional[List[str]]:
    for candidate in PROGRAM_CANDIDATES:
        if not candidate:
            continue
        if os.path.isabs(candidate):
            if os.path.exists(candidate):
                return [candidate]
            continue
        resolved = shutil.which(candidate)
        if resolved:
            return [resolved]
    return None


def est_programme_en_cours(nom: str) -> bool:
    for proc in psutil.process_iter(["name"]):
        if proc.info.get("name") == nom:
            return True
    return False


def arreter_programme(nom: str) -> None:
    for proc in psutil.process_iter(["name"]):
        if proc.info.get("name") == nom:
            try:
                proc.kill()
            except Exception:
                pass


def est_service_en_cours(service: str) -> bool:
    result = subprocess.run(
        ["sc", "query", service], capture_output=True, text=True, check=False
    )
    if not result.stdout:
        return False
    for line in result.stdout.splitlines():
        if "STATE" in line:
            return "RUNNING" in line
    return False


def demarrer_service(service: str) -> None:
    subprocess.run(["sc", "start", service], check=False)
    time.sleep(2)


def arreter_service(service: str) -> None:
    subprocess.run(["sc", "stop", service], check=False)
    time.sleep(3)

    result = subprocess.run(
        ["sc", "queryex", service], capture_output=True, text=True, check=False
    )
    for line in result.stdout.splitlines():
        if "PID" in line:
            try:
                pid_str = line.split(":", 1)[1].strip()
                pid = int(pid_str)
                psutil.Process(pid).kill()
            except Exception:
                pass


def demarrer_programme() -> None:
    commande = _trouver_commande_programme()
    if not commande:
        return
    try:
        subprocess.Popen(commande)
    except Exception:
        pass


class RipWPCControl:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.meta = load_manifest()
        self.translations = load_translations()
        self.lang = detect_language(self.translations)
        self.status_var = tk.StringVar(value=self._t("status_initializing"))
        self.button_var = tk.StringVar(value=self._t("status_checking"))
        self.action_lock = threading.Lock()

        tk.Label(root, text=self._t("title"), font=(None, 16, "bold")).pack(pady=(10, 0))

        header = tk.Frame(root)
        header.pack(pady=(4, 0))
        header_text = self._t("header").format(
            version=self.meta.get("version", "N/A"),
            author=self.meta.get("author", "N/A"),
        )
        tk.Label(header, text=header_text).pack(side=tk.LEFT)
        tk.Button(header, text="?", width=3, command=self.show_description).pack(side=tk.LEFT, padx=(6, 0))

        tk.Label(root, textvariable=self.status_var, justify="center").pack(pady=(8, 4))
        self.toggle_button = tk.Button(
            root,
            textvariable=self.button_var,
            command=self.on_toggle,
            width=32,
            height=2,
        )
        self.toggle_button.pack(pady=(0, 10))

        # Initial status check in a background thread to avoid blocking UI
        threading.Thread(target=self._initial_status_check, daemon=True).start()

    def _initial_status_check(self) -> None:
        """Check service/program state on startup and update UI."""
        self._update_status_labels()
        # Start regular polling after initial check
        self.root.after(POLL_INTERVAL_MS, self.refresh_status)

    def on_toggle(self) -> None:
        if self.action_lock.locked():
            return
        threading.Thread(target=self._run_toggle, daemon=True).start()

    def _run_toggle(self) -> None:
        with self.action_lock:
            self._set_button_state("disabled")
            running = self._aggregated_state()
            if running:
                arreter_programme(PROCESS_NAME)
                arreter_service(SERVICE_NAME)
            else:
                demarrer_service(SERVICE_NAME)
                demarrer_programme()
            self._update_status_labels()
            self._set_button_state("normal")

    def _aggregated_state(self) -> bool:
        return est_service_en_cours(SERVICE_NAME) or est_programme_en_cours(PROCESS_NAME)

    def refresh_status(self) -> None:
        """Periodic status update (called after initial startup check)."""
        self._update_status_labels()
        self.root.after(POLL_INTERVAL_MS, self.refresh_status)

    def _update_status_labels(self) -> None:
        service_running = est_service_en_cours(SERVICE_NAME)
        program_running = est_programme_en_cours(PROCESS_NAME)
        status_text = (
            f"{self._t('status_service_running' if service_running else 'status_service_stopped')} · "
            f"{self._t('status_program_running' if program_running else 'status_program_stopped')}"
        )
        self.status_var.set(status_text)
        if service_running or program_running:
            self.button_var.set(self._t("button_stop"))
        else:
            self.button_var.set(self._t("button_start"))

    def _set_button_state(self, state: str) -> None:
        self.root.after(0, lambda: self.toggle_button.config(state=state))

    def show_description(self) -> None:
        desc_map = self.meta.get("description", {}) if isinstance(self.meta, dict) else {}
        desc = ""
        if isinstance(desc_map, dict):
            desc = desc_map.get(self.lang, "") or desc_map.get("en", "")
        if not desc:
            desc = self._t("description_default")
        messagebox.showinfo(self.meta.get("name", "RipWPC"), desc)

    def _t(self, key: str) -> str:
        lang_map = self.translations.get(self.lang, {})
        fallback_map = self.translations.get("en", {})
        return lang_map.get(key) or fallback_map.get(key, key)


def hide_console() -> None:
    """Masque la fenêtre console si elle existe."""
    try:
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)
    except Exception:
        pass


def main() -> None:
    hide_console()
    root = tk.Tk()
    root.title("RipWPC")
    root.resizable(False, False)
    RipWPCControl(root)
    root.mainloop()


if __name__ == "__main__":
    main()
