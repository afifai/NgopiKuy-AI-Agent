from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any


HERMES_HOME = Path(os.environ.get("HERMES_HOME", os.path.expanduser("~/.hermes_kopi")))
PAIRING_PATH = HERMES_HOME / "config" / "user_pairing.json"


@dataclass
class IdentityResult:
    ok: bool
    sheet_name: str | None = None
    role: str = "member"
    reason: str = ""


def _norm(s: str | None) -> str:
    return (s or "").strip().lower()


def load_pairing() -> dict[str, Any]:
    if not PAIRING_PATH.exists():
        return {
            "version": 1,
            "strict_mode": True,
            "allow_username_fallback": False,
            "allow_full_name_fallback": False,
            "members": [],
        }

    with PAIRING_PATH.open("r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, dict):
        raise ValueError("user_pairing.json harus object JSON")

    data.setdefault("strict_mode", True)
    data.setdefault("allow_username_fallback", False)
    data.setdefault("allow_full_name_fallback", False)
    data.setdefault("members", [])
    return data


def resolve_member(
    telegram_user_id: int | None = None,
    username: str | None = None,
    full_name: str | None = None,
) -> IdentityResult:
    cfg = load_pairing()
    strict = bool(cfg.get("strict_mode", True))
    by_username = bool(cfg.get("allow_username_fallback", False))
    by_full_name = bool(cfg.get("allow_full_name_fallback", False))
    members = cfg.get("members", []) or []

    # 1) Primary: telegram_user_id exact match
    if telegram_user_id is not None:
        for m in members:
            if not m.get("active", True):
                continue
            if m.get("telegram_user_id") == telegram_user_id:
                return IdentityResult(True, m.get("sheet_name"), m.get("role", "member"), "ok:user_id")

    # 2) Optional fallback: username (exact, normalized)
    if by_username and username:
        u = _norm(username).lstrip("@")
        candidates = [
            m for m in members
            if m.get("active", True) and _norm(m.get("username")).lstrip("@") == u
        ]
        if len(candidates) == 1:
            m = candidates[0]
            return IdentityResult(True, m.get("sheet_name"), m.get("role", "member"), "ok:username")
        if len(candidates) > 1:
            return IdentityResult(False, None, "member", "username_duplicated")

    # 3) Optional fallback: full name (exact, normalized)
    if by_full_name and full_name:
        n = _norm(full_name)
        candidates = [
            m for m in members
            if m.get("active", True) and _norm(m.get("full_name")) == n
        ]
        if len(candidates) == 1:
            m = candidates[0]
            return IdentityResult(True, m.get("sheet_name"), m.get("role", "member"), "ok:full_name")
        if len(candidates) > 1:
            return IdentityResult(False, None, "member", "full_name_duplicated")

    if strict:
        return IdentityResult(False, None, "member", "not_registered")

    # non-strict mode: allow direct display name as sheet name fallback
    if full_name:
        return IdentityResult(True, full_name.strip(), "member", "warn:non_strict_full_name")

    return IdentityResult(False, None, "member", "unresolved")


def is_admin(role: str | None) -> bool:
    return _norm(role) == "admin"
