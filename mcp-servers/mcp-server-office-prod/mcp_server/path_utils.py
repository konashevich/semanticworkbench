"""Utilities for normalizing user-provided paths.

Accepts:
- Absolute Windows paths (e.g., C:\\Users\\...)
- Relative paths (resolved against cwd, parent dirs, and optional env roots)
- file:// URIs (e.g., file:///C:/Users/... or UNC URIs)
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Iterable
from urllib.parse import urlparse, unquote
from urllib.request import url2pathname


_ENV_ROOT_VARS: tuple[str, ...] = (
    # Project-specific
    "OFFICE_MCP_WORKSPACE_ROOT",
    "MCP_OFFICE_WORKSPACE_ROOT",
    "MCP_WORKSPACE_ROOT",
    # Common patterns
    "WORKSPACE_ROOT",
    "WORKSPACE_FOLDER",
    # VS Code terminal envs
    "VSCODE_WORKSPACE_FOLDER",
    "VSCODE_WORKSPACE_ROOT",
)


def _strip_quotes(s: str) -> str:
    s = s.strip()
    if (s.startswith("\"") and s.endswith("\"")) or (s.startswith("'") and s.endswith("'")):
        return s[1:-1]
    return s


def _from_file_uri(uri: str) -> Path:
    parsed = urlparse(uri)
    if parsed.scheme.lower() != "file":
        # Not actually a file URI; treat as path
        return Path(uri)
    # UNC paths have a netloc, local drives typically don't
    if parsed.netloc:
        loc = f"//{parsed.netloc}{parsed.path}"
    else:
        loc = parsed.path
    # Convert percent-encoding and URI path -> native path
    native = url2pathname(unquote(loc))
    return Path(native)


def _candidate_roots() -> Iterable[Path]:
    # 1) Current working directory
    yield Path.cwd()
    # 2) Parents of CWD (walk up a few levels)
    p = Path.cwd()
    for _ in range(6):
        p = p.parent
        if not p or str(p) == str(p.parent):
            break
        yield p
    # 3) Env-provided workspace roots
    for var in _ENV_ROOT_VARS:
        val = os.environ.get(var)
        if val:
            try:
                yield Path(os.path.expanduser(os.path.expandvars(val)))
            except Exception:
                continue


def resolve_user_path(input_path: str) -> Path:
    """Best-effort resolution of a user-provided path/URI to an absolute Path.

    Strategy:
    - Strip surrounding quotes and whitespace
    - If file:// URI, convert to local path
    - Expand ~ and %VAR%/$(VAR) env vars
    - If absolute and exists -> return
    - If relative -> try against candidate roots (cwd, parents, env roots)
    - If still not found -> return absolute path under cwd (may not exist)
    """
    if not input_path:
        return Path("")
    raw = _strip_quotes(input_path)
    # Handle file:// URIs
    if raw.lower().startswith("file:"):
        p = _from_file_uri(raw)
    else:
        # Expand env vars and user home
        expanded = os.path.expandvars(os.path.expanduser(raw))
        p = Path(expanded)

    # If it's absolute and exists, we're done
    try:
        if p.is_absolute() and p.exists():
            return p.resolve()
    except Exception:
        pass

    # If it's relative, try a set of candidate roots
    rel = p if not p.is_absolute() else Path(p.name) if not p.exists() else None
    if rel is not None:
        for root in _candidate_roots():
            candidate = (root / rel).resolve()
            try:
                if candidate.exists():
                    return candidate
            except Exception:
                continue

    # As a fallback, normalize to an absolute path under CWD (may not exist)
    try:
        return p.resolve() if p.is_absolute() else (Path.cwd() / p).resolve()
    except Exception:
        # Last resort: return as-is Path
        return p
