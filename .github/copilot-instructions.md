Short, actionable guidance for AI coding agents working in this repository.

Overview
- This repository (Semantic Workbench) is a multi-project monorepo composed of: a Python backend (`workbench-service`), a React frontend (`workbench-app`), many example assistants, and several `mcp-servers` (Python packages that implement the Model Context Protocol).
- Key Office-related project: `mcp-servers/mcp-server-office` (dev) and `mcp-servers/mcp-server-office-prod` (prod copy). These interact with native Microsoft Office via COM (pywin32) and therefore require a Windows host for reliable end-to-end testing.

What to know first (big picture)
- Architecture: frontend <-> workbench-service (Python REST API) <-> assistant services (MCP servers). Assistants and MCP servers communicate via the MCP protocol over stdio, SSE, or HTTP tunnels.
- Important directories and examples:
  - `workbench-service/` — backend service (see README for run/debug instructions).
  - `workbench-app/` — React UI (Node 20.x as documented in `workbench-app/run.sh`).
  - `mcp-servers/` — many MCP servers; see `mcp-server-office/` for Office tooling.
  - `libraries/python/` — shared Python libraries, especially `mcp-extensions` which is a local dependency referenced by several servers.

Developer workflows and commands (concrete)
- Preferred development environment: GitHub Codespaces / devcontainer (`.devcontainer/README.md`) — use it when possible to avoid local environment drift.
- General repo bootstrap (from repo root):
  - `make` — runs project-level installs (see top-level `Makefile`).
  - To run the full workbench locally: `tools/run-workbench-chatbot.ps1` (Windows) or `tools/run-workbench-chatbot.sh` (Unix).
- Running the Office MCP server (dev):
  - The server is designed to be run with `uv run -m mcp_server.start` (defaults to `stdio`). Example: `uv run -m mcp_server.start --transport sse --port 25566`.
  - For local development we use a shared venv at the repo root named `.venv_shared` and thin wrappers `mcp-servers/mcp-server-office/run_shared_venv.ps1` and `mcp-servers/mcp-server-office-prod/run_shared_venv.ps1`. Prefer invoking those wrappers on Windows to avoid `uv` per-folder venv churn.
  - Building a Windows standalone: `make package` in the `mcp-server-office` directory (Windows, PyInstaller required).

Project-specific conventions and patterns
- Python versions: many packages (including `mcp-server-office`) pin Python to `>=3.11, <3.13`. The shared venv uses Python 3.12. Avoid Python 3.13 for those packages.
- Local editable dependencies: packages in `libraries/python/` (notably `mcp-extensions`) are installed locally in editable mode for development. When editing servers that depend on them, ensure `pip install -e "libraries/python/mcp-extensions"` is run in the active venv.
- MCP server runtimes and wrappers:
  - Do not rely on per-folder `uv run` venv creation on Windows — it causes file lock/hardlink problems. Use the shared venv and the `run_shared_venv.ps1` wrappers instead.
  - MCP client configs are stored as `mcp.json` in workspace `.vscode` and user-level settings. There is a workspace override at `.vscode/mcp.json` that points `office` and `officeDev` at the wrapper scripts.

Integration points and external dependencies
- Office COM: `pywin32` — requires running on Windows with Office installed. Tests that exercise Word/PPT will only reliably run on a Windows host with Office available.
- Tunneling: `mcp-tunnel` and dev tunnels (ngrok/dev tunnel) are used for exposing MCP servers over the network. See `mcp-servers/mcp-server-office/README.md` for `--use-ngrok-tunnel` usage and `mcp-tunnel` examples.
- Cloud LLMs and credentials: many examples rely on OpenAI/Anthropic/Azure keys set in environment or `.env` files. Do not hardcode keys in code; follow `.env.example` patterns in server folders.

Testing and debugging
- Unit tests: many servers include pytest tests under `tests/`. Run targeted tests using the repository's Python venv: `.\.venv_shared\Scripts\pytest -q mcp-servers\mcp-server-office\tests\test_word_editor.py` (Windows PowerShell example).
- Manual E2E on Office: start the `mcp-server-office` wrapper, then call the MCP tools from a client (Workbench or `mcp-tunnel`) to exercise real Office automation. Only run one Office MCP server at a time to avoid COM contention.
- Common debugging issues:
  - Permission/lock errors when `uv` manages venvs on Windows. Fix: use the shared venv wrappers.
  - Python version mismatch errors when packaging or installing — check `pyproject.toml` `requires-python` ranges.

Code patterns to follow (concrete examples)
- MCP server entrypoint: `mcp-servers/<server>/mcp_server/start.py` — follow existing `mcp_server.start` structure when adding servers.
- Tools and prompts: Office server organizes logic under `mcp_server/markdown_edit/`, `mcp_server/prompts/` and `mcp_server/evals/`. Reuse those patterns for tool separation and unit-testable logic.
- Local helper libraries: prefer adding shared code into `libraries/python/` and updating `pyproject.toml` `tool.uv.sources` if the package should be installed as an editable local dependency.

Safety and operational notes for agents
- Avoid proposing changes that assume the presence of Office on CI or Linux — any Office-dependent features must be gated and tested only on Windows hosts.
- When modifying startup or packaging logic, include Windows-specific build notes (PyInstaller) and the `--use-ngrok-tunnel` flag examples.

File references (use these as precise anchors when suggesting edits)
- `mcp-servers/mcp-server-office/README.md` — Office server usage and examples.
- `mcp-servers/mcp-server-office/mcp_server/start.py` — server entrypoint.
- `libraries/python/mcp-extensions/` — local editable dependency required by Office server.
- `tools/sync-office-dev-to-prod.ps1` — script to mirror dev→prod server folders (useful when recommending sync or release flows).
- `.vscode/mcp.json` — workspace-level MCP client config pointing to wrapper scripts.

If you make changes
- Run unit tests for the affected server (`pytest`) and validate package installs in the shared venv (`.venv_shared`).
- When touching Office automation code, include a short E2E test description and note that it must be run on Windows with Office installed.

Questions for the maintainer
- Should user-level `mcp.json` edits be considered canonical, or should agents only update the workspace `.vscode/mcp.json`? (Currently both exist; workspace override is recommended.)

If something in this guide is unclear or you want additional sections (CI, release steps, or contributor checklist), tell me which and I will iterate.
