# Release checklist

## Before tagging

- Update `version` in `pyproject.toml`.
- Run `python -m ruff check .`.
- Run `python -m ruff format --check .`.
- Run `python -m mypy`.
- Run `python -m pytest -q`.
- Run `python -m pytest -m performance -q`.
- Build with `python -m PyInstaller compare_tool.spec --noconfirm --clean`.
- Launch `dist\compare_tool.exe`.
- Compare the sample Excel files and confirm the output opens.
- Confirm the GUI title/header show the intended version.

## Tagging

```powershell
git tag v0.1.1
git push origin HEAD
git push origin v0.1.1
```

## After tagging

- Confirm the `Build Windows App` workflow succeeds.
- Download the `compare_tool-v<version>-windows` artifact.
- Smoke test the artifact on a clean folder.
- Publish release notes with supported formats and known limitations.
