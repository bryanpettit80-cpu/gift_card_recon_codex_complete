# Repository Instructions

This project is the durable Gift Card Reconciliation checkout. Treat the tracked GitHub clone as the source of truth; extracted `-main` folders are not canonical.

## Workflow

- Keep weekly operator flows simple and repeatable.
- Preserve the operator-first Dropbox layout unless the requested change explicitly updates it.
- Do not commit raw merchant data, generated workbooks, or local output files unless the user explicitly asks for that artifact to be versioned.
- Prefer narrow, test-backed changes over broad refactors.

## Validation

Run this before committing changes:

```powershell
Push-Location .\_program
python -m pip install -e ".[dev]"
python -m pytest -q
Pop-Location
```

