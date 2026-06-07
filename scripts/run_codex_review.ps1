# Optional: uses Codex CLI to review the repo and run tests.
# Run this after install.ps1 succeeds.
$ErrorActionPreference = "Stop"

codex exec --sandbox workspace-write "Review this gift card reconciliation repo. Run the pytest suite. Fix only genuine bugs. Do not change business rules unless a test proves they are wrong. Preserve POS controls on the Reconciliation tab."
