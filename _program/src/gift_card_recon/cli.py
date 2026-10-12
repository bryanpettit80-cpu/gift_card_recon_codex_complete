"""One public command router; all reporting uses the validated services."""
from __future__ import annotations

import argparse
import sys


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="gift-card-recon", description="Validated gift-card reporting and read-only readiness.")
    parser.add_argument("command", choices=("weekly", "monthly", "readiness"), help="Workflow to run; use COMMAND --help for its options.")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if not args or args[0] in {"-h", "--help"}:
        build_parser().print_help()
        return 0
    if args[0].startswith("-"):
        raise SystemExit(
            "The legacy direct-total reporting command is retired. "
            "Use gift-card-recon weekly or gift-card-recon monthly with the validated source evidence. "
            "Use gift-card-recon readiness for a read-only assessment."
        )
    command = build_parser().parse_args(args[:1]).command
    if command == "weekly":
        from gift_card_recon.auto_run import main as run
    elif command == "monthly":
        from gift_card_recon.monthly_close_cli import main as run
    else:
        from gift_card_recon.coordinator import main as run
    return run(args[1:])


if __name__ == "__main__":
    raise SystemExit(main())
