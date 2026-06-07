# Source Inventory

## Local Context Package

- Layer path: `C:\CodexAgents\gift_card_recon_codex_complete\semantic_layers\restaurant-pos-semantic-layer`
- Created because `C:\CodexAgents\gift_card_recon_codex_complete\.agents\skills` and the Codex profile state folder were not writable in this session.
- This package can be moved into `.agents\skills\restaurant-pos-semantic-layer` later if workspace skill writes become available.

## Primary POS Export

- Source file: `C:\Users\Ruth's Chris GM\Dropbox\GETLinkedData-VB\Micros3700.7z`
- Archive size: 801,818 bytes
- Extracted size observed: 7,698,253 bytes
- Archive timestamp observed: 2026-06-05 06:21 local time
- Format: 7-Zip archive containing CSV-style `.TXT` exports with single-quoted text values and no headers.
- Restaurant identified in `RESTDEF.TXT`: Ruth's Chris, Virginia Beach, VA.

## Derived Working Copy

- Extracted inspection folder used in this workspace: `C:\CodexAgents\gift_card_recon_codex_complete\_inspect_micros3700`
- This folder was created for analysis only. Re-extract the archive if the source archive changes.

## Current Source Coverage

- POS exports: available through local files and archive extraction.
- Gift card reconciliation rules: available through `.agents\skills\gift-card-reconciliation\SKILL.md`.
- Data warehouse: deferred by user; use files, exports, Excel, CSV, or pasted query results for now.
- Team communication: deferred by user.
- BI/dashboard tools: deferred by user.
- Behavior analytics tools: deferred by user.
- Notebook tools: deferred by user.

## Sensitive Data

- Employee definition and labor files can include staff names, identifiers, job rates, pay, clock-in/out data, and other personnel fields.
- Avoid exposing employee-level detail unless the user specifically requests labor analysis.

