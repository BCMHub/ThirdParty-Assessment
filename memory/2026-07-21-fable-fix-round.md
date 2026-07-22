# Debug report — Fable fix round

- Symptom: `modEmail.Reconcile` did not compile under `Option Explicit`; build acceptance depended on a configured Outlook/MAPI runtime; custom-property reconciliation could trust a false-empty `Items.Restrict`; and an `Unresolved` event did not block a duplicate email intent.
- Root cause: `sendingAccount` was omitted from `Reconcile`'s local declarations; build and runtime shared a single side-effecting check path and the controlled reopen fired `Workbook_Open`; `SearchFolder` only used its manual scan after a raised error even though custom-property `Restrict` can return empty without error; and both the summary cache and integrity duplicate set omitted `Unresolved`.
- Fix: declared `sendingAccount`; threaded `structuralOnly` through `SelfTest`, `Check`, and `RunChecks`; guarded runtime environment checks and mutation/UI side effects; passed the flag at both build gates and disabled events during the controlled reopen; made the manual `UserProperties` iteration authoritative; and treated `Unresolved` as blocking in both preparation and structural integrity validation.
- Evidence: `python3 tests/vba_static_audit.py` passes over 34 modules and 188 procedures; all 34 modules contain `Option Explicit`; a negative control reports the removed `sendingAccount` as undeclared; targeted structural-only/GUID/Unresolved assertions pass; no compiled Office artifacts were introduced.
- Regression test: `tests/vba_static_audit.py`.
- Additional undeclared-variable findings: none beyond the reported `sendingAccount` defect.
- Related: the workspace is not a Git checkout, so recent-history/diff archaeology was unavailable. The current macOS host has no Excel/Outlook or PowerShell runtime, so the Windows COM build itself was not executed.
- Status: DONE_WITH_CONCERNS — source/static verification is complete; Windows Excel execution remains part of release evidence.
