#!/usr/bin/env python3
"""Static release checks for the source-exported VBA project.

This is intentionally stricter than a procedure-balance check: it builds module
and procedure scopes, then rejects executable identifiers that are not declared
locally, declared at module/project scope, exposed by a form, or part of the VBA/
Office host vocabulary used by this project.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
VBA_FILES = sorted([*SRC.glob("*.bas"), *SRC.glob("*.cls"), *SRC.glob("*.frm")])

IDENTIFIER = re.compile(r"[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?")
PROC_HEADER = re.compile(
    r"^(?:(?:Public|Private|Friend|Static)\s+)?"
    r"(?:(Sub|Function)|Property\s+(Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)\s*(.*)$",
    re.IGNORECASE,
)
PROC_END = re.compile(r"^End\s+(?:Sub|Function|Property)$", re.IGNORECASE)
DECLARATION = re.compile(r"^(?:Dim|Static)\s+(.+)$", re.IGNORECASE)
MODULE_DECLARATION = re.compile(r"^(?:Public|Private|Friend|Global|Dim)\s+(.+)$", re.IGNORECASE)
CONST_DECLARATION = re.compile(r"^(?:(?:Public|Private|Friend)\s+)?Const\s+(.+)$", re.IGNORECASE)

KEYWORDS = {
    "and", "as", "byref", "byval", "call", "case", "close", "const", "declare",
    "dim", "do", "each", "else", "elseif", "empty", "end", "enum", "eqv", "erase",
    "error", "event", "exit", "explicit", "false", "for", "friend", "function", "get",
    "global", "gosub", "goto", "if", "imp", "implements", "in", "is", "let", "lib",
    "like", "lock", "loop", "lset", "me", "mod", "new", "next", "not", "nothing",
    "on", "open", "option", "optional", "or", "output", "paramarray", "preserve", "print", "private",
    "property", "ptrsafe", "public", "put", "raiseevent", "redim", "rem", "reset", "resume",
    "rset", "select", "set", "shared", "static", "step", "stop", "sub", "then", "to", "true",
    "type", "typeof", "until", "variant", "wend", "while", "with", "withevents", "write", "xor",
}

# Unqualified functions, host objects, intrinsic constants, and project reference
# types legitimately available without a declaration under Option Explicit.
HOST_SYMBOLS = {
    "activesheet", "activecell", "application", "array", "asc", "ascw", "cbool", "cbyte", "ccur", "cdate", "cdbl", "cdec",
    "cells", "cint", "clng", "clnglng", "clngptr", "csng", "cstr", "cvar", "choose", "chr", "chrw",
    "collection", "command", "createobject", "date", "dateadd", "datediff", "datepart", "dateserial",
    "datevalue", "day", "debug", "dir", "doevents", "environ", "err", "error", "filedatetime", "filelen",
    "fix", "format", "freefile", "getobject", "hex", "hour", "iif", "input", "inputbox", "instr",
    "instrrev", "int", "isarray", "isdate", "isempty", "iserror", "ismissing", "isnull", "isnumeric",
    "isobject", "join", "kill", "lbound", "lcase", "left", "len", "listcolumn", "listobject", "listrow",
    "load", "loc", "lof", "log", "ltrim", "mid", "minute", "mkdir", "month", "msgbox", "now", "oct",
    "range", "replace", "rgb", "right", "rnd", "round", "rtrim", "second", "seek", "sgn", "shell",
    "space", "split", "strcomp", "string", "strptr", "switch", "thisworkbook", "time", "timer", "timeserial",
    "timevalue", "trim", "typename", "ubound", "ucase", "unload", "val", "vartype", "weekday", "worksheet",
    "workbook", "year",
}


@dataclass
class Procedure:
    name: str
    start_line: int
    statements: list[tuple[int, str]] = field(default_factory=list)
    declared: set[str] = field(default_factory=set)
    labels: set[str] = field(default_factory=set)


@dataclass
class Module:
    path: Path
    code_start: int
    statements: list[tuple[int, str]]
    controls: set[str]
    declared: set[str] = field(default_factory=set)
    procedures: list[Procedure] = field(default_factory=list)


def strip_comment_and_strings(line: str) -> str:
    output: list[str] = []
    in_string = False
    index = 0
    while index < len(line):
        char = line[index]
        if char == '"':
            if in_string and index + 1 < len(line) and line[index + 1] == '"':
                output.extend("  ")
                index += 2
                continue
            in_string = not in_string
            output.append(" ")
        elif char == "'" and not in_string:
            break
        else:
            output.append(" " if in_string else char)
        index += 1
    return "".join(output)


def split_statements(line: str) -> list[str]:
    statements: list[str] = []
    start = 0
    in_string = False
    index = 0
    while index < len(line):
        char = line[index]
        if char == '"':
            if in_string and index + 1 < len(line) and line[index + 1] == '"':
                index += 2
                continue
            in_string = not in_string
        elif char == "'" and not in_string:
            break
        elif char == ":" and not in_string and (index + 1 >= len(line) or line[index + 1] != "="):
            statements.append(line[start:index])
            start = index + 1
        index += 1
    statements.append(line[start:index])
    return [statement.strip() for statement in statements if statement.strip()]


def logical_statements(lines: list[str], start: int) -> list[tuple[int, str]]:
    result: list[tuple[int, str]] = []
    pending = ""
    pending_line = start + 1
    for offset, raw in enumerate(lines[start:], start + 1):
        code = raw.rstrip()
        if not pending:
            pending_line = offset
        if re.search(r"\s_\s*(?:'.*)?$", code):
            code = re.sub(r"\s_\s*(?:'.*)?$", " ", code)
            pending += code
            continue
        pending += code
        for statement in split_statements(pending):
            result.append((pending_line, statement))
        pending = ""
    if pending:
        result.append((pending_line, pending.strip()))
    return result


def split_commas(text: str) -> list[str]:
    result: list[str] = []
    start = 0
    depth = 0
    in_string = False
    for index, char in enumerate(text):
        if char == '"':
            in_string = not in_string
        elif not in_string:
            if char == "(":
                depth += 1
            elif char == ")":
                depth = max(0, depth - 1)
            elif char == "," and depth == 0:
                result.append(text[start:index])
                start = index + 1
    result.append(text[start:])
    return result


def declared_names(text: str) -> set[str]:
    names: set[str] = set()
    for part in split_commas(text):
        cleaned = re.sub(r"\b(?:Optional|ByVal|ByRef|ParamArray|WithEvents)\b", " ", part, flags=re.I).strip()
        match = IDENTIFIER.match(cleaned)
        if match:
            names.add(match.group(0).rstrip("$%&!#@").lower())
    return names


def parameters(header_tail: str) -> set[str]:
    left = header_tail.find("(")
    if left < 0:
        return set()
    depth = 0
    right = -1
    for index in range(left, len(header_tail)):
        if header_tail[index] == "(":
            depth += 1
        elif header_tail[index] == ")":
            depth -= 1
            if depth == 0:
                right = index
                break
    return declared_names(header_tail[left + 1:right]) if right >= 0 else set()


def load_module(path: Path) -> Module:
    lines = path.read_text(encoding="cp1252").splitlines()
    option_lines = [index for index, line in enumerate(lines) if line.strip().lower() == "option explicit"]
    if not option_lines:
        raise AssertionError(f"{path.relative_to(ROOT)}: missing Option Explicit")
    start = option_lines[0]
    controls = {
        match.group(1).lower()
        for line in lines[:start]
        if (match := re.match(r"\s*Begin\s+(?:[A-Za-z0-9_.]+)\s+([A-Za-z_][A-Za-z0-9_]*)", line, re.I))
    }
    controls.add(path.stem.lower())
    return Module(path, start, logical_statements(lines, start), controls)


def build_scopes(module: Module) -> None:
    current: Procedure | None = None
    in_type = False
    for line_number, raw in module.statements:
        statement = strip_comment_and_strings(raw).strip()
        if not statement:
            continue
        if re.match(r"^(?:Public|Private)?\s*Type\b", statement, re.I):
            type_match = re.match(r"^(?:Public|Private)?\s*Type\s+([A-Za-z_][A-Za-z0-9_]*)", statement, re.I)
            if type_match:
                module.declared.add(type_match.group(1).lower())
            in_type = True
            continue
        if re.match(r"^End\s+Type$", statement, re.I):
            in_type = False
            continue
        if in_type:
            continue
        header = PROC_HEADER.match(statement)
        if header:
            if current is not None:
                raise AssertionError(f"{module.path.relative_to(ROOT)}:{line_number}: nested/unclosed procedure")
            current = Procedure(header.group(3), line_number)
            current.declared.update(parameters(header.group(4)))
            module.declared.add(current.name.lower())
            module.procedures.append(current)
            continue
        if PROC_END.match(statement):
            if current is None:
                raise AssertionError(f"{module.path.relative_to(ROOT)}:{line_number}: unmatched procedure end")
            current = None
            continue
        if current is None:
            const_match = CONST_DECLARATION.match(statement)
            module_match = MODULE_DECLARATION.match(statement)
            if const_match:
                module.declared.update(declared_names(const_match.group(1)))
            elif module_match and not re.match(r"^(?:Public|Private)\s+Declare\b", statement, re.I):
                module.declared.update(declared_names(module_match.group(1)))
            declare_match = re.match(r"^(?:Public|Private)\s+Declare(?:\s+PtrSafe)?\s+(?:Function|Sub)\s+([A-Za-z_][A-Za-z0-9_]*)", statement, re.I)
            if declare_match:
                module.declared.add(declare_match.group(1).lower())
        else:
            current.statements.append((line_number, raw))
            declaration = DECLARATION.match(statement)
            if declaration:
                current.declared.update(declared_names(declaration.group(1)))
            label_match = re.match(r"^([A-Za-z_][A-Za-z0-9_]*)$", statement)
            if label_match:
                current.labels.add(label_match.group(1).lower())
    if current is not None:
        raise AssertionError(f"{module.path.relative_to(ROOT)}:{current.start_line}: unclosed procedure {current.name}")


def is_host_symbol(name: str) -> bool:
    return name in HOST_SYMBOLS or name.startswith(("vb", "xl", "mso"))


def audit_identifiers(modules: list[Module]) -> list[str]:
    # A module's procedures/variables are available inside that module. Across
    # modules, this project uses explicit qualification (modFoo.Member), so only
    # component names are admitted globally; this avoids masking a typo with an
    # unrelated private declaration elsewhere in the project.
    project_symbols = {module.path.stem.lower() for module in modules}
    failures: list[str] = []
    for module in modules:
        available_module = project_symbols | module.declared | module.controls
        for procedure in module.procedures:
            available = available_module | procedure.declared | procedure.labels | {procedure.name.lower()}
            for line_number, raw in procedure.statements:
                code = strip_comment_and_strings(raw)
                stripped = code.strip()
                if not stripped or stripped.startswith("#"):
                    continue
                if DECLARATION.match(stripped) or CONST_DECLARATION.match(stripped):
                    continue
                if re.match(r"^(?:On\s+Error\s+GoTo|Resume)\s+[A-Za-z_][A-Za-z0-9_]*$", stripped, re.I):
                    continue
                for match in IDENTIFIER.finditer(code):
                    token = match.group(0).rstrip("$%&!#@").lower()
                    before = code[:match.start()].rstrip()
                    after = code[match.end():].lstrip()
                    if before.endswith((".", "!")) or (before.endswith("&") and token.startswith("h")) or after.startswith(":="):
                        continue
                    if token in KEYWORDS or token in available or is_host_symbol(token):
                        continue
                    failures.append(f"{module.path.relative_to(ROOT)}:{line_number}: undeclared identifier '{match.group(0)}' in {procedure.name}")
    return failures


def targeted_regressions() -> list[str]:
    failures: list[str] = []
    email = (SRC / "modEmail.bas").read_text(encoding="cp1252")
    env = (SRC / "modEnv.bas").read_text(encoding="cp1252")
    build = (ROOT / "Build_Windows.ps1").read_text(encoding="utf-8")
    reconcile = re.search(r"Public Sub Reconcile\b(.*?)End Sub", email, re.S | re.I)
    search_folder = re.search(r"Private Sub SearchFolder\b(.*?)End Sub", email, re.S | re.I)
    if not reconcile or not re.search(r"\bDim\s+sendingAccount\s+As\s+String\b", reconcile.group(1), re.I):
        failures.append("modEmail.Reconcile must declare sendingAccount locally")
    if not search_folder or "UserProperties(CORRELATION_PROPERTY)" not in search_folder.group(1):
        failures.append("modEmail.SearchFolder must scan UserProperties")
    if search_folder and ".Restrict(" in search_folder.group(1):
        failures.append("modEmail.SearchFolder must not trust Restrict for the custom correlation property")
    if not re.search(r'If stateValue = "Prepared" Or stateValue = "DraftCreated" Or stateValue = "Queued" Or stateValue = "Unresolved" Then mActiveEvents', email):
        failures.append("Unresolved events must block PrepareEmail through mActiveEvents")
    if not re.search(r'If stateValue = "Prepared" Or stateValue = "DraftCreated" Or stateValue = "Queued" Or stateValue = "Unresolved" Then\s+key = .*?activeEvents', env, re.S):
        failures.append("structural data integrity must reject duplicate pending/unresolved events")
    if "Optional ByVal structuralOnly As Boolean = False" not in env:
        failures.append("modEnv self-test entry points must expose structuralOnly")
    for guarded_side_effect in (
        "If Not structuralOnly Then modIdentity.ClearEnvironmentBlock",
        "If Not structuralOnly And Not valid Then modIdentity.BlockMutations",
        "If Not structuralOnly Then modDashboard.ApplyRoleBasedUI",
    ):
        if guarded_side_effect not in env:
            failures.append(f"structural-only self-test must guard side effect: {guarded_side_effect}")
    run_checks = re.search(r"Private Function RunChecks\b(.*?)End Function", env, re.S | re.I)
    if not run_checks or not re.search(r"If Not structuralOnly Then.*CheckClassicOutlook.*End If", run_checks.group(1), re.S | re.I) or "TestBackupFolder" not in run_checks.group(1):
        failures.append("runtime-only checks must remain inside the structuralOnly guard")
    if build.count('$false, $true)') != 2:
        failures.append("both Windows build gates must pass structuralOnly:=True")
    reopen = build.find('$reopened = $excel.Workbooks.Open')
    events_off = build.rfind('$excel.EnableEvents = $false', 0, reopen)
    if reopen < 0 or events_off < 0 or build[events_off:reopen].count("\n") > 2:
        failures.append("the controlled reopen must disable events so Workbook_Open cannot run the full runtime gate")
    return failures


def main() -> int:
    try:
        modules = [load_module(path) for path in VBA_FILES]
        for module in modules:
            build_scopes(module)
        failures = audit_identifiers(modules) + targeted_regressions()
    except AssertionError as error:
        failures = [str(error)]
    if failures:
        print("VBA STATIC AUDIT: FAIL")
        print("\n".join(f"- {failure}" for failure in failures))
        return 1
    procedure_count = sum(len(module.procedures) for module in modules)
    print(f"VBA STATIC AUDIT: PASS ({len(modules)} modules, {procedure_count} procedures)")
    print("Option Explicit scope audit: no undeclared executable identifiers found")
    print("Targeted regressions: structural-only gates, GUID scan, and Unresolved guard verified")
    return 0


if __name__ == "__main__":
    sys.exit(main())
