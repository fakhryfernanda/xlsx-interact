#!/usr/bin/env python3
"""xlsx-interact — CLI entry point."""

import argparse
import sys

import commands as cmds
import formatter
import repl as repl_mod
from session import Session


def main():
    parser = argparse.ArgumentParser(description="xlsx-interact")
    parser.add_argument("file", help="Path to .xlsx file")
    parser.add_argument("command", nargs="?", help="Command (omit for REPL)")
    parser.add_argument("args", nargs=argparse.REMAINDER, help="Command arguments")
    parser.add_argument("-f", "--format", choices=["text", "json"], default="text")

    args = parser.parse_args()

    try:
        session = Session(args.file)
    except FileNotFoundError as e:
        sys.exit(str(e))

    if not args.command:
        repl_mod.run(session)
    else:
        fmt = "text"
        rest = []
        i = 0
        while i < len(args.args):
            a = args.args[i]
            if a in ("-f", "--format") and i + 1 < len(args.args):
                fmt = args.args[i + 1]
                i += 2
            elif a == "-f":
                fmt = "json"
                i += 1
            elif a.startswith("-f") and len(a) > 2 and a[2] != " ":
                fmt = a[2:]
                i += 1
            else:
                rest.append(a)
                i += 1
        line = f"{args.command} {' '.join(rest)}"
        cmd_name, kwargs = cmds.parse(line)
        entry = cmds.COMMANDS.get(cmd_name)
        if not entry:
            sys.exit(f"Unknown: {cmd_name}")
        try:
            data = entry["fn"](session, **kwargs)
        except Exception as e:
            sys.exit(f"Error: {e}")
        if isinstance(data, dict) and "error" in data:
            sys.exit(data["error"])
        print(formatter.render(cmd_name, data, fmt))

    session.close()


if __name__ == "__main__":
    main()
