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
        line = f"{args.command} {' '.join(args.args)}"
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
        print(formatter.render(cmd_name, data, args.format))

    session.close()


if __name__ == "__main__":
    main()
