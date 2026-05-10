import atexit
import os
import readline
import sys

import commands as cmds
import formatter

HISTFILE = os.path.expanduser("~/.cache/xlsx_history")


def _setup_history():
    os.makedirs(os.path.dirname(HISTFILE), exist_ok=True)
    try:
        readline.read_history_file(HISTFILE)
    except FileNotFoundError:
        pass
    readline.set_history_length(2000)
    atexit.register(lambda: readline.write_history_file(HISTFILE))


def run(session):
    _setup_history()
    print('xlsx-interact | type "help" for commands, "exit" to quit')
    switched = False
    while True:
        try:
            prompt = f"xlsx [{session.current_sheet}]> " if switched else "xlsx> "
            line = input(prompt).strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break

        if not line:
            continue

        cmd_name, kwargs = cmds.parse(line)

        if cmd_name == "exit":
            break

        entry = cmds.COMMANDS.get(cmd_name)
        if not entry:
            print()
            print(f"Unknown: {cmd_name}. Try 'help'")
            print()
            continue

        try:
            data = entry["fn"](session, **kwargs)
        except SystemExit:
            break
        except Exception as e:
            print(f"Error: {e}")
            print()
            continue

        if cmd_name == "sheet" and not (isinstance(data, dict) and "error" in data):
            switched = True

        if data is None:
            continue

        if isinstance(data, dict) and "error" in data:
            print()
            print(data["error"])
            print()
            continue

        print(formatter.render(cmd_name, data))
        print()
