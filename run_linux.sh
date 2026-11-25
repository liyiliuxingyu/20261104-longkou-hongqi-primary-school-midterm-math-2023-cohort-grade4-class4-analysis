#!/usr/bin/env sh
set -eu
grade4_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
grade4_python=${CODEX_PRIMARY_RUNTIME_PYTHON:-python3}
exec "$grade4_python" "$grade4_dir/run.py" "$@"
