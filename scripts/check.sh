#!/usr/bin/env bash
# Single local gate: lint + format + types + tests + coverage threshold.
#
#   bash scripts/check.sh
#
# Exits 0 only when every step passes.  The script looks for an interpreter
# that already has the dev dependencies (ruff, mypy, pytest, pytest-cov,
# defusedxml) and, if there is none, bootstraps a repo-local .venv from the
# current interpreter.  That keeps `check.sh` working when the ambient
# `python3` is a bare runtime that never had the dev extras installed.
#
# Override the interpreter with PYTHON=/path/to/python.  No network access is
# needed by the test suite itself: it builds its own .xlsx fixtures with the
# standard library.
set -u -o pipefail

cd "$(dirname "${BASH_SOURCE[0]}")/.." || exit 1

COVERAGE_MIN="${COVERAGE_MIN:-87}"
STATUS=0

VENV_DIR="${VENV_DIR:-.venv}"

step() {
    printf '\n\033[1m==> %s\033[0m\n' "$1"
}

fail() {
    printf '\033[31mFAILED: %s\033[0m\n' "$1" >&2
    STATUS=1
}

# An interpreter is usable for this gate only when every tool the gate calls
# is importable from it.
has_dev_deps() {
    "$1" -c 'import defusedxml, mypy, pytest, pytest_cov, ruff' >/dev/null 2>&1
}

# Resolve a name or path to an executable path, or fail.
resolve() {
    if [ -x "$1" ]; then
        printf '%s\n' "$1"
        return 0
    fi
    type -P "$1" 2>/dev/null || return 1
}

venv_python() {
    if [ -x "$VENV_DIR/bin/python" ]; then
        printf '%s\n' "$VENV_DIR/bin/python"
    elif [ -x "$VENV_DIR/Scripts/python.exe" ]; then
        printf '%s\n' "$VENV_DIR/Scripts/python.exe"
    else
        return 1
    fi
}

# 1) Honour PYTHON when it is set and usable.
# 2) Reuse the repo-local venv from an earlier run.
# 3) Walk every python3/python on PATH (an activated venv can shadow a
#    complete interpreter with a bare one), then the usual absolute locations.
find_interpreter() {
    local candidate name
    for candidate in "${PYTHON:-}" "$(venv_python 2>/dev/null || true)"; do
        [ -n "$candidate" ] || continue
        candidate="$(resolve "$candidate" || true)"
        [ -n "$candidate" ] || continue
        if has_dev_deps "$candidate"; then
            printf '%s\n' "$candidate"
            return 0
        fi
    done
    for name in python3 python /usr/local/bin/python3 /usr/bin/python3 /opt/homebrew/bin/python3; do
        if [ "$name" = python3 ] || [ "$name" = python ]; then
            while IFS= read -r candidate; do
                [ -n "$candidate" ] || continue
                if has_dev_deps "$candidate"; then
                    printf '%s\n' "$candidate"
                    return 0
                fi
            done < <(type -ap "$name" 2>/dev/null)
        else
            candidate="$(resolve "$name" || true)"
            [ -n "$candidate" ] || continue
            if has_dev_deps "$candidate"; then
                printf '%s\n' "$candidate"
                return 0
            fi
        fi
    done
    return 1
}

# Interpreter used to create the venv: the requested one if it exists at all,
# otherwise plain python3/python.
base_interpreter() {
    local candidate
    for candidate in "${PYTHON:-}" python3 python; do
        [ -n "$candidate" ] || continue
        candidate="$(resolve "$candidate" || true)"
        [ -n "$candidate" ] && printf '%s\n' "$candidate" && return 0
    done
    return 1
}

PYTHON_CMD="$(find_interpreter || true)"

if [ -z "$PYTHON_CMD" ]; then
    BASE_PYTHON="$(base_interpreter || true)"
    if [ -z "$BASE_PYTHON" ]; then
        echo "python3 is required but was not found on PATH" >&2
        exit 1
    fi

    step "bootstrap ${VENV_DIR}"
    if [ -z "$(venv_python || true)" ]; then
        printf 'No interpreter with the dev dependencies found; creating %s with %s\n' \
            "$VENV_DIR" "$BASE_PYTHON"
        if ! "$BASE_PYTHON" -m venv "$VENV_DIR"; then
            echo "Could not create ${VENV_DIR}. Install the dev extras manually:" >&2
            echo "  $BASE_PYTHON -m pip install -e '.[dev]'" >&2
            exit 1
        fi
    fi

    PYTHON_CMD="$(venv_python || true)"
    if [ -z "$PYTHON_CMD" ] || ! has_dev_deps "$PYTHON_CMD"; then
        printf 'Installing project with dev extras into %s (one-off, needs network or pip cache)\n' \
            "$VENV_DIR"
        if ! "$PYTHON_CMD" -m pip install --disable-pip-version-check -q -e '.[dev]'; then
            echo "Could not install the dev dependencies." >&2
            echo "Run manually: $PYTHON_CMD -m pip install -e '.[dev]'" >&2
            echo "or point PYTHON at an interpreter that already has them." >&2
            exit 1
        fi
    fi

    if [ -z "$PYTHON_CMD" ] || ! has_dev_deps "$PYTHON_CMD"; then
        echo "Bootstrap finished but the dev dependencies are still missing." >&2
        exit 1
    fi
fi

step "interpreter"
printf 'Using %s\n' "$PYTHON_CMD"
"$PYTHON_CMD" -V || exit 1

step "ruff lint"
if ! "$PYTHON_CMD" -m ruff check . ; then fail "ruff lint"; fi

step "ruff format check"
if ! "$PYTHON_CMD" -m ruff format --check . ; then fail "ruff format"; fi

step "mypy (strict)"
if ! "$PYTHON_CMD" -m mypy ; then fail "mypy"; fi

step "pytest with coverage >= ${COVERAGE_MIN}%"
COVERAGE_ARGS=(
    --cov=xlsx_audit
    --cov-report=term-missing
    --cov-fail-under="${COVERAGE_MIN}"
)
# CI sets COVERAGE_XML=1 so the same gate also produces the artefact it uploads.
if [ -n "${COVERAGE_XML:-}" ]; then
    COVERAGE_ARGS+=(--cov-report=xml)
fi
if ! "$PYTHON_CMD" -m pytest -q "${COVERAGE_ARGS[@]}" ; then
    fail "pytest / coverage"
fi

printf '\n'
if [ "$STATUS" -eq 0 ]; then
    printf '\033[32mAll checks passed.\033[0m\n'
else
    printf '\033[31mOne or more checks failed.\033[0m\n' >&2
fi
exit "$STATUS"
