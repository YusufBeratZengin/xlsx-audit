#!/usr/bin/env bash
# Single local gate: lint + format + types + tests + coverage threshold.
#
#   bash scripts/check.sh
#
# Exits 0 only when every step passes.  No network access is required: the test
# suite builds its own .xlsx fixtures with the standard library.
set -u -o pipefail

cd "$(dirname "${BASH_SOURCE[0]}")/.." || exit 1

PYTHON="${PYTHON:-python3}"
COVERAGE_MIN="${COVERAGE_MIN:-87}"
STATUS=0

step() {
    printf '\n\033[1m==> %s\033[0m\n' "$1"
}

fail() {
    printf '\033[31mFAILED: %s\033[0m\n' "$1" >&2
    STATUS=1
}

if ! command -v "$PYTHON" >/dev/null 2>&1; then
    echo "python3 is required but was not found on PATH" >&2
    exit 1
fi

step "interpreter"
"$PYTHON" -V || exit 1

step "ruff lint"
if ! "$PYTHON" -m ruff check . ; then fail "ruff lint"; fi

step "ruff format check"
if ! "$PYTHON" -m ruff format --check . ; then fail "ruff format"; fi

step "mypy (strict)"
if ! "$PYTHON" -m mypy ; then fail "mypy"; fi

step "pytest with coverage >= ${COVERAGE_MIN}%"
if ! "$PYTHON" -m pytest -q \
    --cov=xlsx_audit \
    --cov-report=term-missing \
    --cov-fail-under="${COVERAGE_MIN}" ; then
    fail "pytest / coverage"
fi

printf '\n'
if [ "$STATUS" -eq 0 ]; then
    printf '\033[32mAll checks passed.\033[0m\n'
else
    printf '\033[31mOne or more checks failed.\033[0m\n' >&2
fi
exit "$STATUS"
