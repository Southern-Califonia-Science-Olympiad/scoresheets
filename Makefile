# Scoresheet regression tests.
#
#   make                      run every event's suite
#   make test-Thermodynamics  run one event
#   make test ONLY=tiebreaks  run one scenario across every suite
#   make list                 show the suites that were discovered
#   make help                 usage
#
# Suites are discovered as */tests/test_scoresheet.py, so a new event folder
# is picked up automatically. Event directory names must not contain spaces.

PYTHON  ?= python3
SOFFICE ?= soffice

SUITES := $(sort $(wildcard */tests/test_scoresheet.py))
EVENTS := $(patsubst %/tests/test_scoresheet.py,%,$(SUITES))

# Optional: ONLY=<scenario> limits to one scenario, SHEET=<path> tests a
# different workbook than the one beside the suite.
ARGS := $(if $(SHEET),"$(SHEET)") $(if $(ONLY),--only $(ONLY))

.DEFAULT_GOAL := test
# test-% is deliberately not .PHONY: make skips implicit-rule search for phony
# targets, which would stop the pattern rule below from ever matching.
.PHONY: all test list help clean check

all: test

help:
	@echo "Scoresheet test targets:"
	@echo "  make                      run every event's suite"
	@echo "  make test-<Event>         run one event (e.g. test-Thermodynamics)"
	@echo "  make list                 list discovered suites"
	@echo "  make clean                remove __pycache__ directories"
	@echo ""
	@echo "Variables:"
	@echo "  ONLY=<scenario>           run a single scenario"
	@echo "  SHEET=<path>              test a specific workbook"
	@echo "  PYTHON=<exe>              python interpreter (default: python3)"

list:
	@if [ -z "$(SUITES)" ]; then \
	  echo "no suites found (looked for */tests/test_scoresheet.py)"; \
	else \
	  for s in $(SUITES); do echo "  $$s"; done; \
	fi

check:
	@command -v $(SOFFICE) >/dev/null 2>&1 || { \
	  echo "error: '$(SOFFICE)' not found on PATH."; \
	  echo "LibreOffice is the calculation engine for these tests."; \
	  exit 1; \
	}

test: check
	@if [ -z "$(SUITES)" ]; then \
	  echo "no suites found (looked for */tests/test_scoresheet.py)"; \
	  exit 1; \
	fi
	@fail=0; \
	for s in $(SUITES); do \
	  echo "=== $$s ==="; \
	  $(PYTHON) "$$s" $(ARGS) || fail=1; \
	  echo ""; \
	done; \
	if [ $$fail -ne 0 ]; then \
	  echo "FAILED: one or more suites reported failures"; \
	fi; \
	exit $$fail

test-%: check
	@suite="$*/tests/test_scoresheet.py"; \
	if [ ! -f "$$suite" ]; then \
	  echo "no test suite at $$suite"; \
	  exit 1; \
	fi; \
	$(PYTHON) "$$suite" $(ARGS)

clean:
	@find . -type d -name __pycache__ -prune -exec rm -rf {} + 2>/dev/null || true
	@echo "removed __pycache__ directories"
