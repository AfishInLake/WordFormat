# Makefile

.PHONY: help install build server clean tests

PROJECT_ROOT := $(CURDIR)

## help: Show this help message
help:
	@echo "Usage: make [target]"
	@echo ""
	@echo "Targets:"
	@echo "  install  # Install the project in editable mode"
	@echo "  build    # Run the build script"
	@echo "  server   # Start the Uvicorn server"
	@echo "  clean    # Clean build artifacts"
	@echo "  tests    # Run the tests using pytest"

## install: Install the project in editable mode
install:
	@echo "Setting up development environment..."
	@if [ ! -d ".venv" ]; then \
		echo "Creating virtual environment..."; \
		if command -v uv >/dev/null 2>&1; then \
			uv venv; \
		elif command -v python3 >/dev/null 2>&1; then \
			python3 -m venv .venv; \
		elif command -v python >/dev/null 2>&1; then \
			python -m venv .venv; \
		else \
			echo "Python not found. Please install Python."; exit 1; \
		fi \
	fi
	@if command -v uv >/dev/null 2>&1; then \
		echo "Syncing dependencies with uv..."; \
		uv sync; \
	else \
		echo "uv not found, falling back to pip..."; \
		if [ -f "requirements.txt" ]; then \
			.venv/bin/python -m pip install -r requirements.txt 2>/dev/null || .venv/Scripts/python.exe -m pip install -r requirements.txt; \
		elif [ -f "pyproject.toml" ]; then \
			.venv/bin/python -m pip install -e . 2>/dev/null || .venv/Scripts/python.exe -m pip install -e .; \
		else \
			echo "No requirements.txt or pyproject.toml found."; \
		fi \
	fi
	@.venv/bin/python scripts/download_model.py 2>/dev/null || .venv/Scripts/python.exe scripts/download_model.py
	@echo "Development environment ready!"

## build: Run the build script (scripts/build.bat)
build: install
	@echo "Running build script..."
	@if [ "$$(uname -s)" = "Linux" ] || [ "$$(uname -s)" = "Darwin" ]; then \
		echo "Detected Unix-like system, using build.sh..."; \
		chmod +x scripts/build.sh; \
		scripts/build.sh "$(PROJECT_ROOT)"; \
	elif [ "$$OS" = "Windows_NT" ] || [ -n "$$WINDIR" ]; then \
		echo "Detected Windows, using build.bat..."; \
		scripts/build.bat "$(PROJECT_ROOT)"; \
	else \
		echo "Unsupported platform!"; exit 1; \
	fi

## server: Start the Uvicorn development server
server:
	uvicorn wordformat.api:app --host 0.0.0.0 --port 8000

## tests: Run the tests using pytest
tests:
	@echo "Running tests..."
	@pytest tests/ --cov=src --cov-report=term-missing --cov-fail-under=85

## clean: Clean build artifacts
clean:
	@echo "Cleaning build artifacts..."
	@if [ "$$(uname -s)" = "Linux" ] || [ "$$(uname -s)" = "Darwin" ]; then \
		rm -rf dist/ build/ output/ *.spec 2>/dev/null; \
		echo "Clean complete (Unix)."; \
	elif [ "$$OS" = "Windows_NT" ] || [ -n "$$WINDIR" ]; then \
		scripts/clean.bat; \
	else \
		echo "Unknown OS, skipping clean."; \
	fi
