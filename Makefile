# Makefile

.PHONY: help install build server clean tests

PROJECT_ROOT := $(CURDIR)

## help: Show this help message
help:
	@echo "Usage: make [target]"
	@echo ""
	@echo "Targets:"
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
	@echo "Development environment ready!"

## build: Run the build script (scripts/build.bat)
build:
	@echo "Running build script..."
	@scripts/build.bat "$(PROJECT_ROOT)";

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
	@if [ "$$OS" = "Windows_NT" ]; then \
		scripts/clean.bat; \
	else \
		echo "Cleaning build artifacts..."; \
		rm -rf dist build output 2>/dev/null; \
		rm -f *.spec 2>/dev/null; \
		echo "Clean complete."; \
	fi
