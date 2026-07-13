# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

WordFormat is a Python CLI tool that checks and auto-corrects formatting in academic Word documents (.docx). It uses an ONNX BERT model to classify paragraphs (heading, abstract, body text, keywords, references, etc.), then validates and applies formatting rules defined in a YAML config file.

**Entry points**: `wordf` and `wordformat` commands both map to `wordformat.cli:main`.

## Git conventions

- **不要添加 `Co-Authored-By` 到 commit message 中。**

## Build & test commands

```bash
# Install dev environment (creates venv, installs deps, downloads ONNX model)
make install

# Run all tests (with coverage, must reach 87%)
make tests

# Run a single test file
pytest tests/rules/test_abstract.py -v

# Run a specific test
pytest tests/rules/test_abstract.py::TestAbstractTitleContentENBase -v

# Run tests matching a keyword
pytest tests/ -k "Abstract" -v

# Lint & format
ruff check src/
ruff format src/

# Run API server locally
make server
# or: wordf startapi

# Lint only (no tests)
make lint

# Build distributable package
make build

# Build Vue UI and copy into api/static
make build-ui

# Clean build artifacts
make clean

# Run pre-commit checks on all files
pre-commit run --all-files
```

## Install variants

```bash
pip install -e "."              # core (includes FastAPI server)
pip install -e ".[test]"        # core + pytest plugins
pip install -e ".[dev]"         # everything (test, pre-commit, ruff, pyinstaller)
```

## Pre-commit hooks

Pre-commit runs on `pre-commit` and `pre-push` stages, configured in `.pre-commit-config.yaml`:

- **sync-version** — mirrors `pyproject.toml` version into `_version.py`
- **end-of-file-fixer** / **trailing-whitespace** / **debug-statements** / **check-yaml** / **pretty-format-json** — standard fixers
- **pyupgrade** — auto-modernizes Python syntax (`--py3-plus`)
- **ruff** — lint + format (auto-fix)

## Ruff config (pyproject.toml)

- **Line length**: 108 (pycodestyle), 200 (docstrings)
- **Complexity**: max 10 (mccabe)
- **Quote style**: double, space indent
- **Per-file ignores**: `__init__.py` (F401/F403/E501), `tree.py`/`cli.py` (T201 print), `body.py`/`heading.py`/`numbering.py`/`_text.py` (C901 complexity)

## Environment variables

| Variable | Purpose | Default |
|----------|---------|---------|
| `WORDFORMAT_BASE_DIR` | Override working directory | auto-detected from project root |
| `WORDFORMAT_API_KEY` | API key for model service | `""` |
| `WORDFORMAT_MODEL` | Model identifier | `""` |
| `WORDFORMAT_MODEL_URL` | Model service URL | `""` |
| `BATCH_SIZE` | ONNX inference batch size | `64` |
| `HOST` / `PORT` | API server bind | `127.0.0.1` / `8000` |

## Architecture: core data flow

The tool operates in two phases that are intentionally separated so users can inspect and manually adjust the intermediate JSON before applying formatting:

```
wordf gj (generate JSON)          wordf cf / wordf af (check/apply format)
─────────────────────          ───────────────────────────────────
.docx → ONNX classify →        JSON tree → match paragraphs by
        flat JSON per para               position order (zip) →
                                          per-node style check/apply
```

**Phase 1 — Classification** (`classify/tag.py` → `base.py`):
- `DocxBase.parse()` loads a .docx, iterates paragraphs, batches them through ONNX BERT inference (`agent/onnx_infer.py`), and returns a flat list of `{category, text, score, comment}` dicts saved as JSON.

**Phase 2 — Tree building** (`structure/`):
- `DocumentBuilder.build_from_json()` feeds the flat JSON list into `DocumentTreeBuilder.build_tree()`, which creates a hierarchical `FormatNode` tree. `LEVEL_MAP` in `structure/settings.py` maps category strings to numeric levels; `CATEGORY_TO_CLASS` is auto-built from `@register` decorators on `FormatNode` subclasses.
- `node_factory.create_node()` instantiates the right `FormatNode` subclass and calls `load_config()`, which walks the YAML dict along `NODE_TYPE` path, merges with class `DEFAULTS`, and stores a `DotDict`.

**Phase 3 — Matching & formatting** (`pipeline/stages.py`):
- Each paragraph in the document is matched to a tree node by position order: `_flatten_tree_nodes()` converts the tree to a flat DFS list, then `zip()` pairs nodes with `document.paragraphs` by index.
- Before formatting, `node.apply_replace(doc)` checks for a `replace` field in the JSON value dict; if present, it substitutes the paragraph's run text with the replacement string.
- The tree is also mutated: `promote_bodytext_in_subtrees_of_type()` replaces generic `body_text` nodes under specific parents (e.g. `AbstractTitleCN`) with typed content nodes (e.g. `AbstractContentCN`).
- `apply_format_check_to_all_nodes()` recursively traverses the tree. For each node it calls `node.check_format(doc)` or `node.apply_format(doc)`, which delegate to `node._base(doc, p, r)` — the boolean flags control whether paragraph styles (`p`) and run styles (`r`) are checked (diffed) or applied (written).

**Phase 4 — Numbering** (apply mode only, `numbering.py`):
- `process_heading_numbering()` strips manual heading numbers from run text and applies Word auto-numbering definitions.

## Key directories & files

| Path | Purpose |
|------|---------|
| `src/wordformat/cli.py` | CLI entry point (`gj`/`cf`/`af`/`tree`/`config`/`startapi` subcommands) |
| `src/wordformat/classify/tag.py` | Classification entry point: loads doc, calls `DocxBase`, returns JSON |
| `src/wordformat/base.py` | `DocxBase`: docx loading + ONNX batch inference |
| `src/wordformat/pipeline/stages.py` | Orchestrator: tree flattening, position-based paragraph matching, subtree promotion, text replace, calls check/apply on each node |
| `src/wordformat/pipeline/orchestrate.py` | Top-level orchestration: `auto_format_thesis_document()` |
| `src/wordformat/rules/node.py` | `FormatNode` base class and `TreeNode` |
| `src/wordformat/rules/abstract.py` | Abstract title/content/title-content nodes (CN + EN) |
| `src/wordformat/rules/heading.py` | Heading level 1/2/3 nodes (no longer overrides `load_config`) |
| `src/wordformat/rules/keywords.py` | Keywords nodes with tag detection and count validation |
| `src/wordformat/rules/body.py` | Body text node with punctuation checking |
| `src/wordformat/rules/caption.py` | Caption formatting and numbering check/apply |
| `src/wordformat/rules/references.py` | References title and entry nodes |
| `src/wordformat/rules/acknowledgement.py` | Acknowledgements title and content nodes |
| `src/wordformat/rules/object.py` | FigureImage and TableObject nodes |
| `src/wordformat/style/diff.py` | `CharacterStyle` and `ParagraphStyle` with diff/apply logic |
| `src/wordformat/style/defs.py` | Enums: `FontSize`, `FontColor`, `FontName`, `Alignment`, `LineSpacingRule`, etc. |
| `src/wordformat/config/models.py` | `NodeConfigRoot` (dict subclass with dot access + `collect_style_configs()`) |
| `src/wordformat/config/loader.py` | `LazyConfig` (normal class, not singleton) for YAML loading |
| `src/wordformat/config/dotdict.py` | `DotDict` (dict with dot access), `BASE_FORMAT` defaults, `deep_merge` |
| `src/wordformat/structure/registry.py` | `@register` decorator, `export_defaults()` — auto-registration of FormatNode subclasses |
| `src/wordformat/structure/settings.py` | `CATEGORY_TO_CLASS` and `LEVEL_MAP` (auto-built from registry) |
| `src/wordformat/numbering.py` | Heading auto-numbering (clear manual + apply Word numbering) |
| `src/wordformat/tree.py` | `Tree`, `TreeNode`, `Stack` data structures |
| `src/wordformat/api/__init__.py` | FastAPI app. Serves Vue SPA from `api/static/`. |
| `src/wordformat/agent/` | `onnx_infer.py` (ONNX model inference), `message.py` (messaging) |
| `src/wordformat/settings.py` | Global config: `BASE_DIR`, `SERVER_HOST`, model/env settings |
| `tests/` | 21 test files mirroring src/ structure, 955 tests, 87% coverage |
| `presets/` | Per-university preset YAML configs |
| `wordformat-skill/` | AI assistant skill definition for Claude Code integration |

## FormatNode subclass pattern

Each `FormatNode` subclass declares class-level attributes and optionally implements `_base(doc, p, r)`:

- `NODE_TYPE` — dot-separated YAML path (e.g. `"abstract.chinese.chinese_title"`)
- `NODE_LABEL` — Chinese label for comment annotations
- `DEFAULTS` — dict of default formatting values, merged with YAML overrides at load time
- `RULES` — optional dict of `rule_name → handler_method` for config-gated business rules
- `DEFAULT_RULES` — always-on rules (paragraph_style + character_style)

Use `@register("category_name", level=N)` decorator to auto-register in `CATEGORY_TO_CLASS` and `LEVEL_MAP`. Import the class in `rules/__init__.py`.

## Test conventions

- Tests are organized in `tests/` directory mirroring `src/wordformat/` structure (e.g. `tests/rules/` ↔ `src/wordformat/rules/`).
- Tests use `python-docx` `Document()` to create in-memory test docs. Patches via `unittest.mock.patch`.
- `conftest.py` provides shared fixtures: `doc`, `config_path`, `root_config`, `temp_docx`, `temp_json`, and `reset_config`/`reset_style_warning` autouse fixtures.

## Config model structure

Config is defined by `DEFAULTS` dict on each `FormatNode` subclass. YAML overrides are merged at load time. `NodeConfigRoot` (dict subclass in `config/models.py`) provides dot-notation access and `collect_style_configs()`. Use `wordf config -o config.yaml` to export a complete config template with all defaults. `BASE_FORMAT` in `config/dotdict.py` provides shared global format defaults.
