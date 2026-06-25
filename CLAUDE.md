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

# Run all tests (with coverage, must reach 85%)
make tests

# Run a single test file
pytest tests/test_rules.py -v

# Run a specific test
pytest tests/test_rules.py::TestAbstractTitleContentENBase -v

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
pip install -e "."              # core only
pip install -e ".[api]"         # core + FastAPI server
pip install -e ".[test]"        # core + pytest plugins
pip install -e ".[dev]"         # everything (api, test, pre-commit, ruff, pyinstaller)
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
- **Per-file ignores**: `__init__.py` (F401/F403/E501), `tree.py`/`cli.py` (T201 print), `body.py`/`heading.py`/`numbering.py`/`set_style.py`/`utils.py` (C901 complexity)

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

**Phase 1 — Classification** (`set_tag.py` → `base.py`):
- `DocxBase.parse()` loads a .docx, iterates paragraphs, batches them through ONNX BERT inference (`agent/onnx_infer.py`), and returns a flat list of `{category, text, score, comment}` dicts saved as JSON.

**Phase 2 — Tree building** (`word_structure/`):
- `DocumentBuilder.build_from_json()` feeds the flat JSON list into `DocumentTreeBuilder.build_tree()`, which creates a hierarchical `FormatNode` tree. `LEVEL_MAP` in `word_structure/settings.py` maps category strings to numeric levels; `CATEGORY_TO_CLASS` maps categories to `FormatNode` subclasses.
- `node_factory.create_node()` instantiates the right `FormatNode` subclass and calls `load_config()`, which resolves the node's `CONFIG_PATH` (e.g. `"abstract.english"`) against the Pydantic root config model (`NodeConfigRoot`).

**Phase 3 — Matching & formatting** (`set_style.py`):
- Each paragraph in the document is matched to a tree node by position order: `_flatten_tree_nodes()` converts the tree to a flat DFS list, then `zip()` pairs nodes with `document.paragraphs` by index.
- Before formatting, `node.apply_replace(doc)` checks for a `replace` field in the JSON value dict; if present, it substitutes the paragraph's run text with the replacement string.
- The tree is also mutated: `promote_bodytext_in_subtrees_of_type()` replaces generic `body_text` nodes under specific parents (e.g. `AbstractTitleCN`) with typed content nodes (e.g. `AbstractContentCN`).
- `apply_format_check_to_all_nodes()` recursively traverses the tree. For each node it calls `node.check_format(doc)` or `node.apply_format(doc)`, which delegate to `node._base(doc, p, r)` — the boolean flags control whether paragraph styles (`p`) and run styles (`r`) are checked (diffed) or applied (written).

**Phase 4 — Numbering** (apply mode only, `numbering.py`):
- `process_heading_numbering()` strips manual heading numbers from run text and applies Word auto-numbering definitions.

## Key directories & files

| Path | Purpose |
|------|---------|
| `src/wordformat/cli.py` | CLI entry point (`gj`/`cf`/`af`/`tree`/`startapi` subcommands) |
| `src/wordformat/set_style.py` | Orchestrator: tree flattening (`_flatten_tree_nodes`), position-based paragraph matching, subtree promotion, text replace, calls check/apply on each node |
| `src/wordformat/set_tag.py` | Classification entry point: loads doc, calls `DocxBase`, returns JSON |
| `src/wordformat/base.py` | `DocxBase`: docx loading + ONNX batch inference |
| `src/wordformat/rules/node.py` | `FormatNode` base class with `load_config`, `check_format`, `apply_format`, `add_comment` |
| `src/wordformat/rules/abstract.py` | Abstract title/content/title-content nodes (CN + EN) |
| `src/wordformat/rules/heading.py` | Heading level 1/2/3 nodes with custom config loading |
| `src/wordformat/rules/keywords.py` | Keywords nodes with tag detection and count validation |
| `src/wordformat/style/check_format.py` | `CharacterStyle` (run-level) and `ParagraphStyle` (para-level) with diff/apply logic |
| `src/wordformat/style/style_enum.py` | Enums: `FontSize`, `FontColor`, `FontName`, `Alignment`, `LineSpacingRule`, etc. |
| `src/wordformat/config/datamodel.py` | Pydantic v2 models for all config sections (`GlobalFormatConfig`, `AbstractConfig`, `HeadingsConfig`, `NumberingConfig`, etc.) |
| `src/wordformat/config/config.py` | `LazyConfig` singleton for YAML loading and caching |
| `src/wordformat/numbering.py` | Heading auto-numbering (clear manual + apply Word numbering) |
| `src/wordformat/tree.py` | `Tree`, `TreeNode`, `Stack` data structures |
| `src/wordformat/word_structure/` | Tree building from flat JSON (`DocumentBuilder`, `DocumentTreeBuilder`, `node_factory`, `settings`) |
| `src/wordformat/api/__init__.py` | FastAPI app (unusual: defined in `__init__.py`, not a separate module). 3 main endpoints + download. Serves Vue SPA from `api/static/`. |
| `src/wordformat/agent/` | `onnx_infer.py` (ONNX model inference), `message.py` (messaging) |
| `src/wordformat/settings.py` | Global config: `BASE_DIR`, `SERVER_HOST`, model/env settings |
| `tests/` | 5 test files with ~850 tests, ~93% coverage |
| `presets/` | Per-university preset YAML configs |
| `wordformat-skill/` | AI assistant skill definition for Claude Code integration |

## FormatNode subclass pattern

Each `FormatNode` subclass declares three class-level attributes and implements `_base(doc, p, r)`:

- `NODE_TYPE` — dot-separated path used by `TreeNode.load_config()` to extract dict config from the full YAML tree
- `CONFIG_MODEL` — Pydantic model class for type-safe config access via `self.pydantic_config`
- `CONFIG_PATH` — dot-separated path used by `FormatNode.load_config()` to resolve the Pydantic config object via `getattr`

When adding a new node type, register it in `word_structure/settings.py` (`CATEGORY_TO_CLASS` and `LEVEL_MAP`) and in `rules/__init__.py`.

## Test conventions

- Tests are organized by module: `test_core.py` (tree, stack, numbering, utils, DocxBase), `test_style.py` (style enums, CharacterStyle, ParagraphStyle), `test_rules.py` (all FormatNode subclasses), `test_integration.py` (config, cross-module, CLI, API), `test_coverage_boost.py` (edge case coverage).
- Known bugs are documented in tests with `"""BUG: ..."""` docstrings — the test asserts the current (buggy) behavior. When fixing, update the test to assert the correct behavior.
- Tests use `python-docx` `Document()` to create in-memory test docs rather than real .docx files. Patches are applied via `unittest.mock.patch` at the module level (e.g. `patch("wordformat.base.onnx_batch_infer", ...)`).
- `conftest.py` provides shared fixtures: `doc` (in-memory Document), a mock ONNX session, and config reset between tests.

## Config model structure

The YAML config is validated by `NodeConfigRoot` (in `datamodel.py`), which wraps:
- `abstract: AbstractConfig` → `chinese: AbstractChineseConfig`, `english: AbstractEnglishConfig`
- `headings: HeadingsConfig` → `level_1/level_2/level_3: HeadingLevelConfig`
- `body_text: GlobalFormatConfig`
- `references: GlobalFormatConfig`
- `captions: dict[str, GlobalFormatConfig]`
- `numbering: NumberingConfig`
- `style_checks_warning: WarningFieldConfig`
- `replace: dict[str, str]`

`GlobalFormatConfig` carries all glyph + paragraph format fields (font, size, color, bold, italic, alignment, spacing, indent, etc.). `AbstractChineseConfig` and `AbstractEnglishConfig` each hold a `title` and `content` sub-config (both `GlobalFormatConfig` subclasses), enabling `AbstractTitleContentCN`/`AbstractTitleContentEN` to apply different styles to title-runs vs content-runs within the same paragraph.
