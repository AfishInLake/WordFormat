# JSON 驱动结构编辑（表格增删 / 多段落条目 / UI 入口）实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让 agent（和人）只改 JSON 就能增删表格、用一条 JSON 条目驱动多个物理段落，并在 Vue UI 上得到插入/删除/改文本的入口。

**Architecture:** 上一轮已落地「段落身份」：`gj` 把 Word 原生的 `w14:paraId` 写回 docx 并导出为 `para_id`，`af --sync` 在建树前按该身份做结构对齐（改/插/删/重排段落），文本改写按字符落到原有 run。本轮把同一套身份机制延伸到**表格**（用表格首单元格段落的 `paraId` 作表格身份，不发明私有属性），把「一条 ↔ 一段」的既有不变量用**同步前展开**的方式兼容多段落写法，最后在 UI 上按身份暴露行级操作。

**Tech Stack:** Python 3.11+、python-docx 1.2.0（lxml）、pytest + coverage、loguru、ruff；Vue 3.5 + Vite 6（`WordFormatUI/`，无测试框架）。

**Spec:** 无独立 spec 文档；需求与边界来自上一轮段落身份改造的收尾说明与 `wordformat-skill/SKILL.md:136` 的既有契约，本文件即实现规格。

## Global Constraints

每个任务的隐含要求，逐条照抄自仓库现状：

- 覆盖率门禁：`pytest tests/ --cov=wordformat --cov-report=term-missing --cov-fail-under=87`（`make tests`），总覆盖率当前 91%，不得降到 87% 以下。
- lint/format 作用域**只有 `src/`**：`ruff check src/` 与 `ruff format src/` 必须全绿。**禁止**对 `tests/` 跑 `ruff format`/`ruff check --fix`——`tests/` 上有 245 处历史告警、11 个文件未格式化，改动它们会制造无关 diff（本计划新增的测试文件单独 `ruff format` 是可以的）。
- ruff：行宽 108（docstring 200）、双引号、复杂度 ≤10；日志统一 `from loguru import logger`（`base.py` 等模块已如此）。
- 段落/表格身份取值域：`[1, 0x7FFFFFFE]` 的 8 位大写十六进制，见 `src/wordformat/para_ids.py:26-28`；新值必须避开文档内已用值（含表格单元格段落），见 `_existing_ids`（`para_ids.py:41-43`）。
- 不得引入私有 XML 属性：只用 Word 自己维护的 `w14:paraId`，避免 Word 兼容性/校验风险（理由见 `para_ids.py:10-13`）。
- 保守降级原则：身份不完整时**绝不允许**按不存在的键删对象——沿用 `_identity_ready`（`sync_apply.py:176-185`）的思路，宁可退回旧的位置/文本行为并 `logger.warning`。
- 既有 1598 个测试必须保持全绿；`wordf gj/cf/af/jd/md` 的现有行为（不带 `para_id` 的旧 JSON）不得破坏。
- commit message：中文、`feat:`/`fix:`/`chore:` 前缀，**禁止 `Co-Authored-By`**（`AGENTS.md`）。只在被要求时提交。
- UI 改动必须在浏览器里真跑一遍并给截图（`make build-ui` → `wordf startapi`），不能只说「类型检查通过」。

## 范围与交付边界

三个里程碑**互相独立、各自可发布**，按顺序做；M5 只依赖「条目带 `para_id`/`table_id`、新条目不带」这个契约，不依赖 M3/M4 的实现细节，可拆给另一轮执行。

明确**不做**（YAGNI，需要时另立计划）：

- 表格**单元格内容**改写：`table_rows` 保持只读（要改内容走 `jd` 或在 Word 里改）。
- 「一个 `FormatNode` 持有多段」的新概念：多段落一律在同步前展开成多条条目。
- 嵌套表格、跨列合并单元格的表格身份特例。
- UI 拖拽排序、批量多选删除、UI 单测框架引入。

## 文件结构

| 文件 | 职责 | 本轮动作 |
|------|------|----------|
| `src/wordformat/para_ids.py` | 身份原语：读/补 `w14:paraId` | 修改：新增 `get_table_id`、`ensure_table_ids`、`_table_anchor_p`（表格身份与段落身份同源，放同一模块） |
| `src/wordformat/base.py` | gj：分类 + 身份落盘 + 表格导出 | 修改：`_annotate_paragraph_ids` 同时补表格身份；`_interleave_tables` 输出 `table_id` |
| `src/wordformat/pipeline/sync_apply.py` | af --sync 结构对齐 | 修改：表格按身份归位/删除/新增；`sync_skip` 判定从 stage 移入本模块；新增 `expand_multi_paragraph_items` |
| `src/wordformat/rules/object.py` | `TableObject` 建表 | 修改：`_try_insert_table` 建表后补表格身份 |
| `src/wordformat/pipeline/orchestrate.py` | 流水线装配 | 修改：`_validate_jd_categories` 拒绝 `paragraphs`（jd 不支持） |
| `tests/test_table_sync.py` | M3 表格身份端到端 | 新建 |
| `tests/test_multi_paragraph.py` | M4 多段落展开 | 新建 |
| `tests/test_para_ids.py` / `tests/test_gj_para_ids.py` | 身份原语 / gj 导出 | 修改：补表格身份用例 |
| `tests/test_sync_apply.py` | 同步行为 | 修改：`sync_skip` 迁移后的断言 |
| `WordFormatUI/src/components/DocTagChecker.vue` | 节点列表编排（`nodeData`、上传/校验/格式化调用） | 修改：结构脏检测 + sync 开关 + 行操作处理 |
| `WordFormatUI/src/components/NodeListPanel.vue` | 行渲染 | 修改：行内 删除/上方插入/上移/下移 |
| `WordFormatUI/src/components/NodeDetailPanel.vue` | 右侧详情 | 修改：文本编辑（写 `replace`）、身份展示 |
| `src/wordformat/api/static/**` | 已构建的前端产物（git 跟踪） | `make build-ui` 后提交 |
| `wordformat-skill/SKILL.md` | agent 可见契约 | 修改：`:43`、`:136`、`:158-160`、`:167`、`:246` 的表格/段落边界与字段表 |

**参考（无需改动）**：`src/wordformat/pipeline/stages.py:124-150`（`_is_detached_table` + 段落数一致的 zip 不变量）、`src/wordformat/pipeline/stages_md.py:78-83`（jd 侧 `exists`/`sync_skip` 不建占位段）、`src/wordformat/pipeline/orchestrate.py:92-108`（stage 顺序：`StructureSyncStage` 插在 `LoadDocxStage` 之后、`TreeBuildingStage` 之前）、`src/wordformat/api/__init__.py:244-260`（`/apply-format` 已有 `sync: bool = Form(False)`，M5 后端零改动）。

---

## M3：JSON 驱动表格增删

### Task 1: 表格身份原语

**Files:**
- Modify: `src/wordformat/para_ids.py`（在 `ensure_para_ids` 之后追加）
- Test: `tests/test_para_ids.py`（文件末尾追加 `class TestTableIds`）

**Interfaces:**
- Consumes: `W14_PARA_ID`、`get_para_id(p_el)`、`_existing_ids(body)`、`_mint_id(used)`（均已存在于 `para_ids.py:23-52`）
- Produces: `get_table_id(tbl_el) -> str | None`、`ensure_table_ids(document) -> int`、`_table_anchor_p(tbl_el)`（Task 2/3/4 全部依赖这两个公共函数）

- [ ] **Step 1: 写失败的测试**

在 `tests/test_para_ids.py` 末尾追加（`_id_doc` 等既有 helper 若命名不同，按该文件实际的建文档 helper 复用；下面用 `Document()` 现建）：

```python
class TestTableIds:
    def _doc_with_tables(self, n=2):
        from docx import Document

        doc = Document()
        doc.add_paragraph("表前")
        tables = []
        for k in range(n):
            t = doc.add_table(rows=2, cols=2)
            t.cell(0, 0).text = f"甲{k}"
            tables.append(t)
        doc.add_paragraph("表后")
        return doc, tables

    def test_get_table_id_reads_first_cell_paragraph(self):
        from wordformat.para_ids import ensure_table_ids, get_table_id

        doc, tables = self._doc_with_tables(1)
        assert get_table_id(tables[0]._tbl) is None
        assert ensure_table_ids(doc) == 1
        table_id = get_table_id(tables[0]._tbl)
        assert table_id and len(table_id) == 8
        # 身份就落在首单元格首段落上，没有第二个属性体系
        anchor = tables[0].rows[0].cells[0].paragraphs[0]
        assert get_para_id(anchor._p) == table_id

    def test_two_tables_get_distinct_ids_and_idempotent(self):
        from wordformat.para_ids import ensure_table_ids, get_table_id

        doc, tables = self._doc_with_tables(2)
        ensure_table_ids(doc)
        ids = {get_table_id(t._tbl) for t in tables}
        assert len(ids) == 2 and all(ids)
        assert ensure_table_ids(doc) == 0

    def test_nested_table_does_not_steal_outer_identity(self):
        from docx import Document

        from wordformat.para_ids import ensure_table_ids, get_table_id

        doc = Document()
        outer = doc.add_table(rows=1, cols=1)
        inner = outer.cell(0, 0).add_table(rows=1, cols=1)
        ensure_table_ids(doc)
        outer_id = get_table_id(outer._tbl)
        assert outer_id
        # 内层表格不是 body 直接子级，不参与 body 层表格计数与身份分配
        assert get_table_id(inner._tbl) == outer_id or get_table_id(inner._tbl)

    def test_table_ids_survive_save_round_trip(self, tmp_path):
        from docx import Document

        from wordformat.para_ids import ensure_table_ids, get_table_id

        doc, tables = self._doc_with_tables(1)
        ensure_table_ids(doc)
        before = get_table_id(tables[0]._tbl)
        path = tmp_path / "t.docx"
        doc.save(path)
        reloaded = Document(path)
        assert get_table_id(reloaded.tables[0]._tbl) == before
```

- [ ] **Step 2: 跑测试确认失败**

Run: `.venv/bin/python -m pytest tests/test_para_ids.py::TestTableIds -q`
Expected: `ImportError: cannot import name 'ensure_table_ids'`（前两条即失败）

- [ ] **Step 3: 实现最小代码**

在 `src/wordformat/para_ids.py` 末尾追加（模块 docstring 已说明「不另造私有属性」，此处沿用同一理由）：

```python
def _table_anchor_p(tbl_el):
    """表格身份锚点：``w:tbl`` 首行首单元格的首个 ``w:p``（可能为 None）。"""
    anchors = tbl_el.xpath("./w:tr[1]/w:tc[1]/w:p[1]")
    return anchors[0] if anchors else None


def get_table_id(tbl_el) -> str | None:
    """表格身份 = 首单元格段落的 ``w14:paraId``。

    OOXML 不给 ``w:tbl`` 定义 id，但单元格段落照样有 ``w14:paraId``，借它当身份
    可以复用 Word 的维护（用户在 Word 里编辑不会丢），不必引入私有属性。
    """
    anchor = _table_anchor_p(tbl_el)
    return get_para_id(anchor) if anchor is not None else None


def ensure_table_ids(document) -> int:
    """给 body 层每个表格的首单元格段落补齐身份，返回新写入数量。

    只处理 body 直接子级的 ``w:tbl``，与 gj 的表格导出范围（``base.py``
    ``_interleave_tables``）保持一致；嵌套表格不单独寻址。
    """
    body = document.element.body
    used = _existing_ids(body)
    minted = 0
    for tbl in body.iterchildren():
        if tbl.tag != qn("w:tbl"):
            continue
        anchor = _table_anchor_p(tbl)
        if anchor is None or get_para_id(anchor):
            continue
        anchor.set(W14_PARA_ID, _mint_id(used))
        minted += 1
    if minted:
        logger.debug(f"补齐表格身份：新写入 {minted} 个 w14:paraId")
    return minted
```

- [ ] **Step 4: 跑测试确认通过**

Run: `.venv/bin/python -m pytest tests/test_para_ids.py -q`
Expected: 全部 PASS。若 `test_nested_table_does_not_steal_outer_identity` 写成析取式显得含糊，改成断言「内层至少能取到某个 id 且 body 层只补了 1 个」并按实际输出定稿——**不要**为了让断言通过而改实现。

- [ ] **Step 5: 门禁 + 提交**

```bash
.venv/bin/python -m ruff check src/wordformat/para_ids.py && .venv/bin/python -m ruff format src/wordformat/para_ids.py
.venv/bin/python -m pytest tests/test_para_ids.py -q
git add src/wordformat/para_ids.py tests/test_para_ids.py
git commit -m "feat:表格身份沿用 w14:paraId（首单元格段落）"
```

---

### Task 2: gj 导出 table_id 并落盘

**Files:**
- Modify: `src/wordformat/base.py:156-172`（`_annotate_paragraph_ids`）、`src/wordformat/base.py:182-220`（`_interleave_tables`）
- Test: `tests/test_gj_para_ids.py`（追加到 `TestGjParaIdExport`）

**Interfaces:**
- Consumes: `ensure_table_ids`、`get_table_id`（Task 1）
- Produces: gj JSON 的 `table_object` 条目带 `table_id: str`（取不到身份时**不写该键**）；docx 落盘后每个 body 表格都可寻址

- [ ] **Step 1: 写失败的测试**

追加到 `tests/test_gj_para_ids.py` 的 `TestGjParaIdExport` 类内（该文件已有 `_make_docx`（支持 `("table", rows)`）与 `_parse(path, n_texts)`）：

```python
    def test_table_entries_carry_a_table_id(self, tmp_path):
        from wordformat.para_ids import get_table_id

        path = _make_docx(tmp_path, ["甲段", ("table", [["a", "b"]]), "乙段"])
        items = _parse(path, 2)
        table_item = next(it for it in items if it["category"] == "table_object")

        assert table_item["table_id"] == get_table_id(Document(path).tables[0]._tbl)
        assert table_item["exists"] is True

    def test_table_ids_are_persisted_to_disk(self, tmp_path):
        from wordformat.para_ids import get_table_id

        path = _make_docx(tmp_path, ["甲段", ("table", [["a"]])])
        _parse(path, 1)

        assert get_table_id(Document(path).tables[0]._tbl)
```

- [ ] **Step 2: 跑测试确认失败**

Run: `.venv/bin/python -m pytest tests/test_gj_para_ids.py -q`
Expected: 2 条新用例 FAIL（`KeyError: 'table_id'` / `get_table_id(...) is None`）

- [ ] **Step 3: 实现**

`src/wordformat/base.py` 第 176 行附近，把只补段落身份改成同时补表格身份（注意保留「写不回就显式告警」的既有分支）：

```python
    def _annotate_paragraph_ids(self) -> None:
        """补齐段落/表格身份并写回源文档。

        af 是另一个进程读同一个文件，身份只能靠 XML 落盘；写不回时按文本/位置降级，
        必须让用户看见原因，不能静默丢掉精确增删能力。
        """
        minted = ensure_para_ids(self.document) + ensure_table_ids(self.document)
        if not minted:
            return
        try:
            self.document.save(self.docx_file)
        except OSError as exc:
            logger.warning(
                f"段落身份无法写回 {self.docx_file}（{exc}）；"
                "af --sync 将退回按文本对齐，无法按段落精确增删"
            )
```

同文件 import 行改为：

```python
from wordformat.para_ids import ensure_para_ids, ensure_table_ids, get_para_id, get_table_id
```

`_interleave_tables` 里插表条目的 dict 字面量（`base.py:207-219`）加一行 `table_id`，取不到就不写键：

```python
        out = list(result)
        for k, (pos, tbl_el) in enumerate(tables):
            entry = {
                "category": "table_object",
                "paragraph": "",
                "score": 1.0,
                "comment": "原文档表格（内容只读，不参与格式检查）",
                "table_rows": self._extract_table_rows(tbl_el),
                "exists": True,
                "needs_review": False,
            }
            table_id = get_table_id(tbl_el)
            if table_id:
                entry["table_id"] = table_id
            out.insert(min(pos + k, len(out)), entry)
        return out
```

- [ ] **Step 4: 跑测试确认通过**

Run: `.venv/bin/python -m pytest tests/test_gj_para_ids.py tests/test_gj_tables.py tests/test_para_ids.py -q`
Expected: 全 PASS（`test_gj_tables.py` 断言导出顺序/`exists`，必须不受影响）

- [ ] **Step 5: 全量 + 提交**

```bash
.venv/bin/python -m pytest tests/ -q
git add src/wordformat/base.py tests/test_gj_para_ids.py
git commit -m "feat:gj 导出表格身份 table_id 并写回源文档"
```

---

### Task 3: 按身份归位与删除表格

**Files:**
- Modify: `src/wordformat/pipeline/sync_apply.py:43-53`（import）、`:208-242`（`_align_by_identity` 表格分支）、`:245-266`（`sync_document_structure` 日志）
- Test: `tests/test_table_sync.py`（新建，本任务先写删除/重排两类）

**Interfaces:**
- Consumes: `get_table_id`（Task 1）、gj 输出的 `table_id`（Task 2）
- Produces: `_table_identity_ready(items, table_elements) -> bool`；`_align_by_identity` 返回 `(n_edit, n_ins, n_del, n_tbl_ins, n_tbl_del)`；`sync_document_structure` 内部据此打日志（Task 4 复用同一返回值形状）

- [ ] **Step 1: 写失败的测试**

新建 `tests/test_table_sync.py`：

```python
#! /usr/bin/env python
# @Time    : 2026-10-28
# @Author  : afish
# @File    : test_table_sync.py
"""af --sync 的表格身份对齐：按 table_id 归位、删除、新增。"""

import pytest
from docx import Document

from wordformat.para_ids import ensure_para_ids, ensure_table_ids, get_table_id
from wordformat.pipeline.sync_apply import (
    _table_identity_ready,
    sync_document_structure,
)


def _doc_two_tables():
    doc = Document()
    doc.add_paragraph("段落A")
    t1 = doc.add_table(rows=1, cols=1)
    t1.cell(0, 0).text = "表一"
    t2 = doc.add_table(rows=1, cols=1)
    t2.cell(0, 0).text = "表二"
    doc.add_paragraph("段落B")
    ensure_para_ids(doc)
    ensure_table_ids(doc)
    return doc


def _body_tags(doc):
    from docx.oxml.ns import qn

    return [
        "tbl" if el.tag == qn("w:tbl") else "p"
        for el in doc.element.body.iterchildren()
        if el.tag in (qn("w:p"), qn("w:tbl"))
    ]


def _items(doc):
    """按当前文档生成 gj 形态的条目列表（含 para_id / table_id）。"""
    from wordformat.para_ids import get_para_id

    items = []
    para_iter = iter(doc.paragraphs)
    for el in doc.element.body.iterchildren():
        from docx.oxml.ns import qn

        if el.tag == qn("w:p"):
            para = next(para_iter)
            items.append(
                {
                    "category": "body_text",
                    "paragraph": para.text,
                    "score": 1.0,
                    "para_id": get_para_id(el),
                }
            )
        elif el.tag == qn("w:tbl"):
            items.append(
                {
                    "category": "table_object",
                    "paragraph": "",
                    "score": 1.0,
                    "exists": True,
                    "table_rows": [[c.text for c in r.cells] for r in el.tables[0].rows]
                    if False else [["x"]],
                    "table_id": get_table_id(el),
                }
            )
    return items


class TestTableIdentityReady:
    def test_ready_when_all_sides_have_ids(self):
        doc = _doc_two_tables()
        items = _items(doc)
        assert _table_identity_ready(items, list(doc.element.body.iterchildren())) is False

    def test_not_ready_without_table_id(self):
        doc = _doc_two_tables()
        items = [dict(it) for it in _items(doc)]
        for it in items:
            if it["category"] == "table_object":
                it.pop("table_id")
        tables = [
            el for el in doc.element.body.iterchildren() if get_table_id(el) is not None
        ]
        assert _table_identity_ready(items, tables) is False

    def test_no_table_items_but_paragraphs_have_ids_is_ready(self):
        # JSON 出自同一文档的 gj（段落带 para_id）却没有表格条目 → 认定表格被删
        doc = _doc_two_tables()
        items = [it for it in _items(doc) if it["category"] != "table_object"]
        tables = [
            el for el in doc.element.body.iterchildren() if get_table_id(el) is not None
        ]
        assert _table_identity_ready(items, tables) is True

    def test_no_table_items_and_no_para_id_is_not_ready(self):
        doc = _doc_two_tables()
        items = [{"category": "body_text", "paragraph": "段落A"}]
        tables = [
            el for el in doc.element.body.iterchildren() if get_table_id(el) is not None
        ]
        assert _table_identity_ready(items, tables) is False


class TestTableDeleteAndReorder:
    def test_reorder_tables_follow_json_order(self):
        doc = _doc_two_tables()
        items = _items(doc)
        first_id = items[1]["table_id"]
        tables = [it for it in items if it["category"] == "table_object"]
        items.remove(tables[0])
        items.insert(items.index(tables[1]) + 1, tables[0])

        sync_document_structure(doc, items)

        ids = [get_table_id(t._tbl) for t in doc.tables]
        assert ids[0] != first_id and ids[1] == first_id

    def test_dropping_a_table_entry_deletes_that_table(self):
        doc = _doc_two_tables()
        items = _items(doc)
        kept = items[1]["table_id"]
        dropped = items[2]["table_id"]
        items = [it for it in items if it.get("table_id") != dropped]

        sync_document_structure(doc, items)

        assert len(doc.tables) == 1
        assert get_table_id(doc.tables[0]._tbl) == kept
        assert _body_tags(doc) == ["p", "tbl", "p", "sectPr"] or "tbl" in _body_tags(doc)

    def test_without_table_id_keeps_positional_behavior_and_warns(self, caplog):
        doc = _doc_two_tables()
        items = [dict(it) for it in _items(doc)]
        for it in items:
            it.pop("table_id", None)
        before = len(doc.tables)

        sync_document_structure(doc, items)

        assert len(doc.tables) == before
```

- [ ] **Step 2: 跑测试确认失败**

Run: `.venv/bin/python -m pytest tests/test_table_sync.py -q`
Expected: `ImportError: cannot import name '_table_identity_ready'`

- [ ] **Step 3: 实现**

`sync_apply.py` import 段补：

```python
from wordformat.para_ids import (
    ensure_para_ids,
    get_para_id,
    get_table_id,
)
```

在 `_align_by_identity` 之前新增判定函数（保守降级原则的表格版）：

```python
def _table_identity_ready(items: list[dict], table_elements: List) -> bool:
    """表格可按身份归位/删除/新增的条件。

    JSON 完全没有表格条目时，只有段落条目普遍带 ``para_id``（说明这份 JSON 出自同一
    文档的 gj，见过全部表格）才认定「表格被删了」；否则保守地让表格原地不动——
    手写 JSON 忽略表格不等于要删表格。
    """
    table_items = [
        it for it in items if isinstance(it, dict) and it.get("category") == "table_object"
    ]
    if not all(it.get("table_id") for it in table_items):
        return False
    if not all(get_table_id(el) for el in table_elements):
        return False
    if not table_items and table_elements:
        return any(isinstance(it, dict) and it.get("para_id") for it in items)
    return True
```

`_align_by_identity` 表格分支改为按 id 认领，函数签名返回值扩成 5 元组：

```python
def _align_by_identity(document, items: list[dict]) -> tuple[int, int, int, int, int]:
    """按 ``para_id`` 对齐：JSON 为准，段落可增、删、改、移位；表格按 ``table_id`` 同样处理。"""
    para_by_id = {get_para_id(p._p): p for p in document.paragraphs}
    table_elements = [
        el for el in document.element.body.iterchildren() if el.tag == qn("w:tbl")
    ]
    if _table_identity_ready(items, table_elements):
        tbl_by_id: dict | None = {get_table_id(el): el for el in table_elements}
    else:
        tbl_by_id = None
    ordered: List = []
    n_edit = n_ins = n_del = n_tbl_ins = n_tbl_del = 0
    table_cursor = 0

    for item in items:
        if item.get("category") == "table_object":
            if tbl_by_id is not None:
                table_id = item.get("table_id")
                el = tbl_by_id.pop(table_id, None) if table_id else None
                if el is None:
                    # 新表格：先占一个空段，TableObject 节点建表后会把占位段删掉
                    ordered.append(document.add_paragraph("")._p)
                    n_tbl_ins += 1
                else:
                    ordered.append(el)
                continue
            # 旧行为：无表格身份时按 exists 顺序归位，不做删除
            if item.get("exists") and table_cursor < len(table_elements):
                ordered.append(table_elements[table_cursor])
                table_cursor += 1
            continue
        para = _take_existing(para_by_id, item)
        if para is None:
            para = document.add_paragraph(_item_text(item))
            n_ins += 1
        elif _set_paragraph_text(para, _item_text(item)):
            n_edit += 1
        ordered.append(para._p)

    n_del = len(para_by_id)
    _remove_paragraphs(list(para_by_id.values()))
    if tbl_by_id is not None:
        n_tbl_del = len(tbl_by_id)
        for el in tbl_by_id.values():
            el.getparent().remove(el)
    elif table_cursor != len(table_elements):
        logger.warning(
            f"文档中有 {len(table_elements)} 个表格，JSON 只描述到 {table_cursor} 个，"
            "未被引用的表格保持原位不动"
        )
    _reorder_body(document, ordered)
    ensure_para_ids(document)  # 新插入的段落也要有身份，否则下一轮同步断链
    return n_edit, n_ins, n_del, n_tbl_ins, n_tbl_del
```

`_align_by_text` 的返回补成同形状（表格在文本路径一律不支持）：

```python
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        ...
    return n_edit, n_ins, n_del, 0, 0
```

`sync_document_structure` 相应改为：

```python
    if _identity_ready(document, items):
        n_edit, n_ins, n_del, n_tbl_ins, n_tbl_del = _align_by_identity(document, items)
    else:
        if any(isinstance(it, dict) and it.get("para_id") for it in items):
            logger.warning(
                "JSON 带段落身份但文档侧不完整（未重新 gj 或在 Word 里被重建），"
                "已退回按文本对齐"
            )
        n_edit, n_ins, n_del, n_tbl_ins, n_tbl_del = _align_by_text(document, items)

    if not document.paragraphs:
        # Word 要求正文至少保留一个段落
        document.add_paragraph("")

    extra = f"，表格 +{n_tbl_ins} −{n_tbl_del}" if (n_tbl_ins or n_tbl_del) else ""
    logger.info(f"结构同步完成：改 {n_edit} 段，插 {n_ins} 段，删 {n_del} 段{extra}")
```

模块 docstring（`sync_apply.py:17` 那行「表格条目按 ``exists`` 顺序对应原有 ``w:tbl``」）同步改成：`- 表格条目按 ``table_id`` 认领原有 ``w:tbl``：删条目＝删该表格，不带 id 的条目＝在该位置新建表格；无 ``table_id``（旧 JSON）时退回按 ``exists`` 顺序归位且不删表。`

- [ ] **Step 4: 跑测试确认通过**

Run: `.venv/bin/python -m pytest tests/test_table_sync.py -q`
Expected: 全 PASS。写 Step 1 时有两处刻意留白的断言（`_body_tags` 的 `sectPr`、`test_reorder_...` 的 first_id）——先跑一次拿到真实值再把断言写死，**不要**放宽断言迁就实现。`_items()` 里 `if False else` 那行是草稿写法，实现前替换成真实的 `[[c.text for c in r.cells] for r in Table(el, doc._body).rows]`。

- [ ] **Step 5: 全量回归**

Run: `.venv/bin/python -m pytest tests/ -q`
Expected: 除新文件外全绿。若 `tests/test_sync_apply.py` 有断言 5 元组形状或表格警告文案失败，逐条确认是「形状变更」而非行为回退后再改断言。

- [ ] **Step 6: 提交**

```bash
git add src/wordformat/pipeline/sync_apply.py tests/test_table_sync.py
git commit -m "feat:af --sync 按 table_id 归位与删除表格"
```

---

### Task 4: 新增表格条目真正建表

**Files:**
- Modify: `src/wordformat/pipeline/sync_apply.py:269-305`（`StructureSyncStage.process` 的 `sync_skip` 循环）、`:245-266`（`sync_document_structure` 顶部）
- Modify: `src/wordformat/rules/object.py:300-326`（`_try_insert_table`）
- Test: `tests/test_table_sync.py`（追加 `TestTableInsert`）

**Interfaces:**
- Consumes: Task 3 的占位段落（`ordered.append(document.add_paragraph("")._p)`）
- Produces: sync 路径下 `table_object` 新条目 → 文档里真的多出表格，且新表格立即带 `table_id`；`sync_skip` 只在「不支持」时出现（`stages.py:130`、`stages_md.py:81`、`domains/thesis/__init__.py:329` 三个消费点语义不变）

- [ ] **Step 1: 写失败的测试**

追加到 `tests/test_table_sync.py`：

```python
class TestTableInsert:
    def test_new_table_item_creates_a_table_at_its_position(self):
        doc = _doc_two_tables()
        items = _items(doc)
        new_item = {
            "category": "table_object",
            "paragraph": "",
            "score": 1.0,
            "table_rows": [["新增1", "新增2"], ["新增3", "新增4"]],
        }
        items.insert(2, new_item)  # 表一之后、表二之前

        sync_document_structure(doc, items)

        # sync 只负责占位；建表发生在 FormattingExecutionStage
        assert _body_tags(doc) == ["p", "tbl", "p", "tbl", "p", "p", "sectPr"] or \
               _body_tags(doc)[:3] == ["p", "tbl", "p"]
        assert new_item not in [it for it in items if it.get("sync_skip")]

    def test_sync_skip_only_when_identity_unavailable(self):
        doc = _doc_two_tables()
        items = [dict(it) for it in _items(doc)]
        for it in items:
            it.pop("table_id", None)
            it.pop("para_id", None)
        new_item = {
            "category": "table_object",
            "paragraph": "",
            "table_rows": [["x"]],
        }
        items.append(new_item)

        sync_document_structure(doc, items + [{"category": "body_text", "paragraph": "无"}])

        assert new_item.get("sync_skip") is True

    def test_full_af_sync_pipeline_builds_the_new_table(self, tmp_path):
        from wordformat.pipeline.orchestrate import auto_format_thesis_document
        from wordformat.para_ids import get_para_id

        doc = _doc_two_tables()
        src = tmp_path / "t.docx"
        doc.save(src)
        items = _items(Document(src))
        items.insert(
            2,
            {
                "category": "table_object",
                "paragraph": "",
                "score": 1.0,
                "table_rows": [["甲", "乙"], ["丙", "丁"]],
            },
        )
        json_path = tmp_path / "nodes.json"
        json_path.write_text(json.dumps(items, ensure_ascii=False), encoding="utf-8")

        out = auto_format_thesis_document(
            jsonpath=str(json_path),
            docxpath=str(src),
            check=False,
            savepath=str(tmp_path / "out"),
            sync=True,
        )

        result = Document(out)
        assert len(result.tables) == 3
        assert result.tables[1].cell(0, 0).text == "甲"
        assert get_table_id(result.tables[1]._tbl)  # 新表格当场可寻址，下一轮能删它
```

（`import json` 已在文件头补上。）

- [ ] **Step 2: 跑测试确认失败**

Run: `.venv/bin/python -m pytest tests/test_table_sync.py::TestTableInsert -q`
Expected: 第 3 条失败——`StructureSyncStage` 仍把新表格标成 `sync_skip`，`len(result.tables) == 2`

- [ ] **Step 3: 实现**

`sync_apply.py`：把「新表格不支持就标 `sync_skip`」的判定从 stage 移到 `sync_document_structure` 顶部，因为只有那里同时看得见 document 与 items：

```python
def _mark_unsupported_new_tables(document, items: list[dict]) -> None:
    """身份不可用时，JSON 新增的表格条目仍按老行为忽略（建表走 jd）。"""
    table_elements = [
        el for el in document.element.body.iterchildren() if el.tag == qn("w:tbl")
    ]
    if _identity_ready(document, items) and _table_identity_ready(items, table_elements):
        return
    for item in items:
        if (
            isinstance(item, dict)
            and item.get("category") == "table_object"
            and not item.get("exists")
        ):
            item["sync_skip"] = True
            logger.warning(
                "本次同步无法按身份新增表格（缺 para_id/table_id 或在按文本对齐），"
                "该条目已忽略；如需增删表格请用 jd 重建或补齐身份后重试"
            )


def sync_document_structure(document, items: list[dict]) -> None:
    """以 JSON 条目列表为准，原地对 document 段落/表格做增/删/改/重排。

    Args:
        document: python-docx Document（会被就地修改）
        items: 编辑后的 JSON 扁平条目列表（category/paragraph/para_id/table_id/replace）
    """
    _mark_unsupported_new_tables(document, items)
    if _identity_ready(document, items):
        ...
```

`StructureSyncStage.process` 删掉原来的 `for item in items: ... sync_skip` 整段（`sync_apply.py:281-292`），只留：

```python
    def process(self, ctx):
        from wordformat.structure.document_builder import DocumentBuilder

        items = ctx.paragraphs or DocumentBuilder.load_paragraphs(ctx.json_path)
        sync_document_structure(ctx.document, items)
        if (
            not any(
                isinstance(it, dict) and it.get("category") != "table_object"
                for it in items
            )
            and ctx.document.paragraphs
        ):
            # 内容被删空时 sync 会补一个空段落（Word 要求），这里补占位条目，
            # 避免后续「节点数 ≠ 段落数」对齐检查失败。
            items.append({"category": "body_text", "paragraph": "", "score": 1.0})
        ctx.paragraphs = items
        return ctx
```

（原 `from wordformat.log_config import logger` 若变成未用 import 一并删除；`sync_document_structure` 内部已有 logger。）

`src/wordformat/rules/object.py` 的 `_try_insert_table` 末尾补身份，让新建表格下一轮即可被 `table_id` 寻址：

```python
        # 将表格从文档末尾移到当前段落后
        if self.paragraph is not None:
            anchor = self.paragraph._element
            anchor.addnext(table._element)
            # jd 生成的空占位段只用于定位；删除以免表格前多出一个空段
            if not self.paragraph.text.strip() and not self.paragraph._p.xpath(
                ".//w:drawing"
            ):
                anchor.getparent().remove(anchor)
        ensure_table_ids(doc)  # 新建表格立刻可被 table_id 寻址，否则要等下一轮 gj
```

同文件 import 段加 `from wordformat.para_ids import ensure_table_ids`。

- [ ] **Step 4: 跑测试确认通过**

Run: `.venv/bin/python -m pytest tests/test_table_sync.py tests/test_sync_apply.py tests/rules/test_object.py -q`
Expected: 全 PASS。`test_sync_apply.py:254-275`（`test_new_table_item_ignored_in_sync`）应仍通过——它的条目不带 `para_id`/`table_id`，走按文本对齐 → 依然 `sync_skip`。若它失败，说明 `_identity_ready` 判定被写错，回到 Step 3 修实现。

- [ ] **Step 5: 全量 + 门禁**

```bash
.venv/bin/python -m pytest tests/ -q --cov=src/wordformat --cov-report=term
.venv/bin/python -m ruff check src/ && .venv/bin/python -m ruff format --check src/
```
Expected: 全绿，TOTAL ≥87%（观察 `sync_apply.py` 覆盖是否因新分支下滑）

- [ ] **Step 6: 真机 CLI 验证（不依赖测试替身）**

```bash
mkdir -p temp/m3 && cp "temp/基于多智能体一致性的异构机器人协同巡检系统设计 3.docx" temp/m3/src.docx
.venv/bin/python -m wordformat.cli config -o temp/m3/config.yaml
.venv/bin/python -m wordformat.cli gj -d temp/m3/src.docx -o temp/m3/gj
```
再用一段 Python：读 gj JSON，删掉第 1 个 `table_object` 条目、在两个段落条目之间插入一条 `{"category": "table_object", "table_rows": [["A","B"],["C","D"]], "score": 1.0}`（不带 `table_id`/`exists`），跑 `wordf af --sync`，断言输出文档 `len(tables)` 与原文档一致（−1+1）、新表格首格 = `"A"`、被删表格的 `table_id` 不再出现在输出文档里。把观察到的数字写进提交说明。清理：`rm -rf temp/m3`。

- [ ] **Step 7: 提交**

```bash
git add src/wordformat/pipeline/sync_apply.py src/wordformat/rules/object.py tests/test_table_sync.py
git commit -m "feat:支持 JSON 新增表格条目在 af --sync 中建表"
```

---

### Task 5: M3 契约文档

**Files:**
- Modify: `wordformat-skill/SKILL.md:43`、`:136`、`:158-160`、`:167`、`:246`

**Interfaces:**
- Consumes: Task 2/3/4 的最终行为
- Produces: agent 可读的 `table_id` 语义（M5 的 UI 判定也以此为准）

- [ ] **Step 1: 改 `:43` 的「不要做」列**

把「已有条目的 `para_id` 不要改也不要清空（新增条目则不写该字段）…；原表格内容只读（不检查不改写），表格增删走 `jd` 或手动」改为：

```
已有条目的 `para_id` / `table_id` 不要改也不要清空（新增条目则不写该字段）；原表格内容只读（不检查不改写），但删条目＝删该表格、不带 `table_id` 的条目＝在该位置新建表格
```

- [ ] **Step 2: 改 `:136` 的 `af --sync` 段落，追加表格语义**

在该条末尾接一句（保持一行）：

```
表格按 `table_id`（gj 写入表格首单元格段落的 `w14:paraId`）认领：删条目＝删那个表格，不带 `table_id` 的 `table_object` 条目＝在 JSON 位置新建表格（`table_rows` 决定内容，`table_style` 选样式档案），条目顺序变了＝移动表格；无 `table_id` 或退回文本对齐时，新表格按老行为忽略、已有表格原地不动并告警。
```

- [ ] **Step 3: 补字段表（`:158-160` 区域）**

在 `exists` 行后插入一行：

```
| `table_id` | 表格身份（首单元格段落的 `w14:paraId`，`gj` 写入）：`af --sync` 按它认领/删除/移动表格；新增表格条目不写该字段 |
```

- [ ] **Step 4: 改 `:167` 与 `:246`**

`:167` 的「表格内容本身不参与格式检查与同步（只读）」改为「表格**单元格内容**不参与格式检查与同步（只读）；表格整体的增删移动由 `table_id` 驱动」。
`:246` 的「原文档表格内容级检查与改写（只读导出）」保留（内容级改写确实还没做），但在同段补一句已完成项：`表格的增/删/移动已可在 af --sync 下由 JSON 驱动。`

- [ ] **Step 5: 校验 + 提交**

```bash
.venv/bin/python -c "import yaml,glob;[yaml.safe_load(open(p)) for p in glob.glob('wordformat-skill/SKILL.md')]" 2>/dev/null; git --no-pager diff --stat wordformat-skill/SKILL.md
git add wordformat-skill/SKILL.md
git commit -m "docs:SKILL 补充表格按 table_id 增删移动的契约"
```

---

## M4：一条 JSON 条目驱动多个段落

### Task 6: 展开多段落条目

**Files:**
- Create: 函数落在 `src/wordformat/pipeline/sync_apply.py`
- Test: `tests/test_multi_paragraph.py`（新建）

**Interfaces:**
- Consumes: JSON 条目新可选字段 `paragraphs: list[str]`（≥2 项，与 `paragraph`/`replace` 互斥）
- Produces: `expand_multi_paragraph_items(items: list[dict]) -> list[dict]`——首条继承 `para_id`，其余为无身份兄弟条目插在紧随其后；`stages.py`/建树/zip 全程仍严格「一条 ↔ 一段」

- [ ] **Step 1: 写失败的测试**

新建 `tests/test_multi_paragraph.py`：

```python
#! /usr/bin/env python
# @Time    : 2026-10-28
# @Author  : afish
# @File    : test_multi_paragraph.py
"""多段落条目（paragraphs）在 af --sync 前展开为多条同身份条目。"""

import pytest

from wordformat.pipeline.sync_apply import expand_multi_paragraph_items


def _body(text, **extra):
    return {"category": "body_text", "paragraph": text, "score": 1.0, **extra}


class TestExpand:
    def test_without_paragraphs_returns_same_list(self):
        items = [_body("甲", para_id="AAAAAAAA"), _body("乙", para_id="BBBBBBBB")]
        assert expand_multi_paragraph_items(items) == items

    def test_two_texts_split_with_id_on_first_only(self):
        items = [_body("甲", para_id="AAAAAAAA", paragraphs=["第一段", "第二段"])]
        out = expand_multi_paragraph_items(items)

        assert [it["paragraph"] for it in out] == ["第一段", "第二段"]
        assert out[0]["para_id"] == "AAAAAAAA"
        assert "para_id" not in out[1]
        assert "paragraphs" not in out[0] and "paragraphs" not in out[1]
        assert out[1]["comment"] == "由多段落条目拆分"

    def test_category_and_score_are_inherited(self):
        items = [
            {
                "category": "heading_level_1",
                "paragraph": "旧",
                "score": 0.42,
                "para_id": "AAAAAAAA",
                "paragraphs": ["第一章 绪论", "本章小结"],
            }
        ]
        out = expand_multi_paragraph_items(items)
        assert [it["category"] for it in out] == ["heading_level_1"] * 2
        assert out[1]["score"] == 0.42

    def test_single_element_list_is_rejected(self):
        with pytest.raises(ValueError, match="至少 2 段"):
            expand_multi_paragraph_items([_body("甲", paragraphs=["只有一段"])])

    def test_conflict_with_replace_is_rejected(self):
        with pytest.raises(ValueError, match="replace"):
            expand_multi_paragraph_items(
                [_body("甲", paragraphs=["a", "b"], replace="覆写")]
            )

    def test_table_and_image_entries_are_rejected(self):
        with pytest.raises(ValueError, match="table_object"):
            expand_multi_paragraph_items(
                [{"category": "table_object", "paragraph": "", "paragraphs": ["a", "b"]}]
            )
        with pytest.raises(ValueError, match="figure_image"):
            expand_multi_paragraph_items(
                [{"category": "figure_image", "paragraph": "/a.png", "paragraphs": ["a", "b"]}]
            )

    def test_non_dict_items_pass_through(self):
        assert expand_multi_paragraph_items(["裸字符串"]) == ["裸字符串"]
```

- [ ] **Step 2: 跑测试确认失败**

Run: `.venv/bin/python -m pytest tests/test_multi_paragraph.py -q`
Expected: `ImportError: cannot import name 'expand_multi_paragraph_items'`

- [ ] **Step 3: 实现**

`sync_apply.py` 在 `_normalize` 附近追加：

```python
_MULTI_PARAGRAPH_UNSUPPORTED = ("table_object", "figure_image")


def expand_multi_paragraph_items(items: list[dict]) -> list[dict]:
    """把带 ``paragraphs`` 的条目拆成多条同类条目（首条继承 ``para_id``）。

    agent 写「一节的多个自然段」时很自然地想用一个条目装一个文本块，但建树/zip
    全程假定「一条 ↔ 一段」；在同步前展开一次即可复用全部既有机制，不必引入
    「节点持有多段」这条新概念。拆分出的兄弟条目不带 ``para_id``，于是
    ``_align_by_identity`` 按 JSON 位置把它们当新段落插入——正是需要的效果。
    """
    out: list[dict] = []
    for idx, item in enumerate(items, start=1):
        if not isinstance(item, dict):
            out.append(item)
            continue
        texts = item.get("paragraphs")
        if texts is None:
            out.append(item)
            continue
        if not isinstance(texts, list) or len(texts) < 2:
            raise ValueError(f"第 {idx} 条 paragraphs 必须是至少 2 段的字符串列表，单段请直接用 paragraph")
        category = item.get("category", "")
        if category in _MULTI_PARAGRAPH_UNSUPPORTED:
            raise ValueError(
                f"第 {idx} 条 {category} 不支持 paragraphs（{category} 的 paragraph 另有含义），请拆成多条条目"
            )
        if "replace" in item:
            raise ValueError(f"第 {idx} 条同时给了 paragraphs 与 replace，两种改写方式只能选一")
        first = {k: v for k, v in item.items() if k != "paragraphs"}
        first["paragraph"] = texts[0]
        out.append(first)
        for text in texts[1:]:
            sibling = {
                k: v
                for k, v in item.items()
                if k not in ("paragraphs", "paragraph", "para_id", "replace")
            }
            sibling["paragraph"] = text
            sibling["comment"] = "由多段落条目拆分"
            out.append(sibling)
    return out
```

- [ ] **Step 4: 跑测试确认通过**

Run: `.venv/bin/python -m pytest tests/test_multi_paragraph.py -q`
Expected: 全 PASS

- [ ] **Step 5: 提交**

```bash
git add src/wordformat/pipeline/sync_apply.py tests/test_multi_paragraph.py
git commit -m "feat:多段落条目展开为多条同身份条目"
```

---

### Task 7: 接线（sync 生效、jd 明确拒绝）

**Files:**
- Modify: `src/wordformat/pipeline/sync_apply.py:276-305`（`StructureSyncStage.process`）
- Modify: `src/wordformat/pipeline/orchestrate.py:210-238`（`_validate_jd_categories`）
- Test: `tests/test_multi_paragraph.py`（追加端到端类）、`tests/test_json_to_docx.py` 或既有 jd 测试文件补一条拒绝用例

**Interfaces:**
- Consumes: `expand_multi_paragraph_items`（Task 6）
- Produces: `af --sync` 支持 `paragraphs`；`jd`（含 `wordf jd` 与 `/json-to-docx`）遇到 `paragraphs` 直接报错而不是静默丢段

- [ ] **Step 1: 写失败的测试**

追加到 `tests/test_multi_paragraph.py`：

```python
import base64
import json

from docx import Document


def _make_doc(texts):
    doc = Document()
    for t in texts:
        doc.add_paragraph(t)
    return doc


def _item(category, text, **extra):
    d = {"category": category, "paragraph": text, "score": 1.0}
    d.update(extra)
    return d


class TestSyncEndToEnd:
    def _run(self, tmp_path, texts, items):
        from wordformat.pipeline.orchestrate import auto_format_thesis_document

        dp = tmp_path / "s.docx"
        _make_doc(texts).save(dp)
        jp = tmp_path / "s.json"
        jp.write_text(json.dumps(items, ensure_ascii=False), encoding="utf-8")
        return auto_format_thesis_document(
            jsonpath=str(jp),
            docxpath=str(dp),
            check=False,
            savepath=str(tmp_path / "out"),
            sync=True,
        )

    def test_one_entry_becomes_two_paragraphs_at_right_place(self, tmp_path):
        from wordformat.para_ids import ensure_para_ids, get_para_id

        doc = _make_doc(["甲", "乙"])
        ensure_para_ids(doc)
        ids = [get_para_id(p._p) for p in doc.paragraphs]
        items = [
            _item("body_text", "甲", para_id=ids[0], paragraphs=["甲改写", "插入的新段"]),
            _item("body_text", "乙", para_id=ids[1]),
        ]
        out = Document(self._run(tmp_path, ["甲", "乙"], items))

        assert [p.text for p in out.paragraphs] == ["甲改写", "插入的新段", "乙"]
        assert get_para_id(out.paragraphs[0]._p) == ids[0]

    def test_jd_rejects_paragraphs_instead_of_dropping_them(self, tmp_path):
        from wordformat.pipeline.orchestrate import json_to_docx

        jp = tmp_path / "nodes.json"
        jp.write_text(
            json.dumps(
                [_item("body_text", "甲", paragraphs=["一", "二"])], ensure_ascii=False
            ),
            encoding="utf-8",
        )
        with pytest.raises(ValueError, match="af --sync"):
            json_to_docx(str(jp), save_dir=str(tmp_path / "out"))
```

- [ ] **Step 2: 跑测试确认失败**

Run: `.venv/bin/python -m pytest tests/test_multi_paragraph.py::TestSyncEndToEnd -q`
Expected: 第 1 条失败（输出仍是 2 段，`paragraphs` 被忽略）；第 2 条失败（jd 静默按 `paragraph="甲"` 建了一段）

- [ ] **Step 3: 实现**

`StructureSyncStage.process` 第一行取值后立刻展开：

```python
        items = ctx.paragraphs or DocumentBuilder.load_paragraphs(ctx.json_path)
        # 「一条 ↔ 一段」是建树/zip 的地基，多段落写法在此展开一次，
        # 后续阶段（含非 sync 路径）看到的仍是扁平单段条目。
        items = expand_multi_paragraph_items(items)
```

（`ctx.paragraphs` 是列表时注意：展开返回**新列表**，后续 `items.append(...)` 占位逻辑照新列表操作，`ctx.paragraphs = items` 已在末尾。）

`_validate_jd_categories`（`orchestrate.py:210`）在既有类别检查后追加 `paragraphs` 拒绝——jd 的 `DocumentCreationStage` 一个节点只建一段，静默丢段不可接受：

```python
    multi = [idx for idx, item in enumerate(items, start=1)
             if isinstance(item, dict) and item.get("paragraphs")]
    if multi:
        raise ValueError(
            f"jd 不支持 paragraphs 多段落条目（第 {', '.join(map(str, multi[:5]))} 条），"
            "会静默丢段；请拆成多条条目，或改用 af --sync"
        )
```

- [ ] **Step 4: 跑测试确认通过**

Run: `.venv/bin/python -m pytest tests/test_multi_paragraph.py tests/test_sync_apply.py tests/test_jd.py -q`
Expected: 全 PASS

- [ ] **Step 5: 全量 + 提交**

```bash
.venv/bin/python -m pytest tests/ -q
git add src/wordformat/pipeline/sync_apply.py src/wordformat/pipeline/orchestrate.py tests/test_multi_paragraph.py
git commit -m "feat:af --sync 支持一条 JSON 驱动多个段落，jd 显式拒绝"
```

---

### Task 8: M4 契约文档

**Files:**
- Modify: `wordformat-skill/SKILL.md:136`、字段表 `:149-161`

- [ ] **Step 1: 字段表加一行**（接在 `replace` 行后）

```
| `paragraphs` | 一条描述多个连续自然段（`list[str]`，≥2 项）：仅 `af --sync` 支持，首段继承该条 `para_id`，其余按 JSON 位置新插；与 `paragraph`/`replace` 互斥，`table_object`/`figure_image` 不适用，`jd` 会报错 |
```

- [ ] **Step 2: `:136` 末尾追加一句**

```
一个条目想承载多个自然段时给 `paragraphs`（≥2 段字符串）：拆分出的后续段落当作「无身份新条目」按 JSON 位置插入，同样只认身份、不看文本相似度。
```

- [ ] **Step 3: 校验 + 提交**

```bash
git --no-pager diff --stat wordformat-skill/SKILL.md
git add wordformat-skill/SKILL.md
git commit -m "docs:SKILL 增加 paragraphs 多段落条目契约"
```

---

## M5：UI 的结构编辑入口（可独立一轮执行）

前端零测试框架（`WordFormatUI/package.json` 只有 `dev/build/preview`），所以每个任务的验收都是**浏览器实跑 + 截图**。后端 `/apply-format` 已有 `sync: bool = Form(False)`（`src/wordformat/api/__init__.py:256`），M5 不改 `src/`。

### Task 9: 结构脏检测 + 同步开关

**Files:**
- Modify: `WordFormatUI/src/components/DocTagChecker.vue`（状态区 `:132-141`、gj 回调 `:233-241`、`callApplyFormatApi` `:303-362`、import JSON 回调 `:434-464`）

**Interfaces:**
- Consumes: 条目上的 `para_id` / `table_id`（M2/M3 已产出）
- Produces: `syncMode: Ref<boolean>`、`structureDirty: Computed<boolean>`、`baselineIds: Ref<string[]>`；请求体在需要时含 `sync: 'true'`

- [ ] **Step 1: 加状态与快照**

`nodeData` 声明后追加：

```js
const syncMode = ref(false)
const baselineIds = ref([]) // gj/导入时的身份快照，用来识别用户改了结构

// 记录当前 JSON 的全部段落/表格身份
const takeBaseline = () => {
  baselineIds.value = nodeData.value
      .map(n => n.para_id || n.table_id)
      .filter(Boolean)
}
```

在 `callGenerateJsonApi` 成功分支里 `isFileLoaded.value = true` 之前调用 `takeBaseline()`；在 JSON 导入回调（`:434-464`）赋值 `nodeData` 后同样调用。

- [ ] **Step 2: 加脏检测计算属性**

```js
// 有身份的行消失＝删段落/删表格；出现无身份的行＝新增。两者都需要 --sync 才能落到文档上。
const structureDirty = computed(() => {
  const now = nodeData.value.map(n => n.para_id || n.table_id).filter(Boolean)
  if (now.length !== baselineIds.value.length) return true
  const baseline = new Set(baselineIds.value)
  if (now.some(id => !baseline.has(id))) return true
  return nodeData.value.some(
      n => n.category === 'table_object' ? !n.table_id && !n.exists && !n.sync_skip
          : !n.para_id && n.category !== 'table_object'
  )
})
```

- [ ] **Step 3: 请求体带 sync，并在未开开关时取得明确同意**

`callApplyFormatApi` 里 `formData.append('json_data', ...)`（`:341`）之后插入：

```js
  if (structureDirty.value && !syncMode.value) {
    const ok = window.confirm(
        'JSON 行数与文档结构已不一致（新增/删除了行）。\n' +
        '点「确定」按 --sync 模式把这些增删真正落到文档；\n' +
        '点「取消」返回，不改动文档。'
    )
    if (!ok) {
      isLoading.value = false
      return
    }
  }
  if (syncMode.value) formData.append('sync', 'true')
```

- [ ] **Step 4: 加开关 UI**

在 `FormatActionPanel.vue`（21 行，只放两个按钮）或 `DocTagChecker.vue` 的动作条内加一个复选框。若放 `FormatActionPanel`，改为 props/emit：

```html
<label class="sync-toggle" title="按 JSON 增删段落/表格（af --sync）">
  <input type="checkbox" v-model="syncMode" />
  <span>结构同步</span>
  <span v-if="structureDirty" class="sync-badge">检测到 {{ dirtyHint }}</span>
</label>
```

- [ ] **Step 5: 浏览器验证**

```bash
make build-ui && .venv/bin/python -m wordformat.cli startapi
```
打开 `http://127.0.0.1:8000/`，上传一个带表格的 docx → 生成节点 JSON → 确认「结构同步」可见、此时无徽标；删一行 → 徽标出现 → 点格式化 → 出现确认框。**截图存档并给用户看**（删行前后各一张）。

- [ ] **Step 6: 提交**

```bash
git add WordFormatUI/src src/wordformat/api/static
git commit -m "feat:UI 增加结构同步开关与结构改动提示"
```

### Task 10: 行级删除 / 上方插入 / 上下移动

**Files:**
- Modify: `WordFormatUI/src/components/NodeListPanel.vue:21-43`（行模板）、props/emits 声明
- Modify: `WordFormatUI/src/components/DocTagChecker.vue`（`selectNode` `:185-188` 附近新增行操作处理）

**Interfaces:**
- Consumes: `filteredNodeData`（显示序）与 `nodeData`（原始序）的双索引换算模式（沿用 `selectNode` 的 `nodeData.value.indexOf(node)`）
- Produces: emits `insert-above(index)` / `delete-node(index)` / `move-node({ index, delta })`，所有 index 均为 `filteredNodeData` 的显示索引

- [ ] **Step 1: NodeListPanel 加行内按钮**

行模板（`@click="$emit('select-node', index)"` 那个 `div`）内追加操作区：

```html
<div class="row-actions" @click.stop>
  <button type="button" title="在上方插入段落" @click="$emit('insert-above', index)">＋</button>
  <button type="button" title="上移" :disabled="index === 0" @click="$emit('move-node', { index, delta: -1 })">↑</button>
  <button type="button" title="下移" :disabled="index === rows.length - 1" @click="$emit('move-node', { index, delta: 1 })">↓</button>
  <button type="button" class="danger" title="删除（--sync 时会删掉对应段落/表格）" @click="$emit('delete-node', index)">✕</button>
</div>
```

在 `defineEmits` 里补 `'insert-above'`、`'delete-node'`、`'move-node'`（该组件当前只 emit `select-node`；若用 `emits: [...]` 选项式写法则加进数组）。

- [ ] **Step 2: DocTagChecker 处理，按对象引用换算原始索引**

```js
// 行操作收到的是过滤列表（可能受搜索/figure_image 隐藏影响）的索引，
// 必须先用节点对象换算回 nodeData 的真实下标，否则会改错行。
const rawIndexOf = (displayIndex) => {
  const node = filteredNodeData.value[displayIndex]
  return node ? nodeData.value.indexOf(node) : -1
}

const handleInsertAbove = (displayIndex) => {
  const raw = rawIndexOf(displayIndex)
  if (raw < 0) return
  // 新行刻意不带 para_id/table_id —— 后端据此判定为「新增」
  nodeData.value.splice(raw, 0, {
    category: 'body_text',
    paragraph: '（新段落，点击右侧编辑文本）',
    score: 1.0,
    comment: '手动插入',
    needs_review: false,
    id: Date.now()
  })
  selectedNodeIndex.value = raw
}

const handleDeleteNode = (displayIndex) => {
  const raw = rawIndexOf(displayIndex)
  if (raw < 0) return
  const node = nodeData.value[raw]
  const what = node.category === 'table_object'
      ? (node.exists ? '这个表格（--sync 时会真的删除文档中的表格）' : '这条新表格')
      : `这一段：${(node.paragraph || '').slice(0, 40)}`
  if (!window.confirm(`确认删除${what}？`)) return
  nodeData.value.splice(raw, 1)
  if (selectedNodeIndex.value === raw) selectedNodeIndex.value = -1
  else if (selectedNodeIndex.value > raw) selectedNodeIndex.value -= 1
}

const handleMoveNode = ({ index, delta }) => {
  const from = rawIndexOf(index)
  const to = from + delta
  if (from < 0 || to < 0 || to >= nodeData.value.length) return
  const [moved] = nodeData.value.splice(from, 1)
  nodeData.value.splice(to, 0, moved)
  selectedNodeIndex.value = to
}
```

把三个 handler 绑到 `<NodeListPanel>` 的对应事件上。

- [ ] **Step 3: 验证 figure_image 隐藏行没被误删**

`visibleNodeData`（`:151-153`）只在显示层滤掉 `figure_image`，`nodeData` 仍保留——这正是删除语义所依赖的：删除 `figure_image` 行在 UI 里不可能发生，图片段因此不会被误删。**浏览器里确认**：图片段不出现在列表、插入的新行在列表中位于所选行上方、删除后行序与 JSON 一致。

- [ ] **Step 4: 端到端 + 截图**

`make build-ui`、重启服务，走完整流程：上传 → 生成 JSON → 在某行上方插入 → 删除另一行 → 勾选「结构同步」→ 格式化 → 下载。用 Python 核对下载文件：
```bash
.venv/bin/python -c "
from docx import Document
d=Document('output/<下载文件名>')
print(len(d.paragraphs))
print([p.text for p in d.paragraphs[:3]])
"
```
断言新插入文本存在、被删段落文本消失。**截图**（列表操作前后、下载校验）给用户看。

- [ ] **Step 5: 提交**

```bash
git add WordFormatUI/src src/wordformat/api/static
git commit -m "feat:UI 支持段落行的插入/删除/上下移动"
```

### Task 11: 右侧文本编辑（写 replace，保留 para_id）

**Files:**
- Modify: `WordFormatUI/src/components/NodeDetailPanel.vue:38`（只读文本区）、`:64-77`（category select 附近）、`:110-112`（emits）
- Modify: `WordFormatUI/src/components/DocTagChecker.vue:191-197`（`handleCategoryChange` 旁新增 `handleReplaceChange`）

**Interfaces:**
- Consumes: 后端 `resolve_item_text`（`src/wordformat/utils/_text.py:9-19`）——`replace` 存在即以它为准，`""` 表示清空
- Produces: `update-replace(text)` emit；行对象新增/覆写 `replace` 字段，`paragraph` 保持原文不动；UI 显示「已改」标记与身份前缀

- [ ] **Step 1: 文本域替换只读展示**

```html
<div class="text-edit">
  <label class="field-label">段落文本</label>
  <textarea v-model="localText" rows="6" @change="$emit('update-replace', localText)"></textarea>
  <p class="original-hint">
    原文：{{ currentNode.paragraph || '（空段）' }}
    <span v-if="isEdited" class="edited-tag">已改</span>
  </p>
</div>
```

```js
const localText = ref('')
watch(
    () => props.node,
    (n) => { localText.value = n ? (n.replace ?? n.paragraph ?? '') : '' },
    { immediate: true }
)
const isEdited = computed(() => props.node && props.node.replace !== undefined
    && props.node.replace !== props.node.paragraph)
```

`emits` 补 `'update-replace'`。

- [ ] **Step 2: DocTagChecker 写回 replace**

```js
// 只写 replace：paragraph 保留 gj 的原文，供左右对照与 diff，
// 后端 resolve_item_text 会优先取 replace。
const handleReplaceChange = (text) => {
  if (!currentNode.value) return
  nodeData.value[selectedNodeIndex.value].replace = text
}
```

绑到 `<NodeDetailPanel @update-replace="handleReplaceChange">`。

- [ ] **Step 3: 展示身份，明确「新行无身份」**

详情头部加一行：

```html
<p class="id-line">
  身份 <code>{{ currentNode.para_id || currentNode.table_id || '（新增，同步时插入）' }}</code>
</p>
```

- [ ] **Step 4: 验证 run 级改写生效（浏览器 + 文档双向核对）**

选一段原文含加粗文字的段落，改中间几个字 → 勾选结构同步 → 格式化 → 下载。核对：新文本落地、原加粗 run 仍在。

```bash
.venv/bin/python -c "
from docx import Document
d = Document('output/<文件>')
for p in d.paragraphs:
    if '改后的关键词' in p.text:
        print([(r.text, r.bold) for r in p.runs]); break
"
```
截图（编辑面板 + 输出段落 run 结构）给用户看。若 UI 无法在浏览器里核对 run，就以 Python 输出为准并说明。

- [ ] **Step 5: 提交**

```bash
git add WordFormatUI/src src/wordformat/api/static
git commit -m "feat:UI 段落文本可编辑（写 replace，保留原文与身份）"
```

### Task 12: M5 文档与收尾门禁

- [ ] **Step 1: SKILL.md 补 UI 入口说明**

在 `:129-136` 的命令说明区补一条：

```
- UI（`wordf startapi`）：节点列表每行支持 上方插入/删除/上下移动，详情面板可编辑段落文本（写入 `replace`）；行数与文档结构不一致时会提示开启「结构同步」（等价 `af --sync`）。删除行即删除文档中该身份对应的段落/表格，勾选前想清楚。
```

- [ ] **Step 2: 全量门禁**

```bash
.venv/bin/python -m pytest tests/ -q --cov=src/wordformat --cov-fail-under=87
.venv/bin/python -m ruff check src/ && .venv/bin/python -m ruff format --check src/
```
Expected: 全绿、覆盖率 ≥87%。

- [ ] **Step 3: 清理临时产物**

`rm -rf temp/m3 temp/e2e* output/*--*.docx`（只删本轮生成的文件；`temp/`、`*.docx` 已被 gitignore，先 `ls` 确认别碰到用户的既有文件），`git status --short` 确认无意外内容。

- [ ] **Step 4: 提交**

```bash
git add wordformat-skill/SKILL.md
git commit -m "docs:记录 UI 结构编辑入口"
```

---

## 自检（写完后回看）

**1. 需求覆盖**
- JSON 增删表格 → Task 1–5（身份原语 / gj 导出 / 删除与重排 / 新增建表 / 契约）。
- 单条目多段落 → Task 6–8（展开 / 接线与 jd 拒绝 / 契约）。
- UI 增删入口 → Task 9–12（开关与脏检测 / 行操作 / 文本编辑 / 文档与门禁）。
- 未覆盖且刻意不做：表格单元格内容改写、嵌套表格寻址、UI 拖拽与批量选择、UI 测试框架——见「范围与交付边界」。

**2. 占位符扫描**：Task 3 Step 1 的 `_items()` 里 `if False else` 与两处「先跑一次拿真实值再写死」的断言是**已知待明确项**，执行时必须先跑再定稿，不许留成 `TODO`；其余步骤均带可运行代码。

**3. 类型一致性**：`_align_by_identity`/`_align_by_text` 统一返回 5 元组 `(n_edit, n_ins, n_del, n_tbl_ins, n_tbl_del)`，只有 `sync_document_structure` 解包；`get_table_id(tbl_el)` 收 XML 元素（不是 `Table` 对象），调用处一律传 `._tbl` 或 `body.iterchildren()` 的元素；`ensure_table_ids(document)` 收 Document（与 `ensure_para_ids` 同形）；`expand_multi_paragraph_items(items) -> list[dict]` 返回新列表，不改原列表。

---

## 执行记录（Task 1–4，2026-10-28）

**Task 3 与 Task 4 合并成一个提交**：占位段落与 `sync_skip` 判定必须同时落地。只合入 Task 3 时
（`StructureSyncStage` 仍把新表格条目一律标 `sync_skip`），`_align_by_identity` 已经给它落了一
个占位段，而占位段没有节点绑定，「段落数 == 条目数」被打破——2 段落 + 2 表格的小文档上直接抛
`ValueError: 当前文档(3 段)与节点 JSON(2 节点)不一致`。

**Task 3 Step 3 的 `_table_identity_ready` 判据改了**：计划里的 `all(it.get("table_id"))` 与
Task 4 自相矛盾（新建的表格条目按定义不带 `table_id`，会让身份模式整表退回按位置归位，
`test_new_table_item_creates_a_table_at_its_position` 必失败）。实现改以 `exists` 区分两种
「没有 table_id」：带 `exists` 却缺 id 只可能是表格身份上线前的旧 JSON，它指向既有表格，
删表不安全 → 退回保守模式；不带 `exists` 就是新建意图，没有 id 是正常的 → 身份模式照常生效。

**Task 4 的「建表」部分没有新增代码**：`_try_insert_table` 本就用 `table_rows` 建表并删占位段，
sync 路径接上占位段后端到端即通；只补了建表后的 `ensure_table_ids(doc)`，让新表格当场可寻址。

**真机验证结果**（temp 下已清理，数字记在这里）：原 1 个 body 表格 → 删该条目 + 在正文插一条
新表格条目 → 输出 1 个表格（内容 `A B / C D`）、被删 id `0A473B6C` 不再出现、段落数 384 不变、
空段 84→84、段落身份完整且唯一；对输出再跑 `gj`（新表格带 id `64843449`）→ 删该条目 →
日志「表格 +0 −1」、输出 0 个表格。`cf` 检查模式不受影响。

**新发现的边界（未修，需在 Task 5 的契约里写明）**：表格条目落在封面/声明等 `other` 类别子树时
不会建表——`DEFAULT_SKIP_CHILDREN_CATEGORIES`/`VOIDNODELIST` 让这些子树整体跳过格式化，占位段
原地留成一个空段。表格条目要放在正文类别的段落之间。
