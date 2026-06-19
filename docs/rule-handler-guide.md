# 规则 Handler 开发指南

声明式规则引擎让开发者只需专注"判断是否有格式问题"，是否生效、何时生效全部由 YAML 配置决定。

## 快速开始

### 1. 定义规则配置模型

```python
# datamodel.py
from pydantic import BaseModel, Field

class BaseRuleConfig(BaseModel):
    """所有规则的基类，enabled 控制开关。"""
    enabled: bool = Field(default=True)

class MyCheckRule(BaseRuleConfig):
    """自定义校验规则的配置参数。"""
    threshold: int = Field(default=3, description="最小数量阈值")
    strict_mode: bool = Field(default=False)


class MyRulesConfig(BaseModel):
    """规则集合。"""
    my_check: MyCheckRule = Field(default_factory=MyCheckRule)
```

### 2. 挂到目标 Config 上

```python
class MyNodeConfig(GlobalFormatConfig):
    # ... 原有格式字段 ...
    rules: MyRulesConfig = Field(default_factory=MyRulesConfig)
```

### 3. 注册 Handler

```python
class MyNode(FormatNode[MyNodeConfig]):
    NODE_TYPE = "my_node"
    CONFIG_MODEL = MyNodeConfig
    CONFIG_PATH = "my_node"

    # 注册：规则名 → handler 方法名
    RULES = {"my_check": "_handle_my_check"}

    def _base(self, doc, p, r):
        """只做格式样式检查（ParagraphStyle / CharacterStyle）。"""
        ...

    def _handle_my_check(self, doc, rule_cfg: MyCheckRule, p: bool):
        """业务规则：只管判断，不关心 enabled。"""
        count = len(self.paragraph.text.split())
        if count < rule_cfg.threshold:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=f"数量不足（最少 {rule_cfg.threshold}，当前 {count}）",
            )
```

### 4. YAML 配置

```yaml
my_node:
  <<: *global_format
  rules:
    my_check:
      enabled: true       # false 则整个 handler 不执行
      threshold: 5
      strict_mode: true
```

---

## Handler 签名

```
handler(self, doc: Document, rule_cfg: RuleConfig, p: bool) -> None
```

| 参数 | 说明 |
|------|------|
| `doc` | python-docx 的 `Document` 对象，用于 `add_comment` |
| `rule_cfg` | 规则对应的 Pydantic 配置对象，`enabled` 已被框架检查过 |
| `p` | `True` = 检查模式（diff），`False` = 应用模式（apply） |

**进入 handler 时 `enabled` 已为 `True`**，不需要在 handler 里再判断。

---

## 示例

### 示例 1：纯检查型规则（关键词数量）

不需要区分 check/apply 模式，只做计数 + 报警。

```python
RULES = {"keyword_count": "_check_keyword_count"}

def _check_keyword_count(self, doc, rule_cfg: KeywordCountRule, p: bool = False):
    keyword_text = "".join(run.text for run in self.paragraph.runs)
    keyword_list = [
        k.strip() for k in re.split(r"[;,]", keyword_text) if k.strip()
    ]
    if len(keyword_list) < rule_cfg.count_min:
        self.add_comment(
            doc=doc, runs=self.paragraph.runs,
            text=f"数量不足（最少 {rule_cfg.count_min}，当前 {len(keyword_list)}）",
        )
```

### 示例 2：区分 check/apply 的规则（题注编号）

check 模式加批注，apply 模式直接改文本。

```python
RULES = {"caption_numbering": "_handle_caption_numbering"}

def _handle_caption_numbering(self, doc, rule_cfg: CaptionNumberingConfig, p: bool):
    prefix = self.pydantic_config.caption_prefix or "图"
    if p:
        _check_caption_numbering(self, doc, prefix, rule_cfg)   # 加批注
    else:
        _apply_caption_numbering(self, prefix, rule_cfg)         # 直接改
```

### 示例 3：注册多个规则

```python
RULES = {
    "keyword_count": "_check_keyword_count",
    "trailing_punctuation": "_check_trailing_punctuation",
}

def _check_keyword_count(self, doc, rule_cfg: KeywordCountRule, p: bool = False):
    ...

def _check_trailing_punctuation(self, doc, rule_cfg: TrailingPunctRule, p: bool = False):
    text = "".join(r.text for r in self.paragraph.runs).strip()
    if text and text[-1] in rule_cfg.forbidden_chars:
        self.add_comment(doc=doc, runs=self.paragraph.runs, text="末尾禁止标点")
```

---

## 执行流程

```
check_format(doc)
  ├── _base(doc, p=True, r=True)      ← 格式样式（自动）
  └── _run_rules(doc, p=True)         ← 业务规则（自动）
        ├── 读 config.rules
        ├── 比对 RULES ↔ config（不匹配则 warning）
        └── 遍历 RULES：
              enabled=True  → handler(doc, rule_cfg, p)
              enabled=False → 跳过
```

## 注意事项

- **`_base()` 只放格式样式检查**（`ParagraphStyle` / `CharacterStyle`），不要在里面直接写业务判断
- **Handler 里不要检查 `enabled`**——框架已过滤，能进来说明已启用
- **`p` 参数默认值设为 `False`**（`p: bool = False`），兼容不需要区分模式的 handler
- **Handler 不应有返回值**——结果全部通过 `self.add_comment()` 上报
- **新增配置模型记得在 `LEAF_TYPES` 注册**（`cli.py:58`），否则 `wordf config` 会把它当容器继续往下展开而非打印字段
