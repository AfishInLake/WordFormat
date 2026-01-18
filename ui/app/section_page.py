# app/section_page.py
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QFrame, QLabel, QTabWidget
from ui.data.constants import TRANSLATIONS
from .override_row import OverrideRow


class SectionPage(QWidget):
    """
    根据一个 YAML 节点动态构建配置界面。
    升级版：支持多级 Tab、智能分组 (GroupBox) 和深度递归。
    """

    def __init__(self, key_name, data_node, yaml_manager):
        super().__init__()
        # 页面整体布局
        self.layout = QVBoxLayout(self)
        self.layout.setAlignment(Qt.AlignTop)
        self.layout.setContentsMargins(30, 30, 30, 30)
        self.layout.setSpacing(30)

        self.data = data_node
        self.global_fmt = yaml_manager.get_global_format()

        # --- 核心修改：根据 Key 名称和数据结构决定布局策略 ---

        # 策略 1: 摘要 (Abstract) - 强制分为 中文/英文/关键字 3个Tab
        if key_name == 'abstract':
            # 注意：这里我们手动指定顺序
            self.build_tabs(data_node, ['chinese', 'english', 'keywords'])

        # 策略 2: 参考文献 & 致谢 (Ref & Ack) - 结构为 title + content
        elif 'title' in data_node and 'content' in data_node:
            self.build_tabs(data_node, ['title', 'content'])

        # 策略 3: 各级标题 (Headings) - 检测 level_ 前缀
        elif any(k.startswith('level_') for k in data_node.keys()):
            levels = sorted([k for k in data_node.keys() if k.startswith('level_')])
            self.build_tabs(data_node, levels)

        # 策略 4: 默认扁平结构 (Global, Figures, Tables, Body)
        # 如果不是上述复杂结构，就放入一个卡片中显示
        else:
            self.build_flat_card(data_node)

        self.layout.addStretch()

    def build_tabs(self, data_node, keys):
        """构建 Tab 页，并在 Tab 内部支持递归分组"""
        tab_widget = QTabWidget()

        for k in keys:
            # 如果 YAML 里没有这个 key (比如 keywords)，就跳过
            if k not in data_node:
                continue

            page = QWidget()
            page_layout = QVBoxLayout(page)
            page_layout.setAlignment(Qt.AlignTop)
            page_layout.setContentsMargins(10, 20, 10, 20)

            # === 关键修改：使用递归构建器代替简单的 populate_rows ===
            # 这样如果 Tab 内部还有 chinese_title 这种字典，会自动生成分组框
            self.build_recursive_content(data_node[k], page_layout)

            # 获取翻译后的标题
            title_text = TRANSLATIONS.get(k, k).upper()
            tab_widget.addTab(page, title_text)

        self.layout.addWidget(tab_widget)

    def build_flat_card(self, data_node):
        """构建单张卡片（用于非 Tab 页面）"""
        card = QFrame()
        card.setProperty("class", "Card")
        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(5)
        card_layout.setContentsMargins(30, 30, 30, 30)

        self.build_recursive_content(data_node, card_layout)

        if card_layout.count() > 0:
            self.layout.addWidget(card)

    def build_recursive_content(self, data_dict, layout):
        """
        智能递归构建内容：
        1. 先渲染当前层级的普通属性 (font_size, bold...)
        2. 再检测是否有子字典 (chinese_title...)，如果有，包裹在 GroupBox 中
        """
        # 1. 筛选出当前层级的直接属性 (非字典)
        # 同时也需要补全 Global Format 中的缺失字段
        all_keys = list(data_dict.keys())
        if self.global_fmt and data_dict is not self.global_fmt:
            for gk in self.global_fmt.keys():
                if gk not in all_keys and gk != '<<':
                    all_keys.append(gk)

        unique_keys = sorted(list(set(all_keys)))

        # 第一轮循环：渲染直接属性 (OverrideRow)
        for key in unique_keys:
            if key == '<<': continue
            val = data_dict.get(key)

            # 如果是 None (说明是继承 global 的)，或者不是字典，则是普通属性
            if not isinstance(val, dict):
                row = OverrideRow(key, val, data_dict, self.global_fmt)
                layout.addWidget(row)

        # 第二轮循环：渲染嵌套字典 (GroupBox)
        # 例如 abstract -> chinese 下面还有 chinese_title 和 chinese_content
        for key in list(data_dict.keys()):
            if key == '<<': continue
            val = data_dict.get(key)

            if isinstance(val, dict):
                # 创建一个内部玻璃分组框
                group_box = QFrame()
                # 复用 GlassBox 样式，或者定义一个新的 InnerGroup 样式
                group_box.setStyleSheet("""
                    QFrame {
                        background-color: rgba(0, 0, 0, 0.2);
                        border: 1px solid rgba(255, 255, 255, 0.05);
                        border-radius: 12px;
                    }
                """)
                group_layout = QVBoxLayout(group_box)
                group_layout.setContentsMargins(20, 20, 20, 20)

                # 分组标题
                header = QLabel(TRANSLATIONS.get(key, key).upper())
                header.setProperty("class", "Header")
                header.setStyleSheet("font-size: 16px; color: #38bdf8; margin-bottom: 10px;")
                group_layout.addWidget(header)

                # 递归调用自己，填充 GroupBox
                self.build_recursive_content(val, group_layout)

                layout.addWidget(group_box)

