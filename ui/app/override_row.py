# app/override_row.py
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QWidget, QHBoxLayout, QLabel, QLineEdit, QCheckBox
from ui.data.constants import OPTIONS_MAP, TRANSLATIONS
from ui.app.safe_widgets import SafeComboBox


class OverrideRow(QWidget):
    """表示一个可编辑的行项控件，通常用于显示和修改键值对的配置项。
    通过 QCheckBox 控制是否启用该行，控件会根据提供的 key 和 current_val
    以及字典数据（data_dict 和 global_dict）来动态展示和处理值
    参数 key: 配置项的标识符
        current_val: 当前值（用来初始化控件显示）
        data_dict: 本地存储数据字典
        global_dict: 全局存储数据字典"""

    def __init__(self, key, current_val, data_dict, global_dict):
        super().__init__()

        # 初始化实例变量
        self.key = key
        self.data_dict = data_dict
        self.global_dict = global_dict or {}  # 保存传入的参数，global_dict 若为空，则赋为一个空字典

        # 创建布局
        layout = QHBoxLayout(self)  # 创建一个水平布局
        layout.setContentsMargins(0, 10, 0, 10)  # 设置边距（上和下 10px）
        layout.setSpacing(20)  # 控件间距为 20px

        # 创建复选框 QCheckBox
        self.chk = QCheckBox()
        self.chk.setCursor(Qt.PointingHandCursor)  # 设置鼠标样式为手指样式

        # 判断是否是全局页面
        # (检查 data_dict 和 global_dict 是否是同一个字典对象，
        # 若是，则意味着这是一个全局页面。)
        is_global_page = (data_dict is global_dict)
        if is_global_page:
            self.chk.setChecked(True)  # 设置复选框为勾选状态并禁用它
            self.chk.setDisabled(True)  # 隐藏复选框的显示样式
            self.chk.setStyleSheet("QCheckBox::indicator { width: 0px; border:none; background:transparent;}")
        else:
            is_present = key in data_dict
            self.chk.setChecked(is_present)  # 判断当前 key 是否存在于 data_dict 中，如果存在，则勾选复选框

        # 设置标签文本
        label_text = TRANSLATIONS.get(key, key)  # 根据 key 查找相应的翻译文本，如果没有翻译，则使用 key 作为默认文本
        self.lbl = QLabel(label_text)  # 创建 QLabel 并设置标签文本
        self.lbl.setProperty("class", "FieldLabel")  # 设置其 CSS 样式类为 "FieldLabel"
        self.lbl.setFixedWidth(120)  # 固定宽度为 120px

        # 创建输入控件
        display_val = current_val
        if display_val is None:
            display_val = self.global_dict.get(key)

        self.widget = self.create_input(key, display_val)  # 调用 create_input 方法创建输入控件
        self.widget.setEnabled(self.chk.isChecked())  # 根据复选框的状态设置是否启用该控件

        # 添加控件到布局
        layout.addWidget(self.chk)  # 复选框
        layout.addWidget(self.lbl)  # 文本框
        layout.addWidget(self.widget, 1)  # 伸展以填充剩余空间（1 表示扩展因子）

        # 连接信号槽
        self.chk.stateChanged.connect(self.on_check_change)  # 连接复选框的状态变化信号到 on_check_change 方法
        if isinstance(self.widget, SafeComboBox):
            self.widget.currentIndexChanged.connect(self.on_val_change)
            self.widget.editTextChanged.connect(self.on_val_change)
        else:
            self.widget.textChanged.connect(self.on_val_change)

    def create_input(self, key, val):
        """kwgs"""
        if key in OPTIONS_MAP:
            cb = SafeComboBox()
            cb.setEditable(True)
            str_val = str(val)
            for label, data in OPTIONS_MAP[key]:
                cb.addItem(label, data)
            idx = cb.findData(val)
            if idx == -1: idx = cb.findData(str_val)  # 尝试匹配 val
            if idx != -1:
                cb.setCurrentIndex(idx)
            else:
                cb.setCurrentText(str_val)
            return cb
        le = QLineEdit(str(val) if val is not None else "")
        return le

    def on_check_change(self, state):
        """ 如果复选框被勾选，启用控件并保存当前值。
            如果复选框取消勾选，禁用控件并根据全局字典设置控件的值"""
        # 需要修改逻辑
        checked = (state == Qt.Checked)
        self.widget.setEnabled(checked)
        if checked:
            self.save_value()
        else:
            if self.key in self.data_dict:
                del self.data_dict[self.key]
            g_val = self.global_dict.get(self.key)
            if g_val is not None:
                self.set_widget_value(g_val)

    def on_val_change(self):
        """ 当控件的值变化时，如果复选框被勾选，则保存新的值"""
        if self.chk.isChecked():
            self.save_value()

    def save_value(self):
        """保存当前控件的值到 data_dict 中。
        对于 SafeComboBox，获取选中的数据或文本；
        对于其他控件，获取文本值。
        如果值可以转为 int 或 float，则尝试转换"""
        val = None
        if isinstance(self.widget, SafeComboBox):
            val = self.widget.currentData()
            if val is None: val = self.widget.currentText()
        else:
            val = self.widget.text()
        if not isinstance(val, (bool, int, float)):
            try:
                val = int(val)
            except:
                try:
                    val = float(val)
                except:
                    pass
        self.data_dict[self.key] = val

    def set_widget_value(self, val):
        """根据给定的值 val 设置控件的显示内容。
        如果是 SafeComboBox，则设置选中的项；否则设置文本框的文本。"""
        if isinstance(self.widget, SafeComboBox):
            idx = self.widget.findData(val)
            if idx != -1:
                self.widget.setCurrentIndex(idx)
            else:
                self.widget.setCurrentText(str(val))
        else:
            self.widget.setText(str(val))