from PyQt5.QtWidgets import QComboBox
class SafeComboBox(QComboBox):
    """SafeComboBox 类的 wheelEvent 方法通过调用 event.ignore()
     来禁用鼠标滚轮滚动时对下拉框的任何影响。
    这通常用于避免用户不小心滚动鼠标时，改变下拉框的选择项"""

    def wheelEvent(self, event):
        event.ignore()