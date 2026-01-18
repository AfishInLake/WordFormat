from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (QDialog,QVBoxLayout,QFrame,
                             QGraphicsDropShadowEffect,QPushButton,QLabel,QHBoxLayout)
# =============================================================================
# 自定义弹窗类 (Deep Ocean Style)
# =============================================================================
class CyberMessageBox(QDialog):
    def __init__(self, parent, title, message, is_error=False):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(400, 240)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 1. 内部容器 (处理圆角和背景)
        container = QFrame()
        container.setStyleSheet("""
            QFrame {
                /* 与主界面一致的深邃渐变背景 */
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, 
                                                  stop:0 #1e293b, stop:1 #0f172a);
                border: 1px solid rgba(255, 255, 255, 0.15);
                border-radius: 20px;
            }
        """)

        # 添加阴影效果
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(40)
        shadow.setColor(QColor(0, 0, 0, 180))
        shadow.setOffset(0, 10)
        container.setGraphicsEffect(shadow)

        inner_layout = QVBoxLayout(container)
        inner_layout.setContentsMargins(30, 30, 30, 30)
        inner_layout.setSpacing(20)

        # 2. 图标 (使用Emoji或字符)
        lbl_icon = QLabel("✓" if not is_error else "✕")
        lbl_icon.setAlignment(Qt.AlignCenter)
        # 成功用青色，失败用红色
        icon_color = "#38bdf8" if not is_error else "#ef4444"
        lbl_icon.setStyleSheet(f"""
            font-size: 48px; 
            font-weight: bold; 
            color: {icon_color};
            background: transparent; 
            border: none;
        """)

        # 3. 消息内容
        lbl_msg = QLabel(message)
        lbl_msg.setAlignment(Qt.AlignCenter)
        lbl_msg.setWordWrap(True)
        lbl_msg.setStyleSheet("""
            font-size: 16px; 
            color: #e2e8f0; 
            font-weight: 500;
            background: transparent;
            border: none;
        """)

        # 4. 确认按钮 (复用主界面 SaveBtn 的样式)
        btn_ok = QPushButton("确 定")
        btn_ok.setCursor(Qt.PointingHandCursor)
        btn_ok.setFixedSize(120, 40)
        # 按钮样式：渐变 + 圆角
        btn_bg = "qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563eb, stop:1 #06b6d4)" if not is_error else "#ef4444"
        btn_hover = "qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #3b82f6, stop:1 #22d3ee)" if not is_error else "#f87171"

        btn_ok.setStyleSheet(f"""
            QPushButton {{
                background-color: {btn_bg};
                border: none;
                border-radius: 20px;
                color: white;
                font-weight: bold;
                font-size: 15px;
            }}
            QPushButton:hover {{
                background-color: {btn_hover};
            }}
        """)
        btn_ok.clicked.connect(self.accept)

        # 按钮居中容器
        btn_box = QHBoxLayout()
        btn_box.addStretch()
        btn_box.addWidget(btn_ok)
        btn_box.addStretch()

        inner_layout.addWidget(lbl_icon)
        inner_layout.addWidget(lbl_msg)
        inner_layout.addLayout(btn_box)

        layout.addWidget(container)

    # 静态方法，方便像 QMessageBox 一样调用
    @staticmethod
    def show_success(parent, message):
        dialog = CyberMessageBox(parent, "Success", message, is_error=False)
        dialog.exec_()

    @staticmethod
    def show_error(parent, message):
        print("message",message)
        dialog = CyberMessageBox(parent, "Error", message, is_error=True)
        dialog.exec_()
