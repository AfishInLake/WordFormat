# =============================================================================
# 样式表：深海科技风格 + 显眼的下拉按钮
# =============================================================================
ARROW_ICON = "url(\"data:image/svg+xml;charset=utf-8,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%2338bdf8' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E\")"

STYLESHEET = f"""
/* 
   1. 全局背景：午夜蓝深邃渐变
*/
QMainWindow {{
    background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, 
                                      stop:0 #1e293b, 
                                      stop:1 #0f172a);
}}

QWidget {{ 
    font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif; 
    font-size: 15px; 
    color: #e2e8f0; 
}}

/* 
   2. 玻璃容器
*/
QFrame#GlassBox {{
    background-color: rgba(30, 41, 59, 0.7);
    border: 1px solid rgba(255, 255, 255, 0.1);
    border-radius: 20px;
}}

QFrame.Card {{
    background-color: rgba(255, 255, 255, 0.03);
    border: 1px solid rgba(255, 255, 255, 0.05);
    border-radius: 16px;
}}

/* 
   3. 侧边栏
*/
QListWidget {{
    background-color: transparent;
    border: none;
    outline: none;
    min-width: 240px;
    padding: 30px 15px;
    border-right: 1px solid rgba(255, 255, 255, 0.05);
}}
QListWidget::item {{
    height: 55px;
    margin-bottom: 10px;
    padding-left: 20px;
    border-radius: 12px;
    color: #94a3b8;
    font-weight: 500;
    font-size: 15px;
    border: 1px solid transparent;
    transition: all 0.2s;
}}
QListWidget::item:hover {{
    background-color: rgba(56, 189, 248, 0.1);
    color: #f1f5f9;
}}
QListWidget::item:selected {{
    background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563eb, stop:1 #06b6d4);
    color: #ffffff;
    font-weight: 700;
    border: none;
    border-left: 4px solid #ffffff; 
}}

/* 
   4. 右侧滚动区 
*/
QScrollArea {{ border: none; background-color: transparent; }}
QWidget#ScrollContents {{ background-color: transparent; }}

QScrollBar:vertical {{
    border: none;
    background: transparent;
    width: 6px;
    margin: 0;
}}
QScrollBar::handle:vertical {{
    background: rgba(148, 163, 184, 0.3);
    min-height: 40px;
    border-radius: 3px;
}}
QScrollBar::handle:vertical:hover {{ background: rgba(148, 163, 184, 0.5); }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}

/* 
   5. Tab 页签
*/
QTabWidget::pane {{ border: none; }}
QTabWidget::tab-bar {{ left: 0px; }}
QTabBar::tab {{
    background: transparent;
    color: #64748b;
    padding: 10px 20px;
    border-bottom: 2px solid transparent;
    margin-right: 15px;
    font-weight: 600;
}}
QTabBar::tab:hover {{ color: #38bdf8; }}
QTabBar::tab:selected {{
    color: #38bdf8;
    border-bottom: 2px solid #38bdf8;
}}

/* 
   6. 输入控件 (LineEdit)
*/
QLineEdit {{
    background-color: rgba(15, 23, 42, 0.6);
    border: 1px solid rgba(148, 163, 184, 0.2);
    border-radius: 8px;
    padding: 0 15px;
    height: 42px;
    color: #f1f5f9;
    font-weight: 500;
}}
QLineEdit:focus {{
    border: 1px solid #38bdf8;
    background-color: rgba(15, 23, 42, 0.8);
}}
QLineEdit:disabled {{
    background-color: rgba(255, 255, 255, 0.03);
    color: #64748b;
    border-color: transparent;
}}

/* 
   7. 下拉框 (QComboBox) - 重点修改部分
*/
QComboBox {{
    background-color: rgba(15, 23, 42, 0.6);
    border: 1px solid rgba(148, 163, 184, 0.2);
    border-radius: 8px;
    padding-left: 15px;
    /* 给右侧箭头留出空间，防止文字重叠 */
    padding-right: 40px; 
    height: 42px;
    color: #f1f5f9;
    font-weight: 500;
}}
QComboBox:focus {{
    border: 1px solid #38bdf8;
    background-color: rgba(15, 23, 42, 0.8);
}}
QComboBox:disabled {{
    background-color: rgba(255, 255, 255, 0.03);
    color: #64748b;
    border-color: transparent;
}}

/* 下拉按钮区域 */
QComboBox::drop-down {{
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 40px; /* 按钮宽度足够大 */

    border-left-width: 1px;
    border-left-color: rgba(148, 163, 184, 0.2);
    border-left-style: solid; /* 左侧加一道分割线 */

    border-top-right-radius: 8px;
    border-bottom-right-radius: 8px;
}}

/* 鼠标悬停在下拉框时的箭头区域背景 */
QComboBox::drop-down:hover {{
    background-color: rgba(56, 189, 248, 0.1);
}}

/* 下拉箭头图标 (使用 Base64 SVG) */
QComboBox::down-arrow {{
    image: {ARROW_ICON};
    width: 16px;
    height: 16px;
}}

/* 展开后的列表视图 */
QComboBox QAbstractItemView {{
    background-color: #1e293b;
    color: white;
    selection-background-color: #2563eb;
    border: 1px solid #334155;
    outline: none;
    padding: 5px;
}}

/* 
   8. 复选框
*/
QCheckBox {{ spacing: 12px; font-weight: 600; color: #94a3b8; }}
QCheckBox::indicator {{
    width: 22px; height: 22px;
    border-radius: 6px;
    border: 1px solid rgba(148, 163, 184, 0.4);
    background: rgba(15, 23, 42, 0.4);
}}
QCheckBox::indicator:checked {{
    background-color: #38bdf8;
    border-color: #38bdf8;
}}

/* 
   9. 按钮组
*/
QPushButton#LoadBtn {{
    background: rgba(255, 255, 255, 0.05);
    border: 1px solid rgba(255, 255, 255, 0.15);
    border-radius: 21px;
    height: 42px;
    padding: 0 25px;
    color: #e2e8f0;
    font-weight: 600;
}}
QPushButton#LoadBtn:hover {{
    background: rgba(255, 255, 255, 0.1);
    border-color: #38bdf8;
    color: #38bdf8;
}}

QPushButton#SaveBtn {{
    background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563eb, stop:1 #06b6d4);
    border: none;
    border-radius: 21px;
    height: 42px;
    padding: 0 35px;
    color: #ffffff;
    font-weight: 700;
    font-size: 15px;
}}
QPushButton#SaveBtn:hover {{
    background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #3b82f6, stop:1 #22d3ee);
}}
QPushButton#SaveBtn:disabled {{
    background-color: rgba(255, 255, 255, 0.1);
    color: #64748b;
}}

/* 文字标签 */
QLabel.BrandLogo {{ 
    font-size: 18px; 
    font-weight: 400; 
    color: #e2e8f0; 
    letter-spacing: 1px;
}}
QLabel.BrandBold {{ font-weight: 800; color: #38bdf8; }}

QLabel.Header {{ font-size: 18px; font-weight: 700; color: #f8fafc; margin-bottom: 5px;}}
QLabel.FieldLabel {{ font-size: 15px; font-weight: 500; color: #cbd5e1; }}
"""
# 定义箭头的 Base64 SVG (青蓝色 Chevron)