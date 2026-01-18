from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor
from ruamel.yaml import YAML
from PyQt5.QtWidgets import ( QMainWindow, QWidget, QVBoxLayout,QHBoxLayout,
                              QLabel, QScrollArea, QFrame, QStackedWidget, QListWidget,
                             QPushButton,QFileDialog, QGraphicsDropShadowEffect,)
from ui.data.constants import MENU_ITEMS
from ui.app.section_page import SectionPage
from ui.app.message_box import CyberMessageBox
class ConfigEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("配置编辑器 | PRO")
        self.resize(1280, 900)
        self.setMinimumSize(1280, 900)
        self.setMaximumSize(1280, 900)

        self.yaml = YAML()
        self.yaml.preserve_quotes = True
        self.yaml.indent(mapping=2, sequence=4, offset=2)

        self.data = None
        self.filepath = None

        self.init_ui()

    def init_ui(self):
        main = QWidget()
        self.setCentralWidget(main)

        main_layout = QVBoxLayout(main)
        main_layout.setContentsMargins(40, 30, 40, 40)
        main_layout.setSpacing(25)

        # 顶部导航
        top_header = QHBoxLayout()

        # Logo
        logo_box = QHBoxLayout()
        logo_box.setSpacing(0)
        brand = QLabel("WordFormat 配置编辑器 | ")
        brand.setProperty("class", "BrandLogo")
        brand_bold = QLabel("PRO")
        brand_bold.setProperty("class", "BrandLogo BrandBold")
        logo_box.addWidget(brand)
        logo_box.addWidget(brand_bold)

        # 按钮
        tools = QHBoxLayout()
        tools.setSpacing(20)

        self.btn_load = QPushButton("加载文件")
        self.btn_load.setObjectName("LoadBtn")
        self.btn_load.setCursor(Qt.PointingHandCursor)
        self.btn_load.clicked.connect(self.load_file)

        self.btn_save = QPushButton("保存修改")
        self.btn_save.setObjectName("SaveBtn")
        self.btn_save.setCursor(Qt.PointingHandCursor)
        self.btn_save.setEnabled(False)
        self.btn_save.clicked.connect(self.save_file)

        tools.addStretch()
        tools.addWidget(self.btn_load)
        tools.addWidget(self.btn_save)

        top_header.addLayout(logo_box)
        top_header.addLayout(tools)

        # 玻璃容器
        glass_box = QFrame()
        glass_box.setObjectName("GlassBox")

        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(60)
        shadow.setColor(QColor(0, 0, 0, 150))
        shadow.setOffset(0, 20)
        glass_box.setGraphicsEffect(shadow)

        hbox = QHBoxLayout(glass_box)
        hbox.setContentsMargins(0, 0, 0, 0)
        hbox.setSpacing(0)

        self.sidebar = QListWidget()
        self.sidebar.currentRowChanged.connect(self.change_page)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.stack = QStackedWidget()
        self.stack.setObjectName("ScrollContents")
        self.scroll_area.setWidget(self.stack)

        hbox.addWidget(self.sidebar)
        hbox.addWidget(self.scroll_area)
        hbox.setStretch(1, 1)

        main_layout.addLayout(top_header)
        main_layout.addWidget(glass_box)

        self.show_welcome()

    def show_welcome(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setAlignment(Qt.AlignCenter)
        lbl = QLabel("请加载配置文件")
        lbl.setStyleSheet("color: #64748b; font-size: 26px; font-weight: 300; letter-spacing: 2px;")
        layout.addWidget(lbl)
        self.stack.addWidget(page)

    def get_global_format(self):
        return self.data.get('global_format', {}) if self.data else {}

    def load_file(self):
        fname, _ = QFileDialog.getOpenFileName(self, "选择配置文件", "", "YAML 文件 (*.yaml *.yml)")
        if not fname: return
        try:
            with open(fname, 'r', encoding='utf-8') as f:
                self.data = self.yaml.load(f)
            self.filepath = fname
            self.btn_save.setEnabled(True)
            self.refresh_ui()
        except Exception as e:
            print("e",e)
            CyberMessageBox.show_error(self, f"加载失败:\n{str(e)}")

    def refresh_ui(self):
        self.sidebar.clear()
        while self.stack.count():
            w = self.stack.widget(0)
            self.stack.removeWidget(w)
            w.deleteLater()

        for name, key in MENU_ITEMS:
            self.sidebar.addItem(name)
            if key in self.data:
                # 参数1: key (例如 'abstract', 用于判断布局策略)
                # 参数2: data_node (当前节点的字典数据)
                # 参数3: yaml_manager (就是 self，因为 ConfigEditor 实现了 get_global_format)
                page = SectionPage(key,self.data[key], self,)
                self.stack.addWidget(page)
            else:
                placeholder = QLabel(f"{name} 无配置项")
                placeholder.setAlignment(Qt.AlignCenter)
                placeholder.setStyleSheet("color: #64748b; font-size: 18px;")
                self.stack.addWidget(placeholder)
        self.sidebar.setCurrentRow(0)

    def change_page(self, idx):
        self.stack.setCurrentIndex(idx)
        self.scroll_area.verticalScrollBar().setValue(0)

    def save_file(self):
        if not self.filepath or not self.data: return
        try:
            with open(self.filepath, 'w', encoding='utf-8') as f:
                self.yaml.dump(self.data, f)
            CyberMessageBox.show_success(self, "配置文件保存成功")
        except Exception as e:
            CyberMessageBox.show_error(self, f"保存失败:\n{str(e)}")