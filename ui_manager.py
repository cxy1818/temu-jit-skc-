
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLineEdit, QLabel, QComboBox,
    QTableWidget, QFrame
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap, QDragEnterEvent, QDropEvent


status_options = ["核价通过", "拉过库存", "已下架", "价格待定", "减少库存为0", "改过体积", "价格错误"]

class ImageDropLabel(QLabel):
    """图片拖拽控件"""
    def __init__(self):
        super().__init__("拖拽或粘贴图片")
        self.setAlignment(Qt.AlignCenter)
        self.setFrameShape(QFrame.Box)
        self.setAcceptDrops(True)
        self.image_path = None

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                self.setPixmap(QPixmap(path).scaled(150, 150, Qt.KeepAspectRatio))
                self.image_path = path

    def clear(self):
        self.setText("拖拽或粘贴图片")
        self.setPixmap(QPixmap())
        self.image_path = None

class SKCUI(QWidget):
    """界面布局，UI独立于逻辑"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SKC 管理器 作者联系方式微信号cxy-cxy-1188")
        self.resize(1200, 650)
        main_layout = QHBoxLayout(self)

        # 左侧栏
        left_col = QVBoxLayout()
        main_layout.addLayout(left_col, 0)

        left_col.addWidget(QLabel("当前项目:"))
        self.project_combo = QComboBox()
        left_col.addWidget(self.project_combo)

        # 项目管理按钮
        self.btn_new_project = QPushButton("新建项目")
        self.btn_switch_project = QPushButton("切换项目")
        self.btn_import_project = QPushButton("导入项目")
        self.btn_export_project = QPushButton("导出项目")
        for btn in [self.btn_new_project, self.btn_switch_project,
                    self.btn_import_project, self.btn_export_project]:
            left_col.addWidget(btn)

        left_col.addWidget(QLabel("货号:"))
        self.entry_product = QLineEdit()
        left_col.addWidget(self.entry_product)

        left_col.addWidget(QLabel("SKC (空格隔开):"))
        skc_h = QHBoxLayout()
        self.entry_skc = QLineEdit()
        skc_h.addWidget(self.entry_skc)
        self.btn_clear_skc = QPushButton("✕")
        self.btn_clear_skc.setFixedWidth(28)
        skc_h.addWidget(self.btn_clear_skc)
        left_col.addLayout(skc_h)

        left_col.addWidget(QLabel("状态:"))
        self.status_combo = QComboBox()
        self.status_combo.addItems(status_options)
        left_col.addWidget(self.status_combo)

        # 操作按钮
        self.btn_add = QPushButton("添加SKC")
        self.btn_batch_modify = QPushButton("批量修改 SKC")
        self.btn_batch_delete = QPushButton("批量删除 SKC")
        self.btn_auto_sort = QPushButton("自动整理/手动保存")
        for btn in [self.btn_add, self.btn_batch_modify, self.btn_batch_delete, self.btn_auto_sort]:
            left_col.addWidget(btn)

        # 图片拖拽区
        left_col.addWidget(QLabel("选择货号添加图片:"))
        self.image_product_combo = QComboBox()
        left_col.addWidget(self.image_product_combo)
        self.image_drop_label = ImageDropLabel()
        self.image_drop_label.setFixedSize(160, 160)
        left_col.addWidget(self.image_drop_label)
        self.btn_add_image = QPushButton("确认添加图片")
        left_col.addWidget(self.btn_add_image)

        # Excel 导入/打开
        self.btn_import_excel = QPushButton("导入 Excel 数据")
        self.btn_open_latest = QPushButton("打开 Excel")
        for btn in [self.btn_import_excel, self.btn_open_latest]:
            left_col.addWidget(btn)

        left_col.addStretch()

        # 数据表格
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["货号", "SKC", "状态"])
        main_layout.addWidget(self.table, 1)
