import sys
import os
import json
import shutil
import tempfile
from datetime import datetime
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit,
                             QLabel, QFileDialog, QTextEdit, QListWidget, QMessageBox, QStyle, QStyleFactory,
                             QProgressBar, QTableWidget, QTableWidgetItem, QHeaderView, QDialogButtonBox,
                             QMainWindow, QToolBar, QAbstractItemView, QMenu, QDialog)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize, QMimeData, QTimer
from PyQt6.QtGui import QIcon, QFont, QPalette, QColor, QDragEnterEvent, QDropEvent, QAction
from docx import Document
import copy

class ReplacementWorker(QThread):
    progress = pyqtSignal(int)
    file_processed = pyqtSignal(str, bool, int)
    finished = pyqtSignal(dict)

    def __init__(self, files, rules, backup_dir):
        super().__init__()
        self.files = files
        self.rules = rules
        self.backup_dir = backup_dir

    def run(self):
        total_files = len(self.files)
        stats = {
            "total_files": total_files,
            "changed_files": 0,
            "total_replacements": 0
        }
        for i, file_path in enumerate(self.files):
            try:
                backup_path = os.path.join(self.backup_dir, os.path.basename(file_path))
                shutil.copy2(file_path, backup_path)

                doc = Document(file_path)
                changed = False
                file_replacements = 0
                for old_text, new_text in self.rules:
                    replacements = self.replace_text_in_document(doc, old_text, new_text)
                    if replacements > 0:
                        changed = True
                        file_replacements += replacements
                if changed:
                    doc.save(file_path)
                    stats["changed_files"] += 1
                    stats["total_replacements"] += file_replacements
                    self.file_processed.emit(file_path, True, file_replacements)
                else:
                    self.file_processed.emit(file_path, False, 0)
                    os.remove(backup_path)
            except Exception as e:
                self.file_processed.emit(file_path, False, 0)
            self.progress.emit(int((i + 1) / total_files * 100))
        self.finished.emit(stats)

    def replace_text_in_document(self, doc, old_text, new_text):
        replacements = 0
        for para in doc.paragraphs:
            if old_text in para.text:
                para.text = para.text.replace(old_text, new_text)
                replacements += para.text.count(new_text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if old_text in cell.text:
                        cell.text = cell.text.replace(old_text, new_text)
                        replacements += cell.text.count(new_text)
        return replacements

class LoadingDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("处理中")
        self.setFixedSize(200, 100)
        layout = QVBoxLayout(self)
        self.label = QLabel("正在处理，请稍候...", self)
        layout.addWidget(self.label)
        self.progress_bar = QProgressBar(self)
        layout.addWidget(self.progress_bar)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)

    def update_progress(self, value):
        self.progress_bar.setValue(value)
class WordReplacerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.replacement_history = []
        self.temp_dir = tempfile.mkdtemp()
        self.file_set = set()  # 用于存储已添加的文件路径

        # 启用拖放
        self.setAcceptDrops(True)

    def initUI(self):
        self.setStyle(QStyleFactory.create('Fusion'))
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QPushButton {
                color: #2196F3;
                border: 1px solid #2196F3;
                padding: 5px 10px;
                text-align: center;
                text-decoration: none;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
                background-color: transparent;
            }
            QPushButton:hover {
                background-color: #E3F2FD;
            }
            QTableWidget {
                gridline-color: #d3d3d3;
            }
            QHeaderView::section {
                background-color: #e0e0e0;
                padding: 4px;
                border: 1px solid #d3d3d3;
                font-weight: bold;
            }
        """)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 工具栏
        toolbar = QToolBar()
        self.addToolBar(toolbar)

        # 文件选择
        file_layout = QHBoxLayout()
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.file_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self.show_file_list_context_menu)
        file_layout.addWidget(self.file_list)
        file_buttons_layout = QVBoxLayout()

        add_file_button = QPushButton('添加文件')
        add_file_button.setIcon(QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon)))
        add_file_button.clicked.connect(self.add_files)

        add_folder_button = QPushButton('添加文件夹')
        add_folder_button.setIcon(QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DirIcon)))
        add_folder_button.clicked.connect(self.add_folder)

        remove_button = QPushButton('移除选中')
        remove_button.setIcon(QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon)))
        remove_button.clicked.connect(self.remove_selected)

        file_buttons_layout.addWidget(add_file_button)
        file_buttons_layout.addWidget(add_folder_button)
        file_buttons_layout.addWidget(remove_button)
        file_layout.addLayout(file_buttons_layout)
        layout.addLayout(file_layout)

        # 替换规则表格
        self.rules_table = QTableWidget(0, 2)
        self.rules_table.setHorizontalHeaderLabels(['要替换的文本', '新文本'])
        self.rules_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.rules_table)

        # 添加和删除规则按钮
        rules_buttons_layout = QHBoxLayout()
        add_rule_button = QPushButton('添加规则')
        add_rule_button.setIcon(QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogNewFolder)))
        add_rule_button.clicked.connect(self.add_rule)

        remove_rule_button = QPushButton('删除选中规则')
        remove_rule_button.setIcon(QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogDiscardButton)))
        remove_rule_button.clicked.connect(self.remove_rule)

        rules_buttons_layout.addWidget(add_rule_button)
        rules_buttons_layout.addWidget(remove_rule_button)
        layout.addLayout(rules_buttons_layout)

        # 替换按钮
        replace_button = QPushButton('替换')
        replace_button.setIcon(QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload)))
        replace_button.clicked.connect(self.replace_text)
        layout.addWidget(replace_button)

        # 进度条
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        # 操作日志区域
        log_layout = QVBoxLayout()
        log_layout.addWidget(QLabel('操作日志:'))
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        log_layout.addWidget(self.log_area)

        # 添加清理日志按钮
        clear_log_button = QPushButton('清理日志')
        clear_log_button.setIcon(QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogResetButton)))
        clear_log_button.clicked.connect(self.clear_log)
        log_layout.addWidget(clear_log_button)

        layout.addLayout(log_layout)

        # 撤销按钮
        undo_button = QPushButton('撤销上次替换')
        undo_button.setIcon(QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ArrowBack)))
        undo_button.clicked.connect(self.undo_last_replacement)
        layout.addWidget(undo_button)

        self.setWindowTitle('高级 Word 文本替换器')
        self.setGeometry(300, 300, 800, 600)
    def show_file_list_context_menu(self, position):
        menu = QMenu()
        remove_action = menu.addAction("移除选中")
        open_folder_action = menu.addAction("在资源管理器中打开")

        action = menu.exec(self.file_list.mapToGlobal(position))
        if action == remove_action:
            self.remove_selected()
        elif action == open_folder_action:
            self.open_selected_in_explorer()

    def open_selected_in_explorer(self):
        selected_items = self.file_list.selectedItems()
        if not selected_items:
            return
        file_path = selected_items[0].text()
        os.startfile(os.path.dirname(file_path))

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择 Word 文件", "", "Word 文件 (*.docx)")
        new_files = [file for file in files if file not in self.file_set]
        if new_files:
            self.file_list.addItems(new_files)
            self.file_set.update(new_files)
            self.log(f"已添加 {len(new_files)} 个新文件。")
        if len(new_files) < len(files):
            self.log(f"已跳过 {len(files) - len(new_files)} 个重复文件。")

    def add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            new_files = []
            for root, dirs, files in os.walk(folder):
                for file in files:
                    if file.endswith('.docx'):
                        file_path = os.path.join(root, file)
                        if file_path not in self.file_set:
                            new_files.append(file_path)
                            self.file_set.add(file_path)
            self.file_list.addItems(new_files)
            self.log(f"已从文件夹添加 {len(new_files)} 个新文件。")
            if len(new_files) < len([f for f in os.listdir(folder) if f.endswith('.docx')]):
                self.log("部分文件因重复而被跳过。")

    def remove_selected(self):
        selected_items = self.file_list.selectedItems()
        if not selected_items:
            return
        for item in selected_items:
            file_path = item.text()
            self.file_list.takeItem(self.file_list.row(item))
            self.file_set.remove(file_path)
        self.log(f"已移除 {len(selected_items)} 个文件。")

    def add_rule(self):
        row_position = self.rules_table.rowCount()
        self.rules_table.insertRow(row_position)
        self.rules_table.setItem(row_position, 0, QTableWidgetItem(""))
        self.rules_table.setItem(row_position, 1, QTableWidgetItem(""))
        self.validate_rules()

    def remove_rule(self):
        indices = self.rules_table.selectionModel().selectedRows()
        for index in sorted(indices, reverse=True):
            self.rules_table.removeRow(index.row())
        self.validate_rules()

    def validate_rules(self):
        rules = set()
        for row in range(self.rules_table.rowCount()):
            old_text = self.rules_table.item(row, 0).text().strip()
            new_text = self.rules_table.item(row, 1).text().strip()
            if old_text and new_text:
                if old_text == new_text:
                    self.rules_table.item(row, 0).setBackground(QColor(255, 200, 200))
                    self.rules_table.item(row, 1).setBackground(QColor(255, 200, 200))
                elif (old_text, new_text) in rules:
                    self.rules_table.item(row, 0).setBackground(QColor(255, 255, 200))
                    self.rules_table.item(row, 1).setBackground(QColor(255, 255, 200))
                else:
                    self.rules_table.item(row, 0).setBackground(QColor(255, 255, 255))
                    self.rules_table.item(row, 1).setBackground(QColor(255, 255, 255))
                    rules.add((old_text, new_text))
            else:
                self.rules_table.item(row, 0).setBackground(QColor(255, 255, 255))
                self.rules_table.item(row, 1).setBackground(QColor(255, 255, 255))

    def replace_text(self):
        files = list(self.file_set)
        rules = []
        for row in range(self.rules_table.rowCount()):
            old_text = self.rules_table.item(row, 0).text().strip()
            new_text = self.rules_table.item(row, 1).text().strip()
            if old_text and new_text and old_text != new_text:
                rules.append((old_text, new_text))

        if not files or not rules:
            self.log("警告：请添加文件和有效的替换规则。")
            return

        confirm = QMessageBox.question(self, '确认', f'是否执行替换操作？\n文件数：{len(files)}\n规则数：{len(rules)}',
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm == QMessageBox.StandardButton.No:
            return

        backup_dir = os.path.join(self.temp_dir, f'backup_{len(self.replacement_history)}')
        os.makedirs(backup_dir, exist_ok=True)

        self.log(f"开始替换操作：处理 {len(files)} 个文件，应用 {len(rules)} 条规则。")

        self.worker = ReplacementWorker(files, rules, backup_dir)
        self.worker.progress.connect(self.update_progress)
        self.worker.file_processed.connect(self.update_output)
        self.worker.finished.connect(self.replacement_finished)

        self.loading_dialog = LoadingDialog(self)
        self.worker.progress.connect(self.loading_dialog.update_progress)
        self.loading_dialog.show()

        self.worker.start()

        self.replacement_history.append((files, rules, backup_dir))

    def update_progress(self, value):
        self.progress_bar.setValue(value)
        if value == 100:
            self.progress_bar.setStyleSheet("""
                QProgressBar {
                    border: 1px solid #4CAF50;
                    border-radius: 5px;
                    text-align: center;
                }
                QProgressBar::chunk {
                    background-color: #C8E6C9;
                }
            """)
        else:
            self.progress_bar.setStyleSheet("""
                QProgressBar {
                    border: 1px solid #2196F3;
                    border-radius: 5px;
                    text-align: center;
                }
                QProgressBar::chunk {
                    background-color: #2196F3;
                }
            """)

    def update_output(self, file_path, changed, replacements):
        status = f"已替换 ({replacements} 处)" if changed else "未更改"
        self.log(f"{file_path}: {status}")

    def replacement_finished(self, stats):
        self.loading_dialog.close()
        self.log("替换操作完成。")
        self.progress_bar.setValue(100)

        summary = (f"替换操作摘要:\n"
                   f"处理文件总数: {stats['total_files']}\n"
                   f"发生更改的文件数: {stats['changed_files']}\n"
                   f"总替换次数: {stats['total_replacements']}")

        QMessageBox.information(self, "替换完成", summary)

    def undo_last_replacement(self):
        if not self.replacement_history:
            self.log("没有可撤销的操作。")
            return

        files, rules, backup_dir = self.replacement_history.pop()
        confirm = QMessageBox.question(self, '确认', f'是否撤销上次替换操作？\n文件数：{len(files)}\n规则数：{len(rules)}',
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm == QMessageBox.StandardButton.No:
            return

        self.log("开始撤销上次替换操作...")
        for file_path in files:
            backup_path = os.path.join(backup_dir, os.path.basename(file_path))
            if os.path.exists(backup_path):
                try:
                    shutil.copy2(backup_path, file_path)
                    self.log(f"已恢复文件: {file_path}")
                except Exception as e:
                    self.log(f"恢复文件失败: {file_path}, 错误: {str(e)}")
            else:
                self.log(f"未找到备份文件: {file_path}")

        self.log("撤销操作完成。")

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        self.process_dropped_files(files)

    def process_dropped_files(self, files):
        new_files = []
        for file in files:
            if os.path.isdir(file):
                for root, dirs, files in os.walk(file):
                    for f in files:
                        if f.endswith('.docx'):
                            file_path = os.path.join(root, f)
                            if file_path not in self.file_set:
                                new_files.append(file_path)
                                self.file_set.add(file_path)
            elif file.endswith('.docx') and file not in self.file_set:
                new_files.append(file)
                self.file_set.add(file)

        self.file_list.addItems(new_files)
        self.log(f"通过拖放添加了 {len(new_files)} 个新文件。")
        if len(new_files) < len(files):
            self.log(f"跳过了 {len(files) - len(new_files)} 个重复或非Word文件。")

    def log(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_area.append(f"[{timestamp}] {message}")

    def clear_log(self):
        self.log_area.clear()
        self.log("日志已清理。")

    def closeEvent(self, event):
        # 关闭应用时清理临时目录
        shutil.rmtree(self.temp_dir, ignore_errors=True)
        super().closeEvent(event)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create('Fusion'))
    ex = WordReplacerApp()
    ex.show()
    sys.exit(app.exec())