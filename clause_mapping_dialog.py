# -*- coding: utf-8 -*-
"""
条款映射设置对话框
提供可视化的映射管理界面

功能：
1. 查看所有映射（表格形式）
2. 单条添加/编辑/删除
3. 批量导入（JSON/Excel）
4. 导出映射
5. 搜索过滤

Author: Dachi Yijin
Date: 2025-12-23
"""

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QLineEdit, QLabel, QGroupBox, QMessageBox, QFileDialog,
    QHeaderView, QAbstractItemView, QWidget, QFormLayout, QComboBox,
    QSpinBox, QCheckBox, QSplitter, QFrame, QMenu, QAction
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QIcon

from clause_mapping_manager import ClauseMappingManager, ClauseMapping


class MappingEditDialog(QDialog):
    """单条映射编辑对话框"""

    def __init__(self, parent=None, mapping: ClauseMapping = None, library_clauses: list = None):
        super().__init__(parent)
        self.mapping = mapping
        self.library_clauses = library_clauses or []
        self.result = None
        self._setup_ui()

    def _setup_ui(self):
        self.setWindowTitle("编辑条款映射" if self.mapping else "添加条款映射")
        self.setMinimumWidth(600)
        self.setModal(True)

        layout = QVBoxLayout(self)

        # 表单
        form_layout = QFormLayout()

        # 英文条款名
        self.edit_english = QLineEdit()
        self.edit_english.setPlaceholderText("例如: Off Premises Property Clause")
        if self.mapping:
            self.edit_english.setText(self.mapping.english)
        form_layout.addRow("英文条款名:", self.edit_english)

        # 中文条款名
        self.edit_chinese = QLineEdit()
        self.edit_chinese.setPlaceholderText("例如: 场所外财产条款")
        if self.mapping:
            self.edit_chinese.setText(self.mapping.chinese)
        form_layout.addRow("中文条款名:", self.edit_chinese)

        # 库内条款名（带下拉建议）
        self.combo_library = QComboBox()
        self.combo_library.setEditable(True)
        self.combo_library.setInsertPolicy(QComboBox.NoInsert)
        self.combo_library.lineEdit().setPlaceholderText("例如: 企业财产保险附加场外维修及改造条款")

        # 添加库内条款列表
        if self.library_clauses:
            self.combo_library.addItems(self.library_clauses)

        if self.mapping and self.mapping.library:
            self.combo_library.setCurrentText(self.mapping.library)

        form_layout.addRow("库内条款名:", self.combo_library)

        # 备注
        self.edit_notes = QLineEdit()
        self.edit_notes.setPlaceholderText("可选备注")
        if self.mapping:
            self.edit_notes.setText(self.mapping.notes)
        form_layout.addRow("备注:", self.edit_notes)

        layout.addLayout(form_layout)

        # 按钮
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(self.btn_cancel)

        self.btn_save = QPushButton("保存")
        self.btn_save.setDefault(True)
        self.btn_save.clicked.connect(self._on_save)
        btn_layout.addWidget(self.btn_save)

        layout.addLayout(btn_layout)

    def _on_save(self):
        english = self.edit_english.text().strip()
        chinese = self.edit_chinese.text().strip()
        library = self.combo_library.currentText().strip()

        if not english:
            QMessageBox.warning(self, "提示", "请输入英文条款名")
            return

        if not chinese and not library:
            QMessageBox.warning(self, "提示", "请至少填写中文条款名或库内条款名")
            return

        self.result = {
            'english': english,
            'chinese': chinese,
            'library': library,
            'notes': self.edit_notes.text().strip(),
        }
        self.accept()


class ImportSettingsDialog(QDialog):
    """导入设置对话框"""

    def __init__(self, parent=None, file_type: str = "excel"):
        super().__init__(parent)
        self.file_type = file_type
        self.result = None
        self._setup_ui()

    def _setup_ui(self):
        self.setWindowTitle(f"导入设置 - {'Excel' if self.file_type == 'excel' else 'JSON'}")
        self.setMinimumWidth(400)
        self.setModal(True)

        layout = QVBoxLayout(self)

        # 覆盖选项
        self.check_overwrite = QCheckBox("覆盖已有映射（推荐）")
        self.check_overwrite.setChecked(True)
        self.check_overwrite.setToolTip("如果导入的条款已存在，将更新为新的映射")
        layout.addWidget(self.check_overwrite)

        # Excel 特有选项
        if self.file_type == "excel":
            group = QGroupBox("Excel 列设置")
            form = QFormLayout(group)

            self.spin_header = QSpinBox()
            self.spin_header.setRange(0, 10)
            self.spin_header.setValue(1)
            self.spin_header.setToolTip("跳过的表头行数")
            form.addRow("表头行数:", self.spin_header)

            self.spin_col_en = QSpinBox()
            self.spin_col_en.setRange(1, 26)
            self.spin_col_en.setValue(1)
            self.spin_col_en.setToolTip("英文条款名所在列（1=A列）")
            form.addRow("英文列:", self.spin_col_en)

            self.spin_col_cn = QSpinBox()
            self.spin_col_cn.setRange(1, 26)
            self.spin_col_cn.setValue(2)
            self.spin_col_cn.setToolTip("中文条款名所在列（2=B列）")
            form.addRow("中文列:", self.spin_col_cn)

            self.spin_col_lib = QSpinBox()
            self.spin_col_lib.setRange(0, 26)
            self.spin_col_lib.setValue(3)
            self.spin_col_lib.setToolTip("库内条款名所在列（0=无此列，3=C列）")
            form.addRow("库内条款列:", self.spin_col_lib)

            layout.addWidget(group)

        # 按钮
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        btn_cancel = QPushButton("取消")
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_cancel)

        btn_ok = QPushButton("开始导入")
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(self._on_ok)
        btn_layout.addWidget(btn_ok)

        layout.addLayout(btn_layout)

    def _on_ok(self):
        self.result = {
            'overwrite': self.check_overwrite.isChecked(),
        }

        if self.file_type == "excel":
            self.result['header_row'] = self.spin_header.value()
            self.result['col_english'] = self.spin_col_en.value() - 1
            self.result['col_chinese'] = self.spin_col_cn.value() - 1
            self.result['col_library'] = self.spin_col_lib.value() - 1 if self.spin_col_lib.value() > 0 else -1

        self.accept()


class ClauseMappingDialog(QDialog):
    """条款映射管理主对话框"""

    mappings_changed = pyqtSignal()  # 映射变更信号

    def __init__(self, parent=None, library_clauses: list = None):
        super().__init__(parent)
        self.manager = ClauseMappingManager.get_instance()
        self.library_clauses = library_clauses or []
        self._setup_ui()
        self._load_data()

    def _setup_ui(self):
        self.setWindowTitle("条款映射设置")
        self.setMinimumSize(1000, 600)

        layout = QVBoxLayout(self)

        # 顶部工具栏
        toolbar = QHBoxLayout()

        # 搜索框
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("搜索条款...")
        self.search_edit.textChanged.connect(self._filter_table)
        self.search_edit.setMaximumWidth(300)
        toolbar.addWidget(self.search_edit)

        toolbar.addStretch()

        # 统计标签
        self.label_count = QLabel("共 0 条映射")
        toolbar.addWidget(self.label_count)

        layout.addLayout(toolbar)

        # 映射表格
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            "英文条款名", "中文条款名", "库内条款名", "更新时间", "来源", "备注"
        ])

        # 表格样式
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setSortingEnabled(True)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.doubleClicked.connect(self._on_edit)

        # 列宽
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Interactive)
        header.setSectionResizeMode(1, QHeaderView.Interactive)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.Interactive)

        self.table.setColumnWidth(0, 250)
        self.table.setColumnWidth(1, 150)
        self.table.setColumnWidth(5, 100)

        # 右键菜单
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)

        layout.addWidget(self.table)

        # 底部按钮栏
        btn_layout = QHBoxLayout()

        # 左侧按钮
        self.btn_add = QPushButton("➕ 添加")
        self.btn_add.clicked.connect(self._on_add)
        btn_layout.addWidget(self.btn_add)

        self.btn_edit = QPushButton("✏️ 编辑")
        self.btn_edit.clicked.connect(self._on_edit)
        btn_layout.addWidget(self.btn_edit)

        self.btn_delete = QPushButton("🗑️ 删除")
        self.btn_delete.clicked.connect(self._on_delete)
        btn_layout.addWidget(self.btn_delete)

        btn_layout.addWidget(self._create_separator())

        # 导入导出按钮
        self.btn_import_json = QPushButton("📥 导入 JSON")
        self.btn_import_json.clicked.connect(self._on_import_json)
        btn_layout.addWidget(self.btn_import_json)

        self.btn_import_excel = QPushButton("📥 导入 Excel")
        self.btn_import_excel.clicked.connect(self._on_import_excel)
        btn_layout.addWidget(self.btn_import_excel)

        self.btn_export = QPushButton("📤 导出")
        self.btn_export.clicked.connect(self._on_export)
        btn_layout.addWidget(self.btn_export)

        btn_layout.addStretch()

        # 右侧按钮
        self.btn_save = QPushButton("💾 保存")
        self.btn_save.clicked.connect(self._on_save)
        btn_layout.addWidget(self.btn_save)

        self.btn_close = QPushButton("关闭")
        self.btn_close.clicked.connect(self._on_close)
        btn_layout.addWidget(self.btn_close)

        layout.addLayout(btn_layout)

    def _create_separator(self) -> QFrame:
        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setFrameShadow(QFrame.Sunken)
        return sep

    def _load_data(self):
        """加载映射数据到表格"""
        self.manager.load()
        self._refresh_table()

    def _refresh_table(self):
        """刷新表格显示"""
        mappings = self.manager.get_all_mappings()

        self.table.setRowCount(len(mappings))

        for row, mapping in enumerate(mappings):
            self.table.setItem(row, 0, QTableWidgetItem(mapping.english))
            self.table.setItem(row, 1, QTableWidgetItem(mapping.chinese))
            self.table.setItem(row, 2, QTableWidgetItem(mapping.library))
            self.table.setItem(row, 3, QTableWidgetItem(mapping.updated))
            self.table.setItem(row, 4, QTableWidgetItem(mapping.source))
            self.table.setItem(row, 5, QTableWidgetItem(mapping.notes))

            # 高亮没有库内条款的行
            if not mapping.library:
                for col in range(6):
                    item = self.table.item(row, col)
                    if item:
                        item.setBackground(QColor(255, 255, 200))

        self.label_count.setText(f"共 {len(mappings)} 条映射")

    def _filter_table(self, text: str):
        """过滤表格"""
        text = text.lower()
        for row in range(self.table.rowCount()):
            match = False
            for col in range(3):  # 只搜索前三列
                item = self.table.item(row, col)
                if item and text in item.text().lower():
                    match = True
                    break
            self.table.setRowHidden(row, not match)

    def _show_context_menu(self, pos):
        """显示右键菜单"""
        menu = QMenu(self)

        action_add = menu.addAction("添加")
        action_add.triggered.connect(self._on_add)

        if self.table.selectedItems():
            action_edit = menu.addAction("编辑")
            action_edit.triggered.connect(self._on_edit)

            action_delete = menu.addAction("删除")
            action_delete.triggered.connect(self._on_delete)

            menu.addSeparator()

            action_copy = menu.addAction("复制英文名")
            action_copy.triggered.connect(self._on_copy_english)

        menu.exec_(self.table.mapToGlobal(pos))

    def _on_add(self):
        """添加映射"""
        dialog = MappingEditDialog(self, library_clauses=self.library_clauses)
        if dialog.exec_() == QDialog.Accepted and dialog.result:
            self.manager.add_mapping(
                english=dialog.result['english'],
                chinese=dialog.result['chinese'],
                library=dialog.result['library'],
                notes=dialog.result['notes'],
            )
            self._refresh_table()

    def _on_edit(self):
        """编辑映射"""
        selected = self.table.selectedItems()
        if not selected:
            return

        row = selected[0].row()
        english = self.table.item(row, 0).text()
        mapping = self.manager.get_mapping(english)

        if not mapping:
            return

        dialog = MappingEditDialog(self, mapping=mapping, library_clauses=self.library_clauses)
        if dialog.exec_() == QDialog.Accepted and dialog.result:
            # 如果英文名改变，先删除旧的
            if dialog.result['english'].lower() != english.lower():
                self.manager.delete_mapping(english)

            self.manager.add_mapping(
                english=dialog.result['english'],
                chinese=dialog.result['chinese'],
                library=dialog.result['library'],
                notes=dialog.result['notes'],
            )
            self._refresh_table()

    def _on_delete(self):
        """删除映射"""
        selected_rows = set(item.row() for item in self.table.selectedItems())
        if not selected_rows:
            return

        count = len(selected_rows)
        reply = QMessageBox.question(
            self, "确认删除",
            f"确定要删除选中的 {count} 条映射吗？",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            for row in sorted(selected_rows, reverse=True):
                english = self.table.item(row, 0).text()
                self.manager.delete_mapping(english)
            self._refresh_table()

    def _on_copy_english(self):
        """复制英文名到剪贴板"""
        from PyQt5.QtWidgets import QApplication
        selected = self.table.selectedItems()
        if selected:
            row = selected[0].row()
            english = self.table.item(row, 0).text()
            QApplication.clipboard().setText(english)

    def _on_import_json(self):
        """导入 JSON"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 JSON 文件", "",
            "JSON 文件 (*.json);;所有文件 (*.*)"
        )

        if not file_path:
            return

        dialog = ImportSettingsDialog(self, file_type="json")
        if dialog.exec_() != QDialog.Accepted:
            return

        added, updated, errors = self.manager.import_from_json(
            file_path,
            overwrite=dialog.result['overwrite']
        )

        if errors:
            QMessageBox.warning(self, "导入警告", "\n".join(errors))
        else:
            QMessageBox.information(
                self, "导入完成",
                f"新增 {added} 条，更新 {updated} 条"
            )

        self._refresh_table()

    def _on_import_excel(self):
        """导入 Excel"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel 文件", "",
            "Excel 文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )

        if not file_path:
            return

        dialog = ImportSettingsDialog(self, file_type="excel")
        if dialog.exec_() != QDialog.Accepted:
            return

        result = dialog.result
        added, updated, errors = self.manager.import_from_excel(
            file_path,
            overwrite=result['overwrite'],
            col_english=result['col_english'],
            col_chinese=result['col_chinese'],
            col_library=result['col_library'] if result['col_library'] >= 0 else 2,
            header_row=result['header_row']
        )

        if errors:
            QMessageBox.warning(self, "导入警告", "\n".join(errors))
        else:
            QMessageBox.information(
                self, "导入完成",
                f"新增 {added} 条，更新 {updated} 条"
            )

        self._refresh_table()

    def _on_export(self):
        """导出映射"""
        file_path, selected_filter = QFileDialog.getSaveFileName(
            self, "导出映射", "clause_mappings",
            "JSON 文件 (*.json);;Excel 文件 (*.xlsx)"
        )

        if not file_path:
            return

        if "xlsx" in selected_filter or file_path.endswith('.xlsx'):
            success = self.manager.export_to_excel(file_path)
        else:
            if not file_path.endswith('.json'):
                file_path += '.json'
            success = self.manager.export_to_json(file_path)

        if success:
            QMessageBox.information(self, "导出成功", f"已导出到:\n{file_path}")
        else:
            QMessageBox.warning(self, "导出失败", "导出时发生错误")

    def _on_save(self):
        """保存映射"""
        if self.manager.save():
            QMessageBox.information(self, "保存成功", "映射已保存")
            self.mappings_changed.emit()
        else:
            QMessageBox.warning(self, "保存失败", "保存时发生错误")

    def _on_close(self):
        """关闭对话框"""
        if self.manager.is_modified():
            reply = QMessageBox.question(
                self, "未保存的更改",
                "有未保存的更改，是否保存？",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel
            )

            if reply == QMessageBox.Save:
                self.manager.save()
                self.mappings_changed.emit()
            elif reply == QMessageBox.Cancel:
                return

        self.accept()


# ========================================
# 测试代码
# ========================================

if __name__ == "__main__":
    import sys
    from PyQt5.QtWidgets import QApplication

    app = QApplication(sys.argv)

    # 示例库内条款列表
    library = [
        "企业财产保险附加场外维修及改造条款",
        "企业财产保险附加地震扩展条款",
        "企业财产保险附加盗窃、抢劫保险条款",
        "企业财产保险附加自动恢复保险金额条款",
    ]

    dialog = ClauseMappingDialog(library_clauses=library)
    dialog.show()

    sys.exit(app.exec_())
