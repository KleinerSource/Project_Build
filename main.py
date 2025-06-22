#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
航发属性模板桌面应用程序
使用PySide5开发的Excel数据处理工具，兼容Windows 7系统
"""

import sys
import os
import pandas as pd
from PySide5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
                               QWidget, QLabel, QLineEdit, QListWidget, QScrollArea,
                               QTextEdit, QMessageBox, QGroupBox, QRadioButton,
                               QCheckBox, QButtonGroup, QFrame, QComboBox, QListWidgetItem,
                               QSplitter)
from PySide5.QtCore import Qt, Signal
from PySide5.QtGui import QFont



class PropertyConfigApp(QMainWindow):
    """航发属性配置主应用程序"""
    
    def __init__(self):
        super().__init__()
        self.df = None
        self.current_category = None
        self.property_widgets = {}  # 存储属性控件
        self.init_ui()
        self.load_excel_file()
    
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("欧菲斯航发属性配置工具")
        self.setGeometry(100, 100, 1200, 800)

        # 创建中央窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局（垂直布局，包含标题、内容区域、版权信息）
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # 标题
        title_label = QLabel("欧菲斯航发属性配置工具")
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("padding: 10px; background-color: #f0f0f0; border-radius: 5px;")
        main_layout.addWidget(title_label)

        # 内容区域（水平布局，左右分栏）
        content_layout = QHBoxLayout()
        content_layout.setSpacing(15)

        # 左侧区域：分类搜索
        self.create_left_panel(content_layout)

        # 右侧区域：属性配置
        self.create_right_panel(content_layout)

        main_layout.addLayout(content_layout)

        # 版权信息
        copyright_label = QLabel("版权所有 © KleinerSource")
        copyright_label.setAlignment(Qt.AlignCenter)
        copyright_label.setStyleSheet("color: gray; font-size: 10px; margin-top: 10px;")
        main_layout.addWidget(copyright_label)

    def create_left_panel(self, parent_layout):
        """创建左侧面板（分类搜索区域）"""
        # 左侧容器
        left_widget = QWidget()
        left_widget.setFixedWidth(350)
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(5, 5, 5, 5)

        # 分类搜索区域
        search_group = QGroupBox("分类名称搜索")
        search_layout = QVBoxLayout(search_group)

        # 搜索框
        search_label = QLabel("搜索分类名称:")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("输入关键词搜索分类名称...")
        self.search_input.textChanged.connect(self.filter_categories)

        # 分类名称列表
        self.property_list = QListWidget()
        self.property_list.itemClicked.connect(self.on_category_selected)

        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.property_list)

        left_layout.addWidget(search_group)
        parent_layout.addWidget(left_widget)

    def create_right_panel(self, parent_layout):
        """创建右侧面板（属性配置区域）"""
        # 右侧容器
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(5, 5, 5, 5)

        # 使用垂直分割器来分配属性配置区域和结果显示区域的空间
        splitter = QSplitter(Qt.Vertical)
        splitter.setChildrenCollapsible(False)  # 防止子区域被完全折叠

        # 属性配置区域
        config_widget = self.create_property_config_area()
        splitter.addWidget(config_widget)

        # 结果显示区域
        result_widget = self.create_result_area()
        splitter.addWidget(result_widget)

        # 设置分割器的比例：属性配置区域占85%，结果显示区域占15%
        splitter.setSizes([850, 150])  # 总和为1000，比例为85:15
        splitter.setStretchFactor(0, 1)  # 属性配置区域可以拉伸
        splitter.setStretchFactor(1, 0)  # 结果显示区域不拉伸

        right_layout.addWidget(splitter)
        parent_layout.addWidget(right_widget)


    def create_property_config_area(self):
        """创建属性配置区域"""
        config_group = QGroupBox("属性配置")
        config_layout = QVBoxLayout(config_group)

        # 滚动区域 - 移除最小高度限制，让它可以充分利用可用空间
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        # 移除 setMinimumHeight(300) 以允许更灵活的空间分配

        # 配置内容容器
        self.config_widget = QWidget()
        self.config_layout = QVBoxLayout(self.config_widget)
        # 优化布局间距，使属性条目更加紧凑
        self.config_layout.setSpacing(5)  # 减少属性条目之间的垂直间距
        self.config_layout.setContentsMargins(5, 5, 5, 5)  # 减少容器内边距
        self.scroll_area.setWidget(self.config_widget)

        config_layout.addWidget(self.scroll_area)
        return config_group
    
    def create_result_area(self):
        """创建结果显示区域"""
        result_group = QGroupBox("配置结果")
        result_layout = QVBoxLayout(result_group)

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        # 设置固定高度，只显示3行文本（约75像素）
        self.result_text.setFixedHeight(75)
        self.result_text.setPlaceholderText("选择属性后，配置结果将在此显示...")

        # 设置字体和样式以确保3行文本的可读性
        font = self.result_text.font()
        font.setPointSize(9)  # 稍微小一点的字体以适应3行显示
        self.result_text.setFont(font)

        # 设置文本编辑器的样式
        self.result_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 4px;
                background-color: #f9f9f9;
            }
        """)

        result_layout.addWidget(self.result_text)
        return result_group
    
    def load_excel_file(self):
        """加载Excel文件"""
        excel_file = "航发属性模板.xlsx"
        
        if not os.path.exists(excel_file):
            QMessageBox.critical(self, "错误", f"找不到文件: {excel_file}")
            sys.exit(1)
        
        try:
            self.df = pd.read_excel(excel_file)
            # 获取所有唯一的分类名称
            self.load_categories()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载Excel文件失败: {str(e)}")
            sys.exit(1)

    def load_categories(self):
        """加载所有分类名称到列表中"""
        if self.df is not None:
            # 获取唯一的分类名称（B列）
            unique_categories = self.df['分类名称'].dropna().unique()
            self.property_list.clear()
            for category in sorted(unique_categories):
                self.property_list.addItem(str(category))
    
    def filter_categories(self, text):
        """根据搜索文本过滤分类名称"""
        if self.df is None:
            return

        self.property_list.clear()
        if not text.strip():
            self.load_categories()
            return

        # 搜索包含关键词的分类名称
        unique_categories = self.df['分类名称'].dropna().unique()
        filtered_categories = [category for category in unique_categories if text.lower() in str(category).lower()]

        for category in sorted(filtered_categories):
            self.property_list.addItem(str(category))
    
    def on_category_selected(self, item):
        """当选择分类名称时的处理"""
        self.current_category = item.text()
        self.load_property_config()

    def load_property_config(self):
        """加载选中分类名称的配置"""
        if not self.current_category or self.df is None:
            return

        # 清除之前的配置
        self.clear_config_area()
        self.property_widgets.clear()

        # 获取当前分类名称的所有记录
        category_data = self.df[self.df['分类名称'] == self.current_category]

        # 只显示必填属性（H列为"是"）
        required_properties = category_data[category_data['是否必填'] == '是']

        if required_properties.empty:
            no_required_label = QLabel("该分类没有必填属性")
            no_required_label.setAlignment(Qt.AlignCenter)
            self.config_layout.addWidget(no_required_label)
            return

        # 为每个必填属性创建输入控件
        for _, row in required_properties.iterrows():
            self.create_property_widget(row)

        # 添加弹性空间
        self.config_layout.addStretch()

        # 更新结果显示
        self.update_result_display()

    def create_property_widget(self, row):
        """为单个属性创建输入控件"""
        property_name = row['属性名称']
        property_type = row['属性类型']
        property_values = row['属性值']

        # 创建属性行容器（水平布局）
        property_frame = QFrame()
        property_frame.setFrameStyle(QFrame.StyledPanel)
        property_frame.setLineWidth(1)
        # 减少属性框架的内边距，使布局更紧凑
        property_frame.setContentsMargins(3, 1, 3, 1)  # 从(8,5,8,5)减少到(6,3,6,3)
        property_layout = QHBoxLayout(property_frame)
        property_layout.setSpacing(8)  # 从15减少到12，保持适当的控件间距

        # 左侧：属性名称标签（固定宽度）
        name_label = QLabel(f"{property_name} (必填)")
        name_label.setStyleSheet("""
            font-weight: bold;
            color: #d32f2f;
            padding: 1px 3px;
            background-color: #fafafa;
            border-radius: 3px;
        """)
        name_label.setFixedWidth(160)
        name_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        name_label.setWordWrap(True)
        # 设置标签的固定高度，与控件保持一致
        name_label.setMinimumHeight(26)
        name_label.setMaximumHeight(26)
        property_layout.addWidget(name_label)

        # 右侧：根据属性类型创建不同的输入控件
        if property_type == "文本框" or pd.isna(property_type):
            # 文本输入框
            text_input = QLineEdit()
            text_input.setPlaceholderText(f"请输入{property_name}")
            text_input.textChanged.connect(self.update_result_display)
            # 减少控件高度，使布局更紧凑
            text_input.setMinimumHeight(26)  # 从30减少到26
            text_input.setMaximumHeight(26)  # 设置最大高度保持一致性
            # 确保文本框可以获得焦点
            text_input.setFocusPolicy(Qt.StrongFocus)
            self.property_widgets[property_name] = text_input
            property_layout.addWidget(text_input)

        elif property_type == "单选菜单":
            # 单选下拉列表
            if pd.notna(property_values):
                combo_box = QComboBox()
                # 减少下拉列表高度，与文本框保持一致
                combo_box.setMinimumHeight(26)  # 从30减少到26
                combo_box.setMaximumHeight(26)  # 设置最大高度保持一致性
                combo_box.addItem("请选择...")  # 默认选项

                options = [opt.strip() for opt in str(property_values).split(',')]
                for option in options:
                    combo_box.addItem(option)

                combo_box.currentTextChanged.connect(self.update_result_display)
                self.property_widgets[property_name] = combo_box
                property_layout.addWidget(combo_box)

        elif property_type == "多选菜单":
            # 单选下拉列表（原多选菜单改为单选）
            if pd.notna(property_values):
                combo_box = QComboBox()
                # 减少下拉列表高度，与其他控件保持一致
                combo_box.setMinimumHeight(26)  # 从30减少到26
                combo_box.setMaximumHeight(26)  # 设置最大高度保持一致性
                combo_box.addItem("请选择...")  # 默认选项

                options = [opt.strip() for opt in str(property_values).split(',')]
                for option in options:
                    combo_box.addItem(option)

                combo_box.currentTextChanged.connect(self.update_result_display)
                self.property_widgets[property_name] = combo_box
                property_layout.addWidget(combo_box)

        self.config_layout.addWidget(property_frame)

    def clear_config_area(self):
        """清除配置区域的所有控件"""
        while self.config_layout.count():
            child = self.config_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

    def update_result_display(self):
        """更新结果显示"""
        if not self.property_widgets:
            self.result_text.clear()
            return

        results = []

        for property_name, widget in self.property_widgets.items():
            value = ""

            if isinstance(widget, QLineEdit):
                # 文本输入框
                value = widget.text().strip()
            elif isinstance(widget, QComboBox):
                # 单选下拉列表（包括原多选菜单）
                current_text = widget.currentText()
                if current_text and current_text != "请选择...":
                    value = current_text
            elif isinstance(widget, list):
                # 兼容旧的单选按钮和复选框（如果还有的话）
                if widget and isinstance(widget[0], QRadioButton):
                    # 单选按钮组
                    for radio_btn in widget:
                        if radio_btn.isChecked():
                            value = radio_btn.text()
                            break
                elif widget and isinstance(widget[0], QCheckBox):
                    # 复选框组
                    selected = [cb.text() for cb in widget if cb.isChecked()]
                    value = ",".join(selected)

            if value:
                results.append(f"{property_name}={value}")

        # 显示结果
        result_text = "|".join(results)
        self.result_text.setPlainText(result_text)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PropertyConfigApp()
    window.show()
    sys.exit(app.exec_())  # PySide5使用exec_()而不是exec()
