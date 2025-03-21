import os
import sys
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, 
    QFileDialog, QMessageBox, QHBoxLayout, QLabel,
    QInputDialog, QLineEdit, QScrollArea
)
from PyQt6.QtGui import QPixmap, QPainter, QFont, QColor, QImage
from PyQt6.QtCore import Qt
from pdf_utils import *  # 自定义PDF处理函数模块
import fitz  # PyMuPDF库，用于PDF渲染

class PDFToolboxApp(QWidget):
    """主应用程序类，继承自QWidget"""
    def __init__(self):
        super().__init__()
        self.init_ui()  # 初始化用户界面
        self.current_pdf_path = None  # 当前显示的PDF路径

    def init_ui(self):
        """初始化用户界面布局和组件"""
        # 窗口基本设置
        self.setWindowTitle("PDF 工具箱")
        self.setGeometry(100, 100, 900, 600)  # 设置窗口位置和大小

        # ===== 主布局 =====
        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(10, 10, 10, 10)  # 设置边距
        main_layout.setSpacing(15)  # 设置组件间距

        # ===== 左侧PDF预览区域 =====
        left_panel = QWidget()
        left_panel.setMinimumWidth(350)  # 设置最小宽度
        left_layout = QVBoxLayout(left_panel)
        
        # 预览标题
        lbl_preview_title = QLabel("PDF预览区域（支持多页滚动）")
        lbl_preview_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_preview_title.setStyleSheet("font-weight: bold; font-size: 14px;")
        left_layout.addWidget(lbl_preview_title)

        # 滚动区域容器
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)  # 允许内容自适应
        
        # 预览内容容器
        self.preview_container = QWidget()
        self.preview_layout = QVBoxLayout(self.preview_container)
        self.preview_layout.setSpacing(10)  # 设置页面间距
        
        self.scroll_area.setWidget(self.preview_container)
        left_layout.addWidget(self.scroll_area)

        main_layout.addWidget(left_panel)

        # ===== 右侧功能按钮区域 =====
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(20, 0, 20, 20)

        # 艺术字标题
        self.lbl_text = QLabel(right_panel)
        self.lbl_text.setFixedSize(300, 150)
        # 使用QPainter绘制动态文字
        pixmap = QPixmap(self.lbl_text.size())
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        font = QFont("Arial", 20, QFont.Weight.Bold)
        painter.setFont(font)
        painter.setPen(QColor(Qt.GlobalColor.red))
        painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, "PDF工具箱")
        painter.end()
        self.lbl_text.setPixmap(pixmap)
        right_layout.addWidget(self.lbl_text)

        # 功能按钮组
        self._init_buttons(right_layout)

        # 退出按钮
        exit_row = QHBoxLayout()
        exit_row.addStretch()  # 添加弹性空间
        self.btn_exit = QPushButton("退出")
        self.btn_exit.clicked.connect(self.close)
        exit_row.addWidget(self.btn_exit)
        right_layout.addLayout(exit_row)

        main_layout.addWidget(right_panel)
        self.setLayout(main_layout)

    def _init_buttons(self, layout):
        """初始化功能按钮"""
        # 按钮文本和对应功能列表
        buttons = [
            ("合并 PDF", self.merge_pdfs),
            ("拆分 PDF", self.split_pdf),
            ("加密 PDF", self.encrypt_pdf),
            ("解密 PDF", self.decrypt_pdf),
            ("提取 PDF 文本", self.extract_text),
            ("HTML 转 PDF", self.html_to_pdf),
            ("图片转 PDF", self.images_to_pdf),
            ("Word 转 PDF", self.word_to_pdf)
        ]

        # 两列布局生成
        button_rows = []
        for i in range(0, len(buttons), 2):
            row_buttons = buttons[i:i+2]
            button_row = QHBoxLayout()
            for text, callback in row_buttons:
                btn = QPushButton(text)
                btn.setStyleSheet("""
                    QPushButton {
                        padding: 12px;
                        font-size: 14px;
                        min-width: 120px;
                    }
                """)
                btn.clicked.connect(callback)
                button_row.addWidget(btn)
            button_rows.append(button_row)

        # 将按钮行添加到布局
        for row in button_rows:
            layout.addLayout(row)

    # --------------------------
    # 核心功能方法
    # --------------------------
    def clear_preview(self):
        """清空预览区域内容"""
        while self.preview_layout.count():
            child = self.preview_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

    def update_preview(self, file_path):
        """更新PDF预览
        Args:
            file_path (str): 要预览的PDF文件路径
        """
        self.clear_preview()
        self.current_pdf_path = file_path
        
        if not file_path:
            return

        try:
            doc = fitz.open(file_path)  # 使用PyMuPDF打开PDF
            container_width = self.scroll_area.viewport().width() - 20
            
            # 逐页渲染
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                page_rect = page.rect
                zoom = container_width / page_rect.width  # 动态计算缩放比例
                
                # 生成页面图像
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                qimage = QImage(
                    pix.samples, 
                    pix.width, 
                    pix.height, 
                    pix.stride,
                    QImage.Format_RGBA8888 if pix.alpha else QImage.Format.Format_RGB888
                )
                
                # 构建页面组件
                page_widget = QWidget()
                page_layout = QVBoxLayout(page_widget)
                
                # 添加页码标签
                page_number = QLabel(f"第 {page_num+1} 页")
                page_number.setStyleSheet("color: #666; font-size: 12px;")
                page_number.setAlignment(Qt.AlignmentFlag.AlignRight)
                
                # 添加页面图像
                page_label = QLabel()
                page_label.setPixmap(QPixmap.fromImage(qimage))
                page_label.setStyleSheet("background-color: white; border: 1px solid #ccc;")
                
                page_layout.addWidget(page_number)
                page_layout.addWidget(page_label)
                self.preview_layout.addWidget(page_widget)

            doc.close()
            self.preview_layout.addStretch()  # 添加底部弹性空间
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法加载PDF：{str(e)}")

    def merge_pdfs(self):
        """合并多个PDF文件"""
        try:
            # 文件选择对话框
            files, _ = QFileDialog.getOpenFileNames(
                self, 
                "选择多个 PDF 文件", 
                "", 
                "PDF Files (*.pdf)"
            )
            if not files:
                return
            
            # 保存路径选择
            output, _ = QFileDialog.getSaveFileName(
                self,
                "保存合并后的 PDF",
                "",
                "PDF Files (*.pdf)"
            )
            if output:
                merge_pdfs(files, output)  # 调用utils模块功能
                self.update_preview(output)
                QMessageBox.information(self, "成功", "PDF 合并成功！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"合并失败：{str(e)}")

    def split_pdf(self):
        try:
            file, _ = QFileDialog.getOpenFileName(self, "选择 PDF 文件", "", "PDF Files (*.pdf)")
            if not file: 
                return
            self.update_preview(file)
            start, end = self.get_page_range()
            if start is None or end is None: 
                return
            output, _ = QFileDialog.getSaveFileName(self, "保存拆分后的 PDF", "", "PDF Files (*.pdf)")
            if output:
                split_pdf(file, start, end, output)
                self.update_preview(output)  # 更新预览
                QMessageBox.information(self, "成功", f"已拆分第 {start}-{end} 页！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"拆分失败：{str(e)}")

    def encrypt_pdf(self):
        try:
            file, _ = QFileDialog.getOpenFileName(self, "选择 PDF 文件", "", "PDF Files (*.pdf)")
            if not file: 
                return
            password, ok = self.get_password("加密")
            if not ok or not password: 
                return
            output, _ = QFileDialog.getSaveFileName(self, "保存加密文件", "", "PDF Files (*.pdf)")
            if output:
                encrypt_pdf(file, password, output)
                self.update_preview(output)  # 更新预览
                QMessageBox.information(self, "成功", "加密成功！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加密失败：{str(e)}")

    def decrypt_pdf(self):
        try:
            file, _ = QFileDialog.getOpenFileName(self, "选择加密的 PDF", "", "PDF Files (*.pdf)")
            if not file: 
                return
            password, ok = self.get_password("解密")
            if not ok or not password: 
                return
            output, _ = QFileDialog.getSaveFileName(self, "保存解密文件", "", "PDF Files (*.pdf)")
            if output:
                decrypt_pdf(file, password, output)
                self.update_preview(output)  # 更新预览
                QMessageBox.information(self, "成功", "解密成功！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"解密失败：{str(e)}")

    def extract_text(self):
        try:
            file, _ = QFileDialog.getOpenFileName(self, "选择 PDF 文件", "", "PDF Files (*.pdf)")
            if not file: 
                return
            save_path, _ = QFileDialog.getSaveFileName(self,"保存Word文档","","Word Documents (*.docx)")
            if not save_path:
                return  # 用户取消保存
            extract_text_from_pdf(file, save_path)
            self.update_preview(file)  # 保持原有PDF预览
        except Exception as e:
            QMessageBox.critical(self, "错误", f"提取失败：{str(e)}")

    def html_to_pdf(self):
        try:
            file, _ = QFileDialog.getOpenFileName(
                self, "选择 HTML 文档", "", "HTML Files (*.html *.htm)"
            )
            if not file: 
                return
            output, _ = QFileDialog.getSaveFileName(self, "保存 PDF", "", "PDF Files (*.pdf)")
            if output:
                html_to_pdf(file, output)
                self.update_preview(output)  # 更新预览
                QMessageBox.information(self, "成功", "转换完成！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"转换失败：{str(e)}")

    def images_to_pdf(self):
        try:
            files, _ = QFileDialog.getOpenFileNames(
                self, "选择图片", "", "图片文件 (*.png *.jpg *.jpeg *.bmp)"
            )
            if not files: 
                return
            output, _ = QFileDialog.getSaveFileName(self, "保存 PDF", "", "PDF Files (*.pdf)")
            if output:
                images_to_pdf(files, output)
                self.update_preview(output)  # 更新预览
                QMessageBox.information(self, "成功", "PDF 生成成功！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成失败：{str(e)}")

    def word_to_pdf(self):
        try:
            file, _ = QFileDialog.getOpenFileName(
                self, "选择 Word 文档", "", "Word 文件 (*.docx *.doc)"
            )
            if not file: 
                return
            output, _ = QFileDialog.getSaveFileName(self, "保存 PDF", "", "PDF Files (*.pdf)")
            if output:
                if os.name == "nt":
                    word_to_pdf(file, output)
                else:
                    word_to_pdf_linux(file, output)
                self.update_preview(output)  
                QMessageBox.information(self, "成功", "转换完成！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"Word 转 PDF 失败：{str(e)}")

    def get_page_range(self):
        """获取用户输入的页码范围
        Returns:
            tuple: (起始页, 结束页) 或 (None, None)
        """
        try:
            start, ok1 = QInputDialog.getInt(
                self, 
                "起始页", 
                "输入起始页码（从1开始）:", 
                min=1, 
                max=1000
            )
            if not ok1: 
                return (None, None)
            
            end, ok2 = QInputDialog.getInt(
                self, 
                "结束页", 
                "输入结束页码:", 
                value=start, 
                min=start, 
                max=1000
            )
            return (start, end) if ok2 else (None, None)
        except:
            return (None, None)
    
    def get_password(self, title):
        """获取密码（带输入验证）"""
        password, ok = QInputDialog.getText(
            self, f"输入密码 - {title}", "密码:", 
            QLineEdit.EchoMode.Password
        )
        return password.strip(), ok

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PDFToolboxApp()
    window.show()
    sys.exit(app.exec())