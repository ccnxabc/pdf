#1、这是一个针对港股招股书，只筛选需要的，凑成一个新的pdf文件。
#2、将该py文件生产exe。
#3、处理过的pdf文件然后丢到chatpdf分析，免费版本的有100页限制。但是太高估chatpdf
#来自chatgpt，部分涉及页面的需要自己修改
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QCheckBox, QPushButton, QVBoxLayout, QWidget, QScrollArea, QHBoxLayout, QLabel, QFileDialog,QMessageBox,QLineEdit

import fitz # PyMuPDF

class PDFEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.pdf_document = None
        self.checkboxes = []
        self.initUI()

    def initUI(self):
        # 设置窗口标题和大小
        self.setWindowTitle('PDF Bookmark Editor')
        self.setGeometry(100, 100, 800, 600)

        # 创建一个水平布局
        main_layout = QHBoxLayout()

        # 创建一个垂直布局用于复选框
        self.checkbox_layout = QVBoxLayout()

        # 创建滚动区域并设置布局
        scroll_area = QScrollArea()
        scroll_area_widget = QWidget()
        scroll_area.setWidget(scroll_area_widget)
        scroll_area.setWidgetResizable(True)
        scroll_area_widget.setLayout(self.checkbox_layout)

        # 创建一个垂直布局用于操作按钮
        operation_layout = QVBoxLayout()

        # 创建选择文件按钮
        select_file_btn = QPushButton('选择PDF文件', self)
        select_file_btn.clicked.connect(self.selectPDFFile)
        operation_layout.addWidget(select_file_btn)

        # 创建保存按钮
        save_btn = QPushButton('保存修改', self)
        save_btn.clicked.connect(self.savePDF)
        operation_layout.addWidget(save_btn)

        # 创建输入框以输入页码范围
        self.page_range_input = QLineEdit(self)
        self.page_range_input.setPlaceholderText('输入页码，如：1-3,5,8-9,将在勾选框精简的文件后添加')
        operation_layout.addWidget(self.page_range_input)

        # 将滚动区域和操作布局添加到主布局
        main_layout.addWidget(scroll_area)
        main_layout.addLayout(operation_layout)

        # 创建一个容器，并设置布局
        container = QWidget()
        container.setLayout(main_layout)

        # 设置主窗口的中心部件
        self.setCentralWidget(container)

    def selectPDFFile(self):
        # 打开文件选择对话框
        file_name, _ = QFileDialog.getOpenFileName(self, '选择PDF文件', '', 'PDF files (*.pdf)')
        if file_name:
            self.pdf_path = file_name
            self.pdf_document = fitz.open(self.pdf_path)
            self.updateCheckboxes()

    def updateCheckboxes(self):
        # 清除现有的复选框
        for checkbox in self.checkboxes:
            checkbox.deleteLater()
        self.checkboxes.clear()

        # 读取PDF书签并为每个书签创建一个复选框
        if self.pdf_document:
            for bookmark in self.pdf_document.get_toc():
                checkbox = QCheckBox(f"{bookmark[1]} (页码: {bookmark[2]})")
                self.checkboxes.append(checkbox)
                self.checkbox_layout.addWidget(checkbox)

    def savePDF(self):

        if self.pdf_document:
            #该文档书签总数
            ShuQian_ZongShu=len(self.checkboxes)
            #指示查到哪个书签
            Which_ShuQian=0
            #该pdf文档有多少页
            self.total_pages=self.pdf_document.page_count

            # 使用文件对话框来选择保存路径和输入文件名
            new_pdf_path, _ = QFileDialog.getSaveFileName(self, '保存PDF文件', '', 'PDF files (*.pdf)')
            if new_pdf_path:
                # 创建一个新的PDF文档
                new_pdf = fitz.open()
                # 遍历复选框，如果被选中，则将对应页码的页面添加到新PDF中
                for checkbox in self.checkboxes:
                    if checkbox.isChecked():
                        page_num = int(checkbox.text().split(': ')[-1].strip(')'))

                        if Which_ShuQian<ShuQian_ZongShu-1:
                            pre_next_shuqian_page=int(self.checkboxes[Which_ShuQian+1].text().split(': ')[-1].strip(')'))-1
                        else:
                            pre_next_shuqian_pagee=self.total_pages

                        new_pdf.insert_pdf(self.pdf_document, from_page=page_num - 1, to_page=pre_next_shuqian_page-1)

                    Which_ShuQian=Which_ShuQian+1

                # 获取用户输入的页码范围,例如：1-3,5,8-9,
                #首先输入要有内容，没内容就往下操作
                if len(self.page_range_input.text())>0:
                    page_ranges = self.page_range_input.text().split(',')
                    selected_pages = set()
                    if len(page_ranges)>0:
                        for page_range in page_ranges:
                            if '-' in page_range:
                                start, end = map(int, page_range.split('-'))
                                selected_pages.update(range(start, end + 1))
                            else:
                                selected_pages.add(int(page_range))

                        # 根据用户输入的页码范围添加页面(例如：1-3,5,8-9)
                        #是否要判断selected_pages是否为空？
                        for page_num in sorted(selected_pages):
                            if 1 <= page_num <= self.total_pages:
                                new_pdf.insert_pdf(self.pdf_document, from_page=page_num - 1, to_page=page_num - 1)


                # 保存新的PDF文件
                new_pdf.save(new_pdf_path)
                new_pdf.close()
                self.pdf_document.close()
                print(f"新文件已保存为: {new_pdf_path}")

                # 弹出确认对话框
                QMessageBox.information(self, '保存成功', f'文件已成功保存到：{new_pdf_path}', QMessageBox.Ok)

# 运行程序
if __name__ == '__main__':
    app = QApplication(sys.argv)
    pdf_editor = PDFEditor()
    pdf_editor.show()
    sys.exit(app.exec_())
