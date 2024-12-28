import docx
from docx.shared import Pt
import docx.styles
import pandas as pd
import sys  
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog, QTableWidgetItem
import mainform_
import random
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
#from docx.enum import WD_LIST_NUMBERING

class testgen(QtWidgets.QMainWindow, mainform_.Ui_MainWindow):
    
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.action_load_file.triggered.connect(self.open_file_dialog)
        self.action_testgen.triggered.connect(self.gen)
        self.action_loadheader.triggered.connect(self.loadHeader)
        self.action_loadfooter.triggered.connect(self.loadFooter)
        self.radioButton_1col.toggled.connect(lambda: self.on_radio_button_clicked(self.radioButton_1col))
        self.radioButton_2col.toggled.connect(lambda: self.on_radio_button_clicked(self.radioButton_2col))
        self.spinBox.valueChanged.connect(lambda: self.checkValue(self.spinBox, self.spinBox_2))
        self.spinBox_2.valueChanged.connect(lambda: self.checkValue(self.spinBox, self.spinBox_2))

    def read_docx_file(self, file_path):
        document = docx.Document(file_path)
        text = []
        for paragraph in document.paragraphs:
            text.append(paragraph.text)
        return '\n'.join(text)

    def loadHeader(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open File", "", "All Files (*)")
        global header_content
        if file_path:
            header_content = self.read_docx_file(file_path)
            self.textBrowser.setText(header_content)
    
    def loadFooter(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open File", "", "All Files (*)")
        global footer_content
        if file_path:
            footer_content = self.read_docx_file(file_path)
            self.textBrowser_2.setText(footer_content)

    def on_radio_button_clicked(self, rbtn):
        global text_col #количество колонок в тесте
        if rbtn.isChecked():
            text_col = int(rbtn.text())
            

    def checkValue(self, spin1, spin2):
        if (spin1.value() == 0 or spin2.value() == 0):
            self.btn_testgen.setEnabled(False)
        else:
            self.btn_testgen.setEnabled(True)
    
    def open_file_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open File", "", "All Files (*)")
        if file_path:
            df = pd.read_excel(file_path)
            global dict_list 
            dict_list= df.to_dict(orient='records')
            table = self.tableWidget
            global row_count
            global column_count
            row_count = (len(dict_list))
            column_count = (len(dict_list[0]))
            table.setColumnCount(column_count) 
            table.setRowCount(row_count)
           
            table.setHorizontalHeaderLabels((list(dict_list[0].keys())))
            table.setColumnWidth(0, 200)
            table.setColumnWidth(1, 600)
            table.setColumnWidth(2, 200)
            table.setColumnWidth(3, 200)
            table.setColumnWidth(4, 200)
            table.setColumnWidth(5, 200)
            table.setColumnWidth(6, 200)
            
            for row in range(row_count): 
                for column in range(column_count):
                    item = list(dict_list[row].values())[column]
                    if ( pd.isna(item)):
                        item = ''
                    table.setItem(row, column, QTableWidgetItem(item))
                    
        else:
            QtWidgets.QMessageBox.critical(self, 'Error', 'Ошибка чтения файла', QtWidgets.QMessageBox.Yes)

    def add_header_section(self, doc, test, style):
        section = doc.sections[0]
        section.start_type = WD_SECTION.NEW_PAGE
        sectPr = section._sectPr 
        cols = sectPr.xpath('./w:cols')[0]
        cols.set(qn('w:num'), '1')
        if self.headergroup.isChecked():
            head = doc.add_paragraph(header_content, style = style)
        head=doc.add_paragraph('Вариант '+str(test+1), style = style)
        
    def add_main_section(self, doc, style, test):
        LEFT_INDENT = Pt(36)
        

    def add_footer_section(self, doc):
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(10)
        sectionf = doc.add_section(WD_SECTION.CONTINUOUS)
        sectionf = doc.sections[2]
        sectPr = sectionf._sectPr
        cols = sectPr.xpath('./w:cols')[0]
        cols.set(qn('w:num'), '1')
        if self.footergroup.isChecked():
            foot = doc.add_paragraph(footer_content).style = style
            #foot.style = style

    def add_keys_table(self, doc, question_count, list_keys):
        # добавляем таблицу с одной строкой для заполнения названий колонок
        table = doc.add_table(1, question_count+1)
        table.style = 'Table Grid'
        # Получаем строку с колонками из добавленной таблицы 
        head_cells = table.rows[0].cells
        #print(list_keys)
        # добавляем названия колонок
        p = head_cells[0].paragraphs[0]
        p.add_run(f'Вариант').bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        for i in range(1, question_count+1):
            #print(i)
            p = head_cells[i].paragraphs[0]
            p.add_run(f'Вопрос {i}').bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for row in range(len(list_keys)):
            cells = table.add_row().cells
            cells[0].text = f'Вариант {row+1}'
            for col in range(1,len(list_keys[row])+1):
                cells[col].text = str(list_keys[row][col-1])

    def gen(self):
        LEFT_INDENT = Pt(36)
        outfile_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "All Files (*)")
              
        if outfile_path:
            doc = docx.Document()
            # задаем стиль текста по умолчанию
            style = doc.styles['Normal']
            style.font.name = 'Arial'
            style.font.size = Pt(14)
            #стиль для шапки
            styleH = doc.styles.add_style('MyHeaderStyle', WD_STYLE_TYPE.PARAGRAPH)
            styleH.font.name = 'Arial'
            styleH.font.size = Pt(10)
            #styleH.font.underline = True
            styleH.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            #_______________________________
            #стиль теста
            style_main = doc.styles.add_style('mainStyle', WD_STYLE_TYPE.PARAGRAPH)
            style_main.font.name = 'Arial'
            style_main.font.size = Pt(12)
            style_main.paragraph_format.algnment = WD_ALIGN_PARAGRAPH.JUSTIFY
            #_______________________________
            test_count = self.spinBox.value()
            question_count = self.spinBox_2.value()
            list_keys= []
            if (test_count != 0 and question_count != 0):
                
                for test in range(test_count):
                    self.add_header_section( doc, test, styleH)
                    section_number = test+1                  
                    doc.add_section(WD_SECTION.CONTINUOUS)
                    section1 = doc.sections[section_number]
                    sectPr = section1._sectPr
                    cols = sectPr.xpath('./w:cols')[0]
                    cols.set(qn('w:num'), str(text_col))
                    cols.set(qn('w:space'), '10')  # Set space between columns to 10 points ->0.01"

                    list_ques = []
                    for ques in range(question_count):
                        n = random.randint(1, row_count-1)
                        ques_text = doc.add_paragraph('Вопрос №' + str(ques+1), style=style_main)
                        ques_text.alignment = WD_ALIGN_PARAGRAPH.CENTER

                        item = list(dict_list[n].values())
                       
                        doc.add_paragraph(item[1], style=style_main)
                        for i in range(2,6):
                            if (pd.isna(item[i]) != True):
                                par =doc.add_paragraph(f'{i-1}. '+ item[i], style=style_main)
                                                        
                        list_ques.insert(ques, item[6])
                        
                        doc.add_paragraph('________________________________')
                    self.add_footer_section(doc)
                    section_number= section_number+1                  
                    list_keys.insert(test, list_ques)
                    doc.add_page_break()
                # Добавление таблицы
                self.add_keys_table(doc, question_count, list_keys)
                
            elif (test_count == 0 or question_count == 0):
                self.btn_testgen.setEnabled(False)     
                                                        
            doc.save(outfile_path)
            doc.styles['MyHeaderStyle'].delete()
            doc.styles['mainStyle'].delete()
            
            QtWidgets.QMessageBox.information(self, 'Information', 'Файл сформирован', QtWidgets.QMessageBox.Yes)
        else: QtWidgets.QMessageBox.critical(self, 'Error', 'Ошибка чтения файла', QtWidgets.QMessageBox.Yes)
     
def main():
    app = QtWidgets.QApplication(sys.argv) 
    window = testgen()  
    window.show() 
    app.exec_() 

if __name__ == '__main__':  
    main()  