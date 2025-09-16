import openpyxl 
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font 
from openpyxl.utils import get_column_letter
import copy
from utils import format_fio


class ExcelProcessor:
    def __init__(self, gui_callback=None):
        self.gui_callback = gui_callback
    
    def update_gui_text(self, text):
        if self.gui_callback:
            self.gui_callback(text)
    
    def create_sheets_for_teacher(self, file_path):
        """Создание отдельных листов для каждого преподавателя"""
        workbook = openpyxl.load_workbook(file_path)
        new_workbook = openpyxl.Workbook()
        count_sheets = 0
        
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            max_row = sheet.max_row
            max_column = sheet.max_column
            new_sheet = new_workbook.create_sheet(sheet_name)
            new_sheet.bestFit = True
            
            self.update_gui_text(f"Копируем значения... {round((count_sheets/len(workbook.sheetnames)) * 100,1)}%") 
            count_sheets += 1
            
            for row in range(1, max_row + 1):
                cell_value = sheet.cell(row, 1).value
                if cell_value is not None and not isinstance(cell_value, int) and 'Преподаватель' in cell_value:
                    fio = cell_value.replace('Преподаватель -', '').strip()
                    formatted_fio = format_fio(fio)
                    new_sheet = new_workbook.create_sheet(formatted_fio)
                else:
                    for column in range(1, max_column + 1):
                        if row > 1:
                            cell = sheet.cell(row - 1, column)
                            new_cell = new_sheet.cell(row, column, value=cell.value)
                            new_cell.font = copy.copy(cell.font)
                            new_sheet.column_dimensions[get_column_letter(column)].width = sheet.column_dimensions[get_column_letter(column)].width

        new_workbook.remove(new_workbook['1'])
        new_workbook.remove(new_workbook['Sheet'])
        save_file = file_path.replace('.xlsx', '_updated.xlsx')
        new_workbook.save(save_file)
        new_workbook.close()
        return save_file

    def set_style_header(self, sheet):
        """Установка стилей для заголовков"""
        sheet.merge_cells('A1:M1')
        sheet.merge_cells('A2:M2')

    def remove_empty_rows(self, file_path):
        """Удаление пустых строк и форматирование"""
        workbook = openpyxl.load_workbook(file_path)
        border_style = Border(left=Side(style='thin'),
                              right=Side(style='thin'),
                              top=Side(style='thin'),
                              bottom=Side(style='thin'))
        font = Font(size=11, bold=True, name='Arial')
        count_sheets = 0
        
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            self.update_gui_text(f"Обработка листа {sheet_name} {round((count_sheets/len(workbook.sheetnames)) * 100,1)}%") 
            count_sheets += 1
            
            self.set_style_header(sheet)
            
            for row in reversed(range(1, sheet.max_row + 1)):
                if all(cell.value in (None, '', '-') for cell in sheet[row]):
                    sheet.delete_rows(row)

            sheet.merge_cells('A3:A4')
            sheet.merge_cells('C3:C4')
            sheet.merge_cells('E3:E4')
            sheet.merge_cells('G3:G4')
            sheet.merge_cells('I3:I4')
            sheet.merge_cells('K3:K4')
            sheet.merge_cells('M3:M4')
            
            sheet.cell(3, 3).value = 'Ауд'
            sheet.cell(3, 3).font = font
            sheet.cell(3, 5).value = 'Ауд'
            sheet.cell(3, 5).font = font
            sheet.cell(3, 7).value = 'Ауд'
            sheet.cell(3, 7).font = font
            sheet.cell(3, 9).value = 'Ауд'
            sheet.cell(3, 9).font = font
            sheet.cell(3, 11).value = 'Ауд'
            sheet.cell(3, 11).font = font
            sheet.cell(3, 13).value = 'Ауд'
            sheet.cell(3, 13).font = font

            for row in range(1, sheet.max_row + 1):
                if row > 2:  
                    for column in range(1, sheet.max_column + 1):
                        sheet.cell(row, column).border = border_style
                        sheet.cell(row, column).alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        
        workbook.save(file_path)
        workbook.close()

    def remove_empty_cells_and_words(self, file_path, remove_default_words, set_colors, word_remove_entry):
        """Удаление пустых ячеек и слов из расписания групп"""
        words_to_remove = ['(лекция)', '(практика)', ', вид занятия'] 
        workbook = openpyxl.load_workbook(file_path) 
        count_sheets = 0
        
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            max_row = sheet.max_row
            max_column = sheet.max_column
            total_sheets = len(workbook.sheetnames)
            
            if sheet.max_column > 4:
                sheet.merge_cells('A1:F1')
                sheet.merge_cells('A3:F3')
                sheet.column_dimensions['D'].width = 8.5
                sheet.column_dimensions['F'].width = 8.5
                sheet.column_dimensions['B'].width = 5.5
                sheet.column_dimensions['A'].width = 7
            else:
                sheet.merge_cells('A1:D1')
                sheet.merge_cells('A3:D3')
                sheet.column_dimensions['B'].width = 10
                sheet.column_dimensions['A'].width = 13
                sheet.column_dimensions['C'].width = 50.5
                sheet.column_dimensions['D'].width = 15.5

            sheet['A1'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            sheet['A1'].font = Font(size=18, bold=True)
            sheet.row_dimensions[2].hidden = True
            sheet.row_dimensions[3].hidden = True
            sheet.row_dimensions[4].hidden = True
        
            self.update_gui_text(f"Обработка листа {sheet_name} {round((count_sheets/total_sheets) * 100,1)}%") 
            count_sheets += 1
            
            for row in reversed(range(5, max_row + 1)): 
                for column in reversed(range(1, max_column + 1)):
                    cell = sheet.cell(row, column)
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    
                    if remove_default_words:
                        for word in words_to_remove:
                            if word.lower() in str(cell.value).lower():
                                cell.value = str(cell.value).replace(word, "") 
                    else:
                        user_words = word_remove_entry.split("-") 
                        for word in user_words:
                            if word.lower() in str(cell.value).lower():
                                cell.value = str(cell.value).replace(word, "")
                        
                    if max_column > 4:
                        if ((sheet[row][2].value is None or sheet[row][2].value == '')  
                        and (sheet[row][4].value is None or sheet[row][4].value == '')):    
                            sheet.row_dimensions[row].hidden = True 
                    else: 
                        if ((sheet[row][2].value is None or sheet[row][2].value == '') 
                        and (sheet[row][3].value is None or sheet[row][3].value == '')):
                            sheet.row_dimensions[row].hidden = True

                if set_colors:
                    if sheet[row][0].value == 'Пн':
                        sheet[row][0].fill = PatternFill(fill_type='solid', fgColor="D9E1F2")
                    if sheet[row][0].value == 'Вт':
                        sheet[row][0].fill = PatternFill(fill_type='solid', fgColor="FCE4D6")
                    if sheet[row][0].value == 'Ср':
                        sheet[row][0].fill = PatternFill(fill_type='solid', fgColor="FFF2CC")
                    if sheet[row][0].value == 'Чт':
                        sheet[row][0].fill = PatternFill(fill_type='solid', fgColor="E2EFDA")
                    if sheet[row][0].value == 'Пт':
                        sheet[row][0].fill = PatternFill(fill_type='solid', fgColor="D6DCE4")
                    if sheet[row][0].value == 'Сб':
                        sheet[row][0].fill = PatternFill(fill_type='solid', fgColor="FFB3B3")

        new_file_path = file_path.replace('.xlsx', '_updated.xlsx')
        workbook.save(new_file_path)
        workbook.close()
        return new_file_path