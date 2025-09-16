import os
import win32com.client
from tkinter import messagebox


class PDFConverter:
    def __init__(self, gui_callback=None):
        self.gui_callback = gui_callback
        self.excel = None
    
    def update_gui_text(self, text):
        if self.gui_callback:
            self.gui_callback(text)
    
    def _setup_page_settings(self, sheet):
        """Настройка параметров страницы"""
        sheet.PageSetup.TopMargin = 1
        sheet.PageSetup.BottomMargin = 0
        sheet.PageSetup.RightMargin = 0
        sheet.PageSetup.LeftMargin = 0
        sheet.PageSetup.Zoom = False
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.CenterHorizontally = True
    
    def _get_sheet_name(self, sheet, schedule_type):
        """Получение имени для PDF файла"""
        if schedule_type != 'Преподаватели':
            if sheet.Cells(5, 5).Value not in (None, ''):
                return f"{sheet.Cells(5, 3).Value},{sheet.Cells(5, 5).Value}"
            return sheet.Cells(5, 3).Value
        else:
            # Пропуск пустых листов для преподавателей
            if all(sheet.Cells(5, col).Value is None for col in [1, 2, 3]):
                return None
            return sheet.Name
    
    def convert_excel_to_pdf(self, file_path, save_file_path, schedule_type):
        """Конвертация Excel в PDF"""
        try:
            self.excel = win32com.client.Dispatch("Excel.Application")
            workbook = self.excel.Workbooks.Open(file_path)
            sheets = workbook.Sheets
            valid_sheets = []
            
            # Предварительный сбор информации о листах
            for sheet in sheets:
                sheet_name = self._get_sheet_name(sheet, schedule_type)
                if sheet_name:
                    valid_sheets.append((sheet, sheet_name))
            
            total_sheets = len(valid_sheets)
            
            for count_sheets, (sheet, sheet_name) in enumerate(valid_sheets, 1):
                pdf_filename = f"{sheet_name}.pdf"
                self._setup_page_settings(sheet)
                
                sheet.ExportAsFixedFormat(0, os.path.join(save_file_path, pdf_filename))
                self.update_gui_text(f"Файл {pdf_filename} сохранен {round((count_sheets/total_sheets) * 100,1)}%")

            workbook.Close(SaveChanges=1)
            messagebox.showinfo(title='Успешно', message="Преобразование завершено") 
        
        except Exception as e:
            messagebox.showerror(title='Ошибка', message=f"Возникла ошибка при преобразовании: {e}")
            self._cleanup_resources(workbook)
        finally:
            self._cleanup_resources(workbook)
    
    def _cleanup_resources(self, workbook):
        """Очистка ресурсов Excel"""
        try:
            if workbook:
                workbook.Close(SaveChanges=0)
        except:
            pass
        
        try:
            if self.excel:
                self.excel.Quit()
        except:
            pass
        
        # Очистка COM объектов
        try:
            del self.excel
            del workbook
        except:
            pass