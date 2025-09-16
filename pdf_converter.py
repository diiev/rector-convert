import os
import win32com.client
from tkinter import messagebox


class PDFConverter:
    def __init__(self, gui_callback=None):
        self.gui_callback = gui_callback
    
    def update_gui_text(self, text):
        if self.gui_callback:
            self.gui_callback(text)
    
    def convert_excel_to_pdf(self, file_path, save_file_path, schedule_type):
        """Конвертация Excel в PDF"""
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            workbook = excel.Workbooks.Open(file_path)
            count_sheets_continue = 0 
            count_sheets = 0

            for sheet in workbook.Sheets:
                if schedule_type != 'Преподаватели':
                    if sheet.Cells(5, 5).Value != '' and sheet.Cells(5, 5).Value != None:
                        cell_value = str(sheet.Cells(5, 3).Value) + ',' + str(sheet.Cells(5, 5).Value) 
                    else:
                        cell_value = sheet.Cells(5, 3).Value
                else: 
                    if sheet.Cells(5, 1).Value is None and sheet.Cells(5, 2).Value is None and sheet.Cells(5, 3).Value is None:
                        count_sheets_continue += 1
                        continue 
                    cell_value = sheet.Name
            
                pdf_filename = f"{cell_value}.pdf"
                sheet.PageSetup.TopMargin = 1
                sheet.PageSetup.BottomMargin = 0
                sheet.PageSetup.RightMargin = 0
                sheet.PageSetup.LeftMargin = 0
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.CenterHorizontally = True
                
                sheet.ExportAsFixedFormat(0, os.path.join(save_file_path, pdf_filename))
                count_sheets += 1
                self.update_gui_text(f"Файл {pdf_filename} сохранен {round((count_sheets/(len(workbook.Sheets) - count_sheets_continue)) * 100,1)}%")

            workbook.Close(SaveChanges=1)
            messagebox.showinfo(title='Успешно', message="Преобразование завершено") 
        
        except Exception as e:
            messagebox.showerror(title='Ошибка', message=f"Возникла ошибка при преобразовании: {e}")
            if 'workbook' in locals():
                workbook.Close(SaveChanges=0)
            if 'excel' in locals():
                excel.Quit()
            excel.Quit()    
            del excel
            del workbook