from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox
import threading
from excel_processor import ExcelProcessor
from pdf_converter import PDFConverter


class ScheduleConverterGUI:
    def __init__(self, root):
        self.root = root
        
        self.remove_default_words = BooleanVar(value=True)
        self.set_colors = BooleanVar(value=True)
        
        self.setup_gui()
        
        self.excel_processor = ExcelProcessor(gui_callback=self.update_label_text)
        self.pdf_converter = PDFConverter(gui_callback=self.update_label_text)
        
        # Установка значений по умолчанию
        self.select_option_combobox.set('Преподаватели')
    
    def update_label_text(self, text):
        self.root.after(0, lambda: self.labelText.config(text=text))
    
    def _create_label(self, text):
        return Label(self.root, text=text, font=('Arial', 10))
    
    def _create_entry(self, width=50):
        return Entry(self.root, width=width, font=('Arial', 10))
    
    def _create_button(self, text, command, padx=50):
        return Button(self.root, text=text, command=command, font=('Arial', 10), padx=padx)
    
    def _create_checkbox(self, text, variable):
        return Checkbutton(self.root, text=text, variable=variable, 
                          font=('Arial', 10), padx=50)
    
    def setup_gui(self):
        self.root.title("Конвертер расписание Ректор-колледж в PDF")
        self._center_window(500, 400)
        
        # Создание элементов интерфейса
        self._create_label("Файл Excel:").pack()
        self.file_path_entry = self._create_entry()
        self.file_path_entry.pack()
        self._create_button("Обзор", self.browse_file).pack()
        
        self._create_label("Куда сохранить (сгенерируются pdf-файлы):").pack()
        self.save_file_path_entry = self._create_entry()
        self.save_file_path_entry.pack()
        self._create_button("Обзор", self.browse_save_path).pack()
        
        # Чекбоксы с уже инициализированными переменными
        self._create_checkbox("Удалить слова по умолчанию (лекция, вид занятия)", 
                             self.remove_default_words).pack()
        self._create_checkbox("Заполнить дни цветом", 
                             self.set_colors).pack()
        
        self._create_label("Выберите тип расписания:").pack()
        self.select_option_combobox = Combobox(self.root, values=['Преподаватели', 'Группы'], 
                                              font=('Arial', 10), width=40)
        self.select_option_combobox.pack()
        
        self._create_label("Или напишите слова для удаления через дефис:").pack()
        self.word_remove_entry = self._create_entry()
        self.word_remove_entry.pack()
        
        self._create_button("Старт", self.run).pack()
        
        self.labelText = Label(self.root, text="", font=('Arial', 10), pady=20, padx=10, justify='left')
        self.labelText.pack()
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def _center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Выберите файл Excel"
        )
        if file_path:
            self.file_path_entry.delete(0, END)
            self.file_path_entry.insert(END, file_path)
    
    def browse_save_path(self):
        save_path = filedialog.askdirectory(title="Выберите папку для сохранения")
        if save_path:
            self.save_file_path_entry.delete(0, END)
            self.save_file_path_entry.insert(END, save_path)
    
    def on_close(self):
        if messagebox.askyesno("Закрыть приложение", "Вы действительно хотите закрыть приложение?"):
            self.root.destroy()
    
    def _validate_inputs(self):
        """Проверка корректности введенных данных"""
        if not self.file_path_entry.get():
            messagebox.showerror("Ошибка", "Выберите файл Excel")
            return False
        
        if not self.save_file_path_entry.get():
            messagebox.showerror("Ошибка", "Выберите папку для сохранения")
            return False
        
        return True
    
    def run(self):
        if not self._validate_inputs():
            return
        
        schedule_type = self.select_option_combobox.get()
        if schedule_type == 'Преподаватели':
            threading.Thread(target=self.get_teacher_schedule, daemon=True).start()
        else:
            threading.Thread(target=self.get_groups_schedule, daemon=True).start()
    
    def get_teacher_schedule(self):
        try:
            file_path = self.file_path_entry.get().replace('//', '\\')
            updated_file = self.excel_processor.create_sheets_for_teacher(file_path)
            self.excel_processor.remove_empty_rows(updated_file)
            
            save_path = self.save_file_path_entry.get().replace('/', '\\')
            self.pdf_converter.convert_excel_to_pdf(updated_file, save_path, 'Преподаватели')
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Ошибка", f"Ошибка обработки: {e}"))
    
    def get_groups_schedule(self):
        try:
            file_path = self.file_path_entry.get().replace('//', '\\')
            updated_file = self.excel_processor.remove_empty_cells_and_words(
                file_path, 
                self.remove_default_words.get(), 
                self.set_colors.get(), 
                self.word_remove_entry.get()
            )
            
            save_path = self.save_file_path_entry.get().replace('/', '\\')
            self.pdf_converter.convert_excel_to_pdf(updated_file, save_path, 'Группы')
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Ошибка", f"Ошибка обработки: {e}"))