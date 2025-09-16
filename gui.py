from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox
import threading
from excel_processor import ExcelProcessor
from pdf_converter import PDFConverter


class ScheduleConverterGUI:
    def __init__(self, root):
        self.root = root
        self.setup_gui()
        
        # Инициализация обработчиков
        self.excel_processor = ExcelProcessor(gui_callback=self.update_label_text)
        self.pdf_converter = PDFConverter(gui_callback=self.update_label_text)
    
    def update_label_text(self, text):
        """Обновление текста метки из других потоков"""
        self.root.after(0, lambda: self.labelText.config(text=text))
    
    def setup_gui(self):
        """Настройка графического интерфейса"""
        self.root.title("Конвертер расписание Ректор-колледж в PDF")
        window_width = 500
        window_height = 358
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x_position = int((screen_width / 2) - (window_width / 2))
        y_position = int((screen_height / 2) - (window_height / 2))
        self.root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        # Элементы интерфейса
        file_path_label = Label(self.root, text="Файл Excel:", font=60)
        file_path_label.pack()
        self.file_path_entry = Entry(self.root, width=50, font=60)
        self.file_path_entry.pack()
        browse_file_button = Button(self.root, text="Обзор", command=self.browse_file, font=60, padx=50)
        browse_file_button.pack()

        save_file_path_label = Label(self.root, text="Куда сохранить(сгенерируются pdf-файлы):", font=60)
        save_file_path_label.pack()
        self.save_file_path_entry = Entry(self.root, width=50, font=60)
        self.save_file_path_entry.pack()
        browse_save_file_button = Button(self.root, text="Обзор", command=self.browse_save_path, font=60, padx=50)
        browse_save_file_button.pack() 

        self.remove_default_words = BooleanVar()
        self.remove_default_words.set(True)
        default_words_checkbox = Checkbutton(self.root, text="Удалить слова по умолчанию(лекция, вид занятия)", 
                                           variable=self.remove_default_words, font=60, padx=50)
        default_words_checkbox.pack()

        self.set_colors = BooleanVar()
        self.set_colors.set(True)
        set_color_checkbox = Checkbutton(self.root, text="Заполнить дни цветом", 
                                       variable=self.set_colors, font=60, padx=50)
        set_color_checkbox.pack()

        file_path_label = Label(self.root, text="Выберите тип расписания(для преподавателей или студентов):", font=60)
        file_path_label.pack()

        choices = ['Преподаватели', 'Группы']
        self.select_option_combobox = Combobox(self.root, values=choices, font=60, width=40)
        self.select_option_combobox.pack()

        word_remove_label = Label(self.root, text="Или напиши слова которые нужно удалить через дефис(-):", font=60)
        word_remove_label.pack()
        self.word_remove_entry = Entry(self.root, width=50, font=60)
        self.word_remove_entry.pack()  

        start_button = Button(self.root, text="Старт", font=60, padx=50)
        start_button.pack()
        start_button.config(command=self.run)  

        self.labelText = Label(self.root, text="", font=60, pady=20, padx=10, justify='left')
        self.labelText.pack()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def browse_file(self):
        """Выбор файла через диалоговое окно"""
        file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        self.file_path_entry.delete(0, END)
        self.file_path_entry.insert(END, file_path)
    
    def browse_save_path(self):
        """Выбор пути сохранения через диалоговое окно"""
        save_file_path = filedialog.askdirectory()
        self.save_file_path_entry.delete(0, END)
        self.save_file_path_entry.insert(END, save_file_path)
    
    def on_close(self):
        """Обработчик закрытия приложения"""
        result = messagebox.askquestion("Закрыть приложение", "Вы действительно хотите закрыть приложение?")
        if result == "yes":
            self.root.destroy()
    
    def run(self):
        """Запуск обработки в зависимости от выбранного типа"""
        if self.select_option_combobox.get() == 'Преподаватели':
            threading.Thread(target=self.get_teacher_schedule).start()
        else:
            threading.Thread(target=self.get_groups_schedule).start()
    
    def get_teacher_schedule(self):
        """Обработка расписания преподавателей"""
        file_path = self.file_path_entry.get().replace('//', '\\')
        updated_file = self.excel_processor.create_sheets_for_teacher(file_path)
        self.excel_processor.remove_empty_rows(updated_file)
        
        save_file_path = self.save_file_path_entry.get().replace('/', '\\')
        self.pdf_converter.convert_excel_to_pdf(updated_file, save_file_path, 'Преподаватели')
    
    def get_groups_schedule(self):
        """Обработка расписания групп"""
        file_path = self.file_path_entry.get().replace('//', '\\')
        updated_file = self.excel_processor.remove_empty_cells_and_words(
            file_path, 
            self.remove_default_words.get(), 
            self.set_colors.get(), 
            self.word_remove_entry.get()
        )
        
        save_file_path = self.save_file_path_entry.get().replace('/', '\\')
        self.pdf_converter.convert_excel_to_pdf(updated_file, save_file_path, 'Группы')