from tkinter import Tk
from gui import ScheduleConverterGUI
import sys

def main():
    try:
        root = Tk()
        app = ScheduleConverterGUI(root)
        root.mainloop()
    except Exception as e:
        print(f"Ошибка запуска приложения: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()