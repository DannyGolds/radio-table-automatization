from comparing import Comparing
import customtkinter as ctk
from tkinter import filedialog, messagebox


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Сравнение таблиц Excel")
        self.geometry("600x380")
        
        # Создаем фреймы для организации интерфейса
        self.frame = ctk.CTkFrame(self)
        self.frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        # Заголовок
        self.label = ctk.CTkLabel(self.frame, text="Сравнение таблиц Excel", font=ctk.CTkFont(size=20, weight="bold"))
        self.label.pack(pady=12, padx=10)
        
        # Поля для выбора файлов
        self.file1_label = ctk.CTkLabel(self.frame, text="Перетащите или выберите через меню первый файл", compound="center")
        self.file1_label.pack(pady=(0, 5))
        
        self.file1_entry = ctk.CTkEntry(self.frame, width=400)
        self.file1_entry.pack(pady=(0, 5))
        
        self.file1_button = ctk.CTkButton(self.frame, text="Выбрать", command=lambda: self.select_file(self.file1_entry))
        self.file1_button.pack(pady=(0, 10))
        
        self.file2_label = ctk.CTkLabel(self.frame, text="Перетащите или выберите через меню второй файл", compound="center")
        self.file2_label.pack(pady=(0, 5))
        
        self.file2_entry = ctk.CTkEntry(self.frame, width=400)
        self.file2_entry.pack(pady=(0, 5))
        
        self.file2_button = ctk.CTkButton(self.frame, text="Выбрать", command=lambda: self.select_file(self.file2_entry))
        self.file2_button.pack(pady=(0, 20))
        
        # Кнопка сравнения
        self.compare_button = ctk.CTkButton(self.frame, text="Сравнить таблицы", command=self.run_comparison)
        self.compare_button.pack(pady=10)
        
        # Статус бар
        self.status_label = ctk.CTkLabel(self.frame, text="", text_color="gray")
        self.status_label.pack(pady=10)
    
    def select_file(self, entry_widget):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
        if file_path:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, file_path)
    
    def run_comparison(self):
        file1 = self.file1_entry.get()
        file2 = self.file2_entry.get()

        if not file1 or not file2:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите оба файла!")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="Результат.xlsx"
        )

        if not output_path:
            return  # Пользователь отменил сохранение

        compare_selected_columns = Comparing(file1, file2, ['C', 'D', 'H', 'I', 'K'], output_path)
        result, message = compare_selected_columns.compare()
        if result:
            result, message = compare_selected_columns.save()
            messagebox.showinfo("Успех", message)
        else:
            messagebox.showerror("Ошибка", message)
        self.update()

if __name__ == "__main__":
    app = App()
    app.mainloop()