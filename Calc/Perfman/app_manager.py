import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
from datetime import datetime
from .settings import Settings  # Updated to relative import
from .timer_d import Timer  # Updated to relative import
from .button_manager import ButtonManager  # Updated to relative import
from .mode_manager import ModeManager  # Updated to relative import
from .file_handler import FileHandler  # Updated to relative import

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Менеджер производительности by v.lazarenko")
        # Инициализация переменных
        self.total_value = 0
        self.button_history = []
        self.settings = Settings()
        self.timer1 = Timer()
        self.timer2 = Timer()
        self.button_manager = ButtonManager(self)
        self.mode_manager = ModeManager(self)
        self.file_handler = FileHandler()
        # Загрузка настроек
        self.settings.load_settings()
        # Создание интерфейса
        self.create_tabs()
        # Создание меню
        self.create_menu()

    def create_tabs(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(expand=True, fill='both')
        input_data_frame = ttk.Frame(notebook)
        notebook.add(input_data_frame, text="Ввод данных")
        self.create_input_data_tab(input_data_frame)
        current_day_frame = ttk.Frame(notebook)
        notebook.add(current_day_frame, text="Текущий день")
        self.create_current_day_tab(current_day_frame)

    def create_input_data_tab(self, frame):
        ttk.Label(frame, text="Логин").grid(row=1, column=0, padx=5, pady=5)
        self.login_entry = ttk.Entry(frame)
        self.login_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Label(frame, text="Всего рабочих дней").grid(row=2, column=0, padx=5, pady=5)
        self.total_days_entry = ttk.Entry(frame)
        self.total_days_entry.grid(row=2, column=1, padx=5, pady=5)
        ttk.Label(frame, text="Осталось рабочих дней").grid(row=3, column=0, padx=5, pady=5)
        self.remaining_days_entry = ttk.Entry(frame)
        self.remaining_days_entry.grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Загрузить данные", command=self.load_data).grid(row=4, column=0, columnspan=2, padx=5, pady=5)

    def create_current_day_tab(self, frame):
        ttk.Label(frame, text="Кэф на конец месяца/Бонус(в баллах):").grid(row=0, column=0, padx=5, pady=5)
        self.index_value_month_label = ttk.Label(frame, text="0")
        self.index_value_month_label.grid(row=0, column=1, padx=5, pady=5)
        self.bonus_label = ttk.Label(frame, text="0")
        self.bonus_label.grid(row=0, column=2, padx=5, pady=5)
        ttk.Label(frame, text="До следующего бонуса/Следующий бонус(в баллах):").grid(row=2, column=0, padx=5, pady=5)
        self.bonus_goal_label = ttk.Label(frame, text="0")
        self.bonus_goal_label.grid(row=2, column=1, padx=5, pady=5)
        self.next_bonus_label = ttk.Label(frame, text="0")
        self.next_bonus_label.grid(row=2, column=2, padx=5, pady=5)
        ttk.Label(frame, text="До нормы(месяц):").grid(row=3, column=0, padx=5, pady=5)
        self.monthly_goal_label = ttk.Label(frame, text="0")
        self.monthly_goal_label.grid(row=3, column=1, padx=5, pady=5)
        ttk.Label(frame, text="До нормы(день):").grid(row=4, column=0, padx=5, pady=5)
        self.daily_goal_label = ttk.Label(frame, text="0")
        self.daily_goal_label.grid(row=4, column=1, padx=5, pady=5)
        current_date = datetime.now().strftime("%d.%m.%Y")
        ttk.Label(frame, text=f"Сделано {current_date}:").grid(row=5, column=0, padx=5, pady=5)
        self.done_today_label = ttk.Label(frame, text="0")
        self.done_today_label.grid(row=5, column=1, padx=5, pady=5)
        self.index_value_month = 0
        self.bonus = 0
        self.bonus_goal = 0
        self.next_bonus = 0
        self.monthly_goal = 0
        self.daily_goal = 0
        self.done_today = 0
        self.sum_points = 0
        self.norm_value_month = 0
        self.value_now = 0
        self.login = ''
        activities_frame = ttk.LabelFrame(frame, text="Активности")
        activities_frame.grid(row=6, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.button_manager.create_buttons(activities_frame)
        ttk.Button(frame, text="Отменить последнее действие", command=self.button_manager.undo_last_action).grid(row=7, column=0, columnspan=2, padx=5, pady=5)
        modes_frame = ttk.LabelFrame(frame, text="Режимы")
        modes_frame.grid(row=8, column=0, columnspan=2, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.mode_manager.create_modes(modes_frame)

    def load_data(self):
        login = self.login_entry.get()
        total_days = int(self.total_days_entry.get()) if self.total_days_entry.get().isdigit() else None
        remaining_days = int(self.remaining_days_entry.get()) if self.remaining_days_entry.get().isdigit() else None
        if not login or total_days is None or remaining_days is None:
            messagebox.showerror("Ошибка", "Пожалуйста, заполните все поля корректными данными.")
            return
        try:
            data = self.file_handler.read_from_xlsx(login, total_days, remaining_days)
            if data:
                self.update_data_fields(data)
            else:
                messagebox.showerror("Ошибка", "Данные не найдены для указанного логина.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при загрузке данных: {str(e)}")

    def update_data_fields(self, data):
        self.login = data['login']
        self.daily_goal = data['daily_goal'] - self.done_today
        self.monthly_goal = data['monthly_goal'] - self.done_today
        self.bonus = data['bonus']
        self.bonus_goal = data['bonus_goal'] - self.done_today
        self.sum_points = data['sum_points'] + self.done_today
        self.norm_value_month = data['norm_value_month']
        self.index_value_month = float(self.sum_points / self.norm_value_month) if self.norm_value_month > 0 else 0
        self.next_bonus = data['next_bonus']
        self.update_goals_labels(self.daily_goal, self.monthly_goal, data['bonus'], self.bonus_goal, self.index_value_month, data['next_bonus'])

    def update_goals_labels(self, daily_goal, monthly_goal, bonus, bonus_goal, index_value_month, next_bonus):
        self.daily_goal_label.config(text=f"{max(0, int(daily_goal))}")
        self.monthly_goal_label.config(text=f"{max(0, int(monthly_goal))}")
        self.index_value_month_label.config(text=(f"{(index_value_month)}")[:5])
        self.bonus_label.config(text=f"{max(0, int(bonus))}")
        self.bonus_goal_label.config(text=f"{max(0, int(bonus_goal))}")
        self.next_bonus_label.config(text=f"{max(0, int(next_bonus))}")

    def create_menu(self):
        menu_bar = tk.Menu(self.root)
        self.root.config(menu=menu_bar)
        settings_menu = tk.Menu(menu_bar, tearoff=0)
        settings_menu.add_command(label="Добавить новую кнопку", command=self.open_add_button_window)
        settings_menu.add_command(label="Удалить существующую кнопку", command=self.open_remove_button_window)
        settings_menu.add_separator()
        settings_menu.add_command(label="Добавить новый режим", command=self.open_add_mode_window)
        settings_menu.add_command(label="Удалить существующий режим", command=self.open_remove_mode_window)
        settings_menu.add_separator()
        settings_menu.add_command(label="Сброс настроек", command=self.reset_settings)
        settings_menu.add_command(label="Сохранить изменения", command=self.save_settings_changes)
        menu_bar.add_cascade(label="Настройки", menu=settings_menu)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Сохранить данные в XLSX", command=self.save_data_to_xlsx)
        menu_bar.add_cascade(label="Файл", menu=file_menu)

    def open_add_button_window(self):
        add_window = tk.Toplevel(self.root)
        add_window.title("Добавить новую кнопку")
        ttk.Label(add_window, text="Название кнопки").grid(row=0, column=0, padx=5, pady=5)
        name_entry = ttk.Entry(add_window)
        name_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(add_window, text="Значение").grid(row=1, column=0, padx=5, pady=5)
        value_entry = ttk.Entry(add_window)
        value_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(add_window, text="Добавить", command=lambda: self.add_button_and_update(name_entry.get(), float(value_entry.get()), add_window)).grid(row=2, column=0, columnspan=2, padx=5, pady=5)

    def add_button_and_update(self, name, value, window):
        if self.mode_manager.is_any_mode_active():
            messagebox.showwarning("Предупреждение", "Нельзя добавить кнопку, пока включен режим.")
            return
        if len(self.settings.buttons) >= 10:
            messagebox.showwarning("Предупреждение", "Превышено максимальное количество кнопок (10).")
            return
        self.settings.add_button(name, value)
        self.button_manager.reload_buttons()
        window.destroy()

    def open_remove_button_window(self):
        remove_window = tk.Toplevel(self.root)
        remove_window.title("Удалить кнопку")
        ttk.Label(remove_window, text="Выберите кнопку для удаления").grid(row=0, column=0, padx=5, pady=5)
        button_var = tk.StringVar(remove_window)
        button_options = list(self.settings.buttons.keys())
        button_dropdown = ttk.OptionMenu(remove_window, button_var, None, *button_options)
        button_dropdown.grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(remove_window, text="Удалить", command=lambda: self.remove_button_and_update(button_var.get(), remove_window)).grid(row=2, column=0, padx=5, pady=5)

    def remove_button_and_update(self, button_name, window):
        if button_name:
            self.settings.remove_button(button_name)
            self.button_manager.reload_buttons()
            messagebox.showinfo("Успех", f"Кнопка '{button_name}' удалена.")
            window.destroy()
        else:
            messagebox.showwarning("Предупреждение", "Выберите кнопку для удаления.")

    def reset_settings(self):
        self.settings.reset_settings()
        self.button_manager.reload_buttons()
        self.mode_manager.reload_modes()
        messagebox.showinfo("Сброс настроек", "Настройки успешно сброшены!")

    def save_settings_changes(self):
        self.settings.save_settings()
        self.mode_manager.save_modes()
        messagebox.showinfo("Сохранение изменений", "Изменения успешно сохранены.")

    def collect_current_day_data(self):
        login = self.login
        monthly_goal = self.monthly_goal_label.cget("text")
        daily_goal = self.daily_goal_label.cget("text")
        done_today = self.done_today_label.cget("text")
        index_value_month = self.index_value_month_label.cget("text")
        bonus = self.bonus_label.cget("text")
        bonus_goal = self.bonus_goal_label.cget("text")
        next_bonus = self.next_bonus_label.cget("text")
        sum_points = self.sum_points
        activities_data = []
        for name, button_data in self.button_manager.buttons.items():
            counter = button_data['counter_label'].cget("text")
            cumulative = button_data['cumulative']
            activities_data.append({
                'name': name,
                'counter': counter,
                'cumulative': cumulative
            })
        modes_data = []
        for name, mode_data in self.mode_manager.modes.items():
            timer = mode_data['timer'].visible_timer_label.cget("text")
            modes_data.append({
                'name': name,
                'timer': timer,
            })
        return {
            'login': login,
            'monthly_goal': monthly_goal,
            'daily_goal': daily_goal,
            'done_today': done_today,
            'index_value_month': index_value_month,
            'bonus': bonus,
            'bonus_goal': bonus_goal,
            'next_bonus': next_bonus,
            'sum_points': sum_points,
            'activities': activities_data,
            'modes': modes_data
        }

    def save_data_to_xlsx(self):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel files", "*.xlsx")],
                                                initialfile=f"report_{datetime.now().strftime('%d.%m.%Y')}.xlsx")
        if filename:
            data = self.collect_current_day_data()
            wb = openpyxl.Workbook()
            ws = wb.active
            cur_day = datetime.now().strftime('%d.%m.%Y')
            headers = [
                "Логин",
                "Месячный план",
                "Дневной план",
                f"Сделано {cur_day}",
                "Кэф на конец месяцa",
                "Бонус за перевыполнение",
                "Баллов до следующего бонуса",
                "Следующий бонус",
                "Всего баллов(по менеджеру)"
            ]
            ws.append(headers)
            ws.append([
                data['login'],
                data['monthly_goal'],
                data['daily_goal'],
                data['done_today'],
                data['index_value_month'],
                data['bonus'],
                data['bonus_goal'],
                data['next_bonus'],
                data['sum_points']
            ])
            ws.append([])
            ws.append(["Название активности", "Количество", "Накопительное значение"])
            for activity in data['activities']:
                ws.append([activity['name'], activity['counter'], activity['cumulative']])
            ws.append([])
            ws.append(["Название режима", "Времени в режиме"])
            for mode in data['modes']:
                ws.append([mode['name'], mode['timer']])
            wb.save(filename)
            messagebox.showinfo("Сохранение данных", f"Данные сохранены в {filename}")

    def open_add_mode_window(self):
        add_window = tk.Toplevel(self.root)
        add_window.title("Добавить новый режим")
        ttk.Label(add_window, text="Название режима").grid(row=0, column=0, padx=5, pady=5)
        name_entry = ttk.Entry(add_window)
        name_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(add_window, text="Баллов в минуту").grid(row=1, column=0, padx=5, pady=5)
        value_entry = ttk.Entry(add_window)
        value_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(add_window, text="Добавить", command=lambda: self.add_mode_and_update(name_entry.get(), value_entry.get(), add_window)).grid(row=2, column=0, columnspan=2, padx=5, pady=5)

    def add_mode_and_update(self, name, value_str, window):
        if self.mode_manager.is_any_mode_active():
            messagebox.showwarning("Предупреждение", "Нельзя добавить режим, пока включен режим.")
            return
        value = int(value_str) if value_str.isdigit() else 0
        if len(self.settings.modes) >= 6:
            messagebox.showwarning("Предупреждение", "Превышено максимальное количество режимов (6).")
            return
        self.settings.add_mode(name, value)
        self.mode_manager.reload_modes()
        window.destroy()

    def open_remove_mode_window(self):
        remove_window = tk.Toplevel(self.root)
        remove_window.title("Удалить режим")
        ttk.Label(remove_window, text="Выберите режим для удаления").grid(row=0, column=0, padx=5, pady=5)
        mode_var = tk.StringVar(remove_window)
        mode_options = list(self.settings.modes.keys())
        mode_dropdown = ttk.OptionMenu(remove_window, mode_var, None, *mode_options)
        mode_dropdown.grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(remove_window, text="Удалить", command=lambda: self.remove_mode_and_update(mode_var.get(), remove_window)).grid(row=2, column=0, padx=5, pady=5)

    def remove_mode_and_update(self, mode_name, window):
        if mode_name:
            if self.mode_manager.is_any_mode_active():
                messagebox.showwarning("Предупреждение", "Нельзя удалить режим, пока включен режим.")
                return
            self.settings.remove_mode(mode_name)
            self.mode_manager.reload_modes()
            if mode_name in self.mode_manager.timers:
                del self.mode_manager.timers[mode_name]
            if mode_name in self.mode_manager.modes:
                del self.mode_manager.modes[mode_name]
            messagebox.showinfo("Успех", f"Режим '{mode_name}' удален.")
            window.destroy()
        else:
            messagebox.showwarning("Предупреждение", "Выберите режим для удаления.")
