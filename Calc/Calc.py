import tkinter as tk
from tkinter import messagebox, ttk
import openpyxl  # Для работы с файлами Excel
import threading
import os
import re

# Глобальные переменные для хранения состояния таймера и значения "Обработано сегодня"
paid_mode_timer = None
break_timer_running = False
break_timer_seconds = 0
processed_today_value = 0
paid_mode_points = 0
button_counters = {2: 0, 4: 0, 5.5: 0, 9: 0, 10: 0, 4: 0, 20: 0}  # Исправлено значение для "Эскалация" на 4
button_widgets = {}
button_counter_entries = {}
total_additions_entries = {}
break_timer_label = None

# Функция для обработки нажатия кнопки на первой вкладке
def process_file():
    login = login_entry.get()
    total_days_str = total_days_entry.get()
    remaining_days_str = remaining_days_entry.get()
    if not total_days_str.isdigit() or int(total_days_str) < 0 or int(total_days_str) > 31:
        messagebox.showwarning("Предупреждение", "Введите корректное количество общего числа рабочих дней (целое число от 0 до 31).")
        return
    if not remaining_days_str.isdigit() or int(remaining_days_str) < 0 or int(remaining_days_str) > 31:
        messagebox.showwarning("Предупреждение", "Введите корректное количество оставшихся рабочих дней (целое число от 0 до 31).")
        return
    total_days = int(total_days_str)
    remaining_days = int(remaining_days_str)
    if login:
        try:
            norm_value_day, norm_value_month = load_norm_from_file(login, total_days, remaining_days)
            if norm_value_day is not None and norm_value_month is not None:
                norm_entry_day.config(state=tk.NORMAL)  # Разрешаем редактирование, чтобы установить значение
                norm_entry_day.delete(0, tk.END)
                norm_entry_day.insert(0, str(round(norm_value_day)))  # Округляем до целого числа
                norm_entry_day.config(state='readonly')  # Запрещаем редактирование после установки значения
                norm_entry_month.config(state=tk.NORMAL)  # Разрешаем редактирование, чтобы установить значение
                norm_entry_month.delete(0, tk.END)
                norm_entry_month.insert(0, str(round(norm_value_month)))  # Округляем до целого числа
                norm_entry_month.config(state='readonly')  # Запрещаем редактирование после установки значения
                messagebox.showinfo("Информация", f"Обрабатывается файл для логина: {login}")
            else:
                messagebox.showwarning("Предупреждение", "Логин не найден в файле или данные некорректны.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при обработке файла: {e}")
    else:
        messagebox.showwarning("Предупреждение", "Введите логин")

# Функция для увеличения значения поля "Обработано сегодня" и уменьшения значений "До нормы"
def increment_processed_today(value):
    global processed_today_value
    processed_today_value += value
    update_processed_today_display()
    # Увеличиваем счетчик нажатий для соответствующей кнопки
    button_counters[value] += 1
    update_button_counter_entries()
    # Обновляем сумму добавлений для текущей кнопки
    total_additions_entries[value].config(state=tk.NORMAL)
    current_total = float(total_additions_entries[value].get() or '0')
    total_additions_entries[value].delete(0, tk.END)
    total_additions_entries[value].insert(0, str(current_total + value))
    total_additions_entries[value].config(state='readonly')
    # Уменьшаем значения "До нормы"
    decrease_norm_values(value)

# Функция для уменьшения значений "До нормы"
def decrease_norm_values(value):
    current_norm_day = float(norm_entry_day.get())
    current_norm_month = float(norm_entry_month.get())
    new_norm_day = max(0, current_norm_day - value)
    new_norm_month = max(0, current_norm_month - value)
    norm_entry_day.config(state=tk.NORMAL)
    norm_entry_day.delete(0, tk.END)
    norm_entry_day.insert(0, str(round(new_norm_day)))
    norm_entry_day.config(state='readonly')
    norm_entry_month.config(state=tk.NORMAL)
    norm_entry_month.delete(0, tk.END)
    norm_entry_month.insert(0, str(round(new_norm_month)))
    norm_entry_month.config(state='readonly')

# Функция для обновления отображения значения "Обработано сегодня"
def update_processed_today_display():
    processed_today_entry.config(state=tk.NORMAL)
    processed_today_entry.delete(0, tk.END)
    processed_today_entry.insert(0, str(int(processed_today_value)))
    processed_today_entry.config(state='readonly')

# Функция для обновления значений счетчиков кнопок
def update_button_counter_entries():
    for value, entry in button_counter_entries.items():
        entry.config(state=tk.NORMAL)
        entry.delete(0, tk.END)
        entry.insert(0, str(button_counters[value]))
        entry.config(state='readonly')

# Функция для загрузки значений "До нормы(день)" и "До нормы(месяц)" из файла по логину
def load_norm_from_file(login, total_days, remaining_days):
    try:
        # Поиск файла по шаблону имени
        pattern = r'Мотивация ООО.*\.xlsx'
        for filename in os.listdir('.'):
            if re.match(pattern, filename):
                workbook = openpyxl.load_workbook(filename)
                sheet = workbook['Данные за период']
                # Получаем заголовки столбцов
                headers = [cell.value for cell in sheet[1]]
                print(f"Заголовки столбцов: {headers}")
                # Проверяем наличие необходимых столбцов
                required_columns = ['Логин', 'График', 'Сумма баллов']
                missing_columns = [col for col in required_columns if col not in headers]
                if missing_columns:
                    messagebox.showerror("Ошибка", f"Отсутствуют необходимые столбцы: {', '.join(missing_columns)}")
                    return None, None
                # Находим индексы столбцов
                login_index = headers.index('Логин')
                grafik_index = headers.index('График')
                points_per_hour_index = headers.index('Баллы в час')
                sum_points_index = headers.index('Сумма баллов')
                # Проходим по строкам таблицы, чтобы найти соответствующий логин (игнорируя регистр)
                login_lower = login.lower()
                found = False
                for row in sheet.iter_rows(min_row=2, values_only=True):  # Начиная со второй строки, пропуская заголовки
                    if row[login_index] and row[login_index].lower() == login_lower:
                        grafik = row[grafik_index]
                        points_per_hour = row[points_per_hour_index]
                        sum_points = row[sum_points_index]
                        if points_per_hour and points_per_hour != 0:
                            # Проверка наличия подстроки "8" или "2/2" в значении графика
                            if '8' in str(grafik):
                                norm_value_day = ((480 * total_days) - sum_points) / remaining_days if remaining_days > 0 else 0
                                norm_value_month = (480 * total_days) - sum_points
                            elif '2/2' in str(grafik):
                                norm_value_day = ((620 * total_days) - sum_points) / remaining_days if remaining_days > 0 else 0
                                norm_value_month = (620 * total_days) - sum_points
                            else:
                                messagebox.showwarning("Предупреждение", f"Неизвестное значение графика: {grafik}")
                                return None, None
                            found = True
                            break
                        else:
                            messagebox.showwarning("Предупреждение", f"Нулевое или отсутствующее значение в столбце 'Баллы в час' для логина: {login}")
                            return None, None
                if not found:
                    messagebox.showwarning("Предупреждение", "Логин не найден в файле")
                    return None, None
                return norm_value_day, norm_value_month
        messagebox.showwarning("Предупреждение", "Файл не найден.")
        return None, None
    except FileNotFoundError:
        messagebox.showerror("Ошибка", "Файл не найден.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при загрузке файла: {e}")
    return None, None

# Функция для запуска главного цикла приложения в отдельном потоке
def run_app():
    global root, break_timer_label, break_button, paid_mode_button
    # Создаем главное окно
    root = tk.Tk()
    root.title("Calc_V")
    root.geometry("800x700")  # Увеличиваем высоту окна

    # Настройка расширяемости строк и столбцов для всего окна
    root.grid_columnconfigure(0, weight=1)
    root.grid_rowconfigure(0, weight=1)

    # Создаем notebook (вкладки)
    notebook = ttk.Notebook(root)
    notebook.grid(row=0, column=0, sticky="nsew")

    # Первая вкладка - Ввод данных
    tab1 = ttk.Frame(notebook)
    notebook.add(tab1, text="Ввод данных")

    # Элементы для первой вкладки
    login_label = tk.Label(tab1, text="Введите логин:")
    login_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
    global login_entry
    login_entry = tk.Entry(tab1)
    login_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
    
    total_days_label = tk.Label(tab1, text="Всего раб. дней в месяце(без отпусков, БЛ и т.д):")
    total_days_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
    global total_days_entry
    total_days_entry = tk.Entry(tab1)
    total_days_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
    
    remaining_days_label = tk.Label(tab1, text="Осталось раб. дней включая сегодня(введите значение):")
    remaining_days_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
    global remaining_days_entry
    remaining_days_entry = tk.Entry(tab1)
    remaining_days_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
    
    process_button = tk.Button(tab1, text="Обработать файл", command=process_file)
    process_button.grid(row=3, column=0, columnspan=2, padx=5, pady=10, sticky="ew")
    
    # Настройка расширяемости строк и столбцов для первой вкладки
    tab1.grid_columnconfigure(1, weight=1)

    # Вторая вкладка - Текущий день
    tab2 = ttk.Frame(notebook)
    notebook.add(tab2, text="Текущий день")

    # Поле "До нормы(месяц)"
    global norm_entry_month
    norm_label_month = tk.Label(tab2, text="До нормы(месяц):")
    norm_label_month.grid(row=0, column=0, padx=10, pady=10, sticky="w")
    norm_entry_month = tk.Entry(tab2, state='readonly')  # Только для чтения
    norm_entry_month.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

    # Поле "До нормы(день)"
    global norm_entry_day
    norm_label_day = tk.Label(tab2, text="До нормы(день):")
    norm_label_day.grid(row=1, column=0, padx=10, pady=10, sticky="w")
    norm_entry_day = tk.Entry(tab2, state='readonly')  # Только для чтения
    norm_entry_day.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

    # Поле "Обработано сегодня"
    global processed_today_entry
    processed_today_label = tk.Label(tab2, text="Обработано сегодня:")
    processed_today_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")
    processed_today_entry = tk.Entry(tab2, state='readonly')  # Только для чтения
    processed_today_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
    processed_today_entry.config(state=tk.NORMAL)
    processed_today_entry.insert(0, "0")  # Устанавливаем начальное значение
    processed_today_entry.config(state='readonly')

    # Кнопки для увеличения значения "Обработано сегодня"
    buttons = {
        2: "Дубль/клиент отказался",
        4: "Переделка",
        5.5: "Касание(кроме По номеру/С2С)",
        9: "По номеру/С2С(только закрытие)",
        10: "Результат",
        4: "Эскалация(кроме АТМ 11/13)",  # Исправлено значение на +4
        20: "Согласование Jira"  # Добавлена новая кнопка
    }
    global button_widgets
    button_widgets = {}
    global button_counter_entries
    button_counter_entries = {}
    global total_additions_entries
    total_additions_entries = {}
    row_num = 3
    # Подписываем столбцы
    activity_count_label = tk.Label(tab2, text="Кол-во активностей")
    activity_count_label.grid(row=row_num, column=1, padx=10, pady=5, sticky="w")
    points_label = tk.Label(tab2, text="Баллов за активность")
    points_label.grid(row=row_num, column=2, padx=10, pady=5, sticky="w")
    row_num += 1
    for value, label in buttons.items():
        btn = tk.Button(tab2, text=label, command=lambda v=value: increment_processed_today(v))
        btn.grid(row=row_num, column=0, padx=10, pady=5, sticky="w")
        button_widgets[value] = btn
        counter_entry = tk.Entry(tab2, width=5, state='readonly')
        counter_entry.grid(row=row_num, column=1, padx=10, pady=5, sticky="ew")
        button_counter_entries[value] = counter_entry
        total_addition_entry = tk.Entry(tab2, width=5, state='readonly')
        total_addition_entry.grid(row=row_num, column=2, padx=10, pady=5, sticky="ew")
        total_additions_entries[value] = total_addition_entry
        row_num += 1

    # Кнопка "Платный режим"
    global paid_mode_var
    paid_mode_var = tk.BooleanVar()
    paid_mode_button = tk.Checkbutton(tab2, text="Платный режим", variable=paid_mode_var, command=toggle_paid_mode,
                                      font=("Helvetica", 14))  # Увеличенный размер шрифта
    paid_mode_button.grid(row=row_num, column=0, padx=10, pady=10, sticky="w")
    
    # Поле для отображения баллов, добавленных платным режимом
    paid_mode_points_label = tk.Label(tab2, text="Баллы, добавленные платным режимом:")
    paid_mode_points_label.grid(row=row_num, column=1, padx=10, pady=10, sticky="w")
    global paid_mode_points_entry
    paid_mode_points_entry = tk.Entry(tab2, state='readonly')
    paid_mode_points_entry.grid(row=row_num, column=2, padx=10, pady=10, sticky="ew")
    paid_mode_points_entry.config(state=tk.NORMAL)
    paid_mode_points_entry.insert(0, "0")
    paid_mode_points_entry.config(state='readonly')
    row_num += 1

    # Кнопка "Сброс" перемещена в правый верхний угол и изменен цвет фона на красный
    reset_button = tk.Button(tab2, text="Сброс", command=reset_values, bg="red", fg="white")
    reset_button.grid(row=0, column=2, padx=10, pady=10, sticky="ne")  # Позиционирование кнопки в правом верхнем углу

    # Кнопка "Перерыв"
    global break_timer_running, break_timer_label
    break_timer_running = False
    break_timer_label = tk.Label(tab2, text="00:00:00", font=("Helvetica", 16))  # Увеличенный размер шрифта
    break_timer_label.grid(row=row_num, column=0, padx=10, pady=10, sticky="w")
    break_button = tk.Button(tab2, text="Перерыв", command=toggle_break_timer, width=20, height=2,
                             font=("Helvetica", 14))  # Увеличенный размер кнопки и шрифта
    break_button.grid(row=row_num, column=1, padx=10, pady=10, sticky="ew")
    row_num += 1

    # Настройка расширяемости строк и столбцов для второй вкладки
    tab2.grid_columnconfigure(0, weight=1)
    tab2.grid_columnconfigure(1, weight=1)
    tab2.grid_columnconfigure(2, weight=1)

    # Упаковываем notebook
    notebook.pack(expand=True, fill="both")

    # Запускаем главный цикл приложения
    root.mainloop()

# Функция для увеличения значения "Обработано сегодня" в платном режиме и уменьшения значений "До нормы"
def increment_paid_mode():
    global processed_today_value, paid_mode_points
    processed_today_value += 1
    paid_mode_points += 1
    update_processed_today_display()
    update_paid_mode_points_display()
    # Уменьшаем значения "До нормы"
    decrease_norm_values(1)

# Функция для обновления отображения баллов, добавленных платным режимом
def update_paid_mode_points_display():
    paid_mode_points_entry.config(state=tk.NORMAL)
    paid_mode_points_entry.delete(0, tk.END)
    paid_mode_points_entry.insert(0, str(int(paid_mode_points)))
    paid_mode_points_entry.config(state='readonly')

# Функция для планирования следующего увеличения значения в платном режиме
def schedule_paid_mode_increment():
    global paid_mode_timer
    if paid_mode_var.get():
        paid_mode_timer = root.after(60000, schedule_paid_mode_increment)  # Вызываем каждую минуту
        increment_paid_mode()
    else:
        if paid_mode_timer is not None:
            root.after_cancel(paid_mode_timer)
            paid_mode_timer = None

# Функция для переключения платного режима
def toggle_paid_mode():
    global paid_mode_timer
    if paid_mode_var.get():
        if break_timer_running:
            messagebox.showwarning("Предупреждение", "Перерыв уже включен. Выключите перерыв, чтобы включить платный режим.")
            paid_mode_var.set(False)
            return
        # Деактивируем кнопки
        disable_buttons(True)
        # Начинаем увеличивать значение через 60 секунд
        root.after(60000, schedule_paid_mode_increment)  # Изменено на вызов через 60 секунд
        # Подсвечиваем чек-бокс желтым цветом
        paid_mode_button.config(bg="yellow")
    else:
        # Активируем кнопки
        disable_buttons(False)
        # Отменяем таймер
        if paid_mode_timer is not None:
            root.after_cancel(paid_mode_timer)
            paid_mode_timer = None
        # Возвращаем чек-боксу стандартный цвет
        paid_mode_button.config(bg="SystemButtonFace")

# Функция для деактивации/активации кнопок
def disable_buttons(disable):
    for btn in button_widgets.values():
        btn.config(state=tk.DISABLED if disable else tk.NORMAL)

# Функция для сброса значений полей
def reset_values():
    global processed_today_value, button_counters, paid_mode_points
    processed_today_value = 0
    paid_mode_points = 0
    update_processed_today_display()
    update_paid_mode_points_display()
    for value in button_counters.keys():
        button_counters[value] = 0
        button_counter_entries[value].config(state=tk.NORMAL)
        button_counter_entries[value].delete(0, tk.END)
        button_counter_entries[value].insert(0, "0")
        button_counter_entries[value].config(state='readonly')
        total_additions_entries[value].config(state=tk.NORMAL)
        total_additions_entries[value].delete(0, tk.END)
        total_additions_entries[value].insert(0, "0")
        total_additions_entries[value].config(state='readonly')

# Функция для переключения таймера перерыва
def toggle_break_timer():
    global break_timer_running, break_timer_label, break_timer_seconds
    if break_timer_running:
        if paid_mode_var.get():
            messagebox.showwarning("Предупреждение", "Платный режим уже включен. Выключите платный режим, чтобы включить перерыв.")
            return
        root.after_cancel(break_timer_id)
        break_timer_running = False
        change_button_color("break_button", "SystemButtonFace")
    else:
        if paid_mode_var.get():
            messagebox.showwarning("Предупреждение", "Платный режим уже включен. Выключите платный режим, чтобы включить перерыв.")
            return
        break_timer_running = True
        update_break_timer()
        change_button_color("break_button", "lightblue")

# Функция для обновления таймера перерыва
def update_break_timer():
    global break_timer_label, break_timer_seconds, break_timer_id
    if break_timer_running:
        hours, remainder = divmod(break_timer_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        time_format = f"{hours:02}:{minutes:02}:{seconds:02}"
        break_timer_label.config(text=time_format)
        break_timer_seconds += 1
        break_timer_id = root.after(1000, update_break_timer)

# Функция для изменения цвета кнопок
def change_button_color(button_name, color):
    if button_name == "break_button":
        break_button.config(bg=color)
    elif button_name == "paid_mode_button":
        paid_mode_button.config(bg=color)

# Запускаем графический интерфейс в отдельном потоке
threading.Thread(target=run_app).start()