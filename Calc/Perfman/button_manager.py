from tkinter import ttk

class ButtonManager:
    def __init__(self, app):
        self.app = app
        self.buttons = {}

    def create_buttons(self, frame):
        self.frame = frame
        self.reload_buttons()

    def reload_buttons(self):
        for widget in self.frame.winfo_children():
            if isinstance(widget, (ttk.Button, ttk.Label)):
                widget.destroy()
        row = 0
        col = 0
        for name, value in self.app.settings.buttons.items():
            self.create_button_row(row, col, name, value)
            col += 2
            if col >= 4:
                col = 0
                row += 1

    def create_button_row(self, row, col, name, value):
        button = ttk.Button(self.frame, text=name, command=lambda n=name, v=value: self.on_button_click(n, v))
        button.grid(row=row, column=col, padx=5, pady=5)
        counter_label = ttk.Label(self.frame, text="0")
        counter_label.grid(row=row, column=col+1, padx=5, pady=5)
        self.buttons[name] = {
            'value': value,
            'button': button,
            'counter_label': counter_label,
            'counter': 0,
            'cumulative': 0
        }

    def on_button_click(self, name, value):
        button_data = self.buttons[name]
        button_data['counter'] += 1
        button_data['cumulative'] += value
        button_data['counter_label'].config(text=str(button_data['counter']))
        self.app.done_today += value
        self.app.daily_goal -= value
        self.app.monthly_goal -= value
        self.app.sum_points += value
        self.app.index_value_month = self.app.sum_points / self.app.norm_value_month if self.app.norm_value_month > 0 else 0
        self.update_bonus_and_goals()
        self.app.button_history.append((name, value))

    def update_bonus_and_goals(self):
        if self.app.index_value_month >= 1.0 and self.app.index_value_month <= 1.11:
            self.app.bonus_goal = (self.app.norm_value_month * 1.11) - self.app.sum_points
            self.app.bonus = 0
            self.app.next_bonus = 1000
        elif self.app.index_value_month > 1.1 and self.app.index_value_month <= 1.21:
            self.app.bonus_goal = (self.app.norm_value_month * 1.21) - self.app.sum_points
            self.app.bonus = 1000
            self.app.next_bonus = 1500
        elif self.app.index_value_month > 1.21 and self.app.index_value_month <= 1.31:
            self.app.bonus_goal = (self.app.norm_value_month * 1.31) - self.app.sum_points
            self.app.bonus = 1500
            self.app.next_bonus = 2000
        elif self.app.index_value_month > 1.31 and self.app.index_value_month <= 1.41:
            self.app.bonus_goal = (self.app.norm_value_month * 1.41) - self.app.sum_points
            self.app.bonus = 2000
            self.app.next_bonus = 2500
        elif self.app.index_value_month > 1.41 and self.app.index_value_month <= 1.51:
            self.app.bonus_goal = (self.app.norm_value_month * 1.51) - self.app.sum_points
            self.app.bonus = 2500
            self.app.next_bonus = 3000
        elif self.app.index_value_month > 1.51 and self.app.index_value_month <= 1.630:
            self.app.bonus_goal = (self.app.norm_value_month * 1.630) - self.app.sum_points
            self.app.bonus = 3000
            self.app.next_bonus = 3750
        elif self.app.index_value_month > 1.625:
            self.app.bonus_goal = 0
            self.app.bonus = 3750
            self.app.next_bonus = 'Ты молодец!'
        else:
            self.app.bonus_goal = 0
            self.app.bonus = 0
            self.app.next_bonus = 0
        self.update_goals()

    def update_goals(self):
        self.app.daily_goal_label.config(text=f"{max(0, int(self.app.daily_goal))}")
        self.app.monthly_goal_label.config(text=f"{max(0, int(self.app.monthly_goal))}")
        self.app.done_today_label.config(text=f"{int(self.app.done_today)}")
        self.app.index_value_month_label.config(text=(f"{(self.app.index_value_month)}")[:5])
        self.app.bonus_label.config(text=f"{max(0, int(self.app.bonus))}")
        self.app.bonus_goal_label.config(text=f"{max(0, int(self.app.bonus_goal))}")
        self.app.next_bonus_label.config(text=f"{self.app.next_bonus}")

    def undo_last_action(self):
        if self.app.button_history:
            last_action = self.app.button_history.pop()
            name, value = last_action
            button_data = self.buttons[name]
            if button_data['counter'] > 0:
                button_data['counter'] -= 1
                button_data['cumulative'] -= value
                button_data['counter_label'].config(text=str(button_data['counter']))
            self.app.done_today -= value
            self.app.daily_goal += value
            self.app.monthly_goal += value
            self.app.sum_points -= value
            self.app.index_value_month = self.app.sum_points / self.app.norm_value_month if self.app.norm_value_month > 0 else 0
            self.update_bonus_and_goals()
