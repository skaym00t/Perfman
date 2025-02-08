import tkinter as tk
from tkinter import ttk
from datetime import datetime
from .timer_d import Timer  # Updated to relative import

class ModeManager:
    def __init__(self, app):
        self.app = app
        self.modes = {}
        self.timers = {}
        self.active_mode = None

    def create_modes(self, frame):
        self.frame = frame
        self.reload_modes()

    def reload_modes(self):
        for widget in self.frame.winfo_children():
            if isinstance(widget, (ttk.Button, ttk.Label, ttk.Frame)):
                widget.destroy()
        row = 0
        col = 0
        for name, value in self.app.settings.modes.items():
            self.create_mode_row(row, col, name, value)
            col += 1
            if col >= 3:
                col = 0
                row += 1

    def create_mode_row(self, row, col, name, value):
        mode_frame = ttk.Frame(self.frame)
        mode_frame.grid(row=row, column=col, padx=5, pady=5)
        timer = Timer()
        self.timers[name] = timer
        button = ttk.Button(mode_frame, text=name, command=lambda n=name, v=value: self.on_mode_click(n, v, timer))
        button.grid(row=0, column=0, padx=5, pady=5)
        self.modes[name] = {
            'value': value,
            'button': button,
            'timer': timer
        }
        timer.create_timer(mode_frame, row=1, col=0)

    def on_mode_click(self, name, value, timer):
        if self.active_mode and self.active_mode == name:
            timer.toggle_timer()
            if not timer.running:
                self.active_mode = None
                self.enable_all_buttons()
        else:
            for mode_name, mode_data in self.modes.items():
                if mode_name != name:
                    mode_data['timer'].stop_timer()
                    mode_data['button'].config(state=tk.NORMAL)
            timer.toggle_timer()
            self.active_mode = name
            self.disable_other_modes(name)
            if value > 0 and timer.running:
                self.start_value_timer(name, value, timer)

    def is_any_mode_active(self):
        for mode_data in self.modes.values():
            if mode_data['timer'].running:
                return True
        return False

    def start_value_timer(self, name, value, timer):
        if timer.running:
            elapsed_seconds = int((timer.elapsed_time + (datetime.now() - timer.start_time)).total_seconds())
            if elapsed_seconds % 60 == 0 and elapsed_seconds > 0:
                self.app.done_today += value
                self.app.daily_goal -= value
                self.app.monthly_goal -= value
                self.app.sum_points += value
                self.app.index_value_month = self.app.sum_points / self.app.norm_value_month if self.app.norm_value_month > 0 else 0
                self.update_bonus_and_goals()
            self.app.root.after(1000, lambda: self.start_value_timer(name, value, timer))

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
        elif self.app.index_value_month > 1.41 and(self.app.index_value_month <= 1.51):
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

    def disable_other_modes(self, active_mode):
        for mode_name, mode_data in self.modes.items():
            if mode_name != active_mode:
                mode_data['button'].config(state=tk.DISABLED)
        active_mode_value = self.modes.get(active_mode, {}).get('value', 0)
        if active_mode_value > 0 or active_mode_value == None:
            self.disable_activity_buttons()
        else:
            self.enable_activity_buttons()

    def disable_activity_buttons(self):
        for widget in self.app.button_manager.frame.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.config(state=tk.DISABLED)

    def enable_activity_buttons(self):
        for widget in self.app.button_manager.frame.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.config(state=tk.NORMAL)

    def enable_all_buttons(self):
        for widget in self.app.button_manager.frame.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.config(state=tk.NORMAL)
        for widget in self.frame.winfo_children():
            if isinstance(widget, ttk.Frame):
                button = widget.winfo_children()[0]
                button.config(state=tk.NORMAL)

    def save_modes(self):
        self.app.settings.save_settings()
