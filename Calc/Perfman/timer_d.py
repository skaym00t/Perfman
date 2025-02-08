from datetime import datetime, timedelta
from tkinter import ttk

class Timer:
    def __init__(self):
        self.running = False
        self.start_time = None
        self.elapsed_time = timedelta(seconds=0)
        self.visible_timer_label = None

    def start_timer(self):
        if not self.running:
            self.start_time = datetime.now()
            self.running = True
            self.update_visible_timer()

    def stop_timer(self):
        if self.running:
            self.elapsed_time += datetime.now() - self.start_time
            self.running = False

    def toggle_timer(self):
        if self.running:
            self.stop_timer()
        else:
            self.start_timer()

    def reset_timer(self):
        self.start_time = datetime.now()
        self.elapsed_time = timedelta(seconds=0)
        if self.visible_timer_label:
            self.visible_timer_label.config(text="0:00:00")

    def update_visible_timer(self):
        if self.running:
            elapsed = self.elapsed_time + (datetime.now() - self.start_time)
            time_str = str(elapsed).split('.')[0]
            if self.visible_timer_label:
                self.visible_timer_label.config(text=time_str)
            self.visible_timer_label.after(1000, self.update_visible_timer)

    def create_timer(self, frame, row, col):
        self.visible_timer_label = ttk.Label(frame, text="0:00:00")
        self.visible_timer_label.grid(row=row, column=col, padx=5, pady=5)
        ttk.Button(frame, text="Сброс", command=self.reset_timer).grid(row=row+2, column=col, padx=5, pady=5)
