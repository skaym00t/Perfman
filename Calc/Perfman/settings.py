import json
import os

class Settings:
    SETTINGS_FILE = ".settings.json"

    def __init__(self):
        self.buttons = {}
        self.modes = {}

    def load_settings(self):
        if os.path.exists(self.SETTINGS_FILE):
            with open(self.SETTINGS_FILE, 'r') as file:
                data = json.load(file)
                self.buttons = data.get('buttons', {})
                self.modes = data.get('modes', {})

    def save_settings(self):
        with open(self.SETTINGS_FILE, 'w') as file:
            json.dump({'buttons': self.buttons, 'modes': self.modes}, file)

    def reset_settings(self):
        self.buttons = {}
        self.modes = {}
        self.save_settings()

    def add_button(self, name, value):
        self.buttons[name] = value
        self.save_settings()

    def remove_button(self, name):
        if name in self.buttons:
            del self.buttons[name]
            self.save_settings()

    def remove_mode(self, name):
        if name in self.modes:
            del self.modes[name]
            self.save_settings()

    def add_mode(self, name, value):
        self.modes[name] = value
        self.save_settings()
