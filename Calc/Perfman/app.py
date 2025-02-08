import tkinter as tk
from .app_manager import App  # Updated to relative import

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
