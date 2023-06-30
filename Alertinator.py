import tkinter as tk

class AlertWindow(tk.Toplevel):

    def __init__(self, message):
        super().__init__()
        self.title("Alert")
        self.geometry("600x300")
        self.choice = None

        label = tk.Label(self, text=message, wraplength=1000, font=("Courier", 11), justify=tk.LEFT)
        label.pack(padx=20, pady=20)

        yes_button = tk.Button(self, text="Yes", command=self.on_yes_button_click, height=2, width=10, bg="#A2FF9F")
        yes_button.pack(side=tk.LEFT, expand=True, fill=tk.X)

        no_button = tk.Button(self, text="No", command=self.on_no_button_click, height=2, width=10, bg="#FFA2A2")
        no_button.pack(side=tk.LEFT, expand=True, fill=tk.X)

    def on_yes_button_click(self):
        self.choice = True
        self.destroy()

    def on_no_button_click(self):
        self.choice = False
        self.destroy()

    def get_answer(self):
        self.wait_window()
        return self.choice