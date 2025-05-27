import tkinter as tk

class AlertWindow(tk.Frame):

    def __init__(self, parent, message, numbered_buttons=0):
        super().__init__(parent, relief="raised")
        self.choice = None
        self.numbered_buttons = numbered_buttons

        label = tk.Label(self, text=message, wraplength=1000, font=("Courier", 11), justify=tk.LEFT)
        label.pack(padx=20, pady=20)

        yes_button = tk.Button(self, text="Yes", command=self.on_yes_button_click, height=2, width=10, bg="#A2FF9F")
        yes_button.pack(side=tk.LEFT, expand=True, fill=tk.X)

        no_button = tk.Button(self, text="No", command=self.on_no_button_click, height=2, width=10, bg="#FFA2A2")
        no_button.pack(side=tk.LEFT, expand=True, fill=tk.X)

        for i in range(numbered_buttons): # Numbered buttons
            numbered_button = tk.Button(self, text=str(i+1), command=lambda i=i: self.on_numbered_button_click(i), height=2, width=10, bg="#A2A2FF")
            numbered_button.pack(side=tk.LEFT, expand=True, fill=tk.X)
            
    def on_numbered_button_click(self, i):
        self.choice = i
        self.destroy()

    def on_yes_button_click(self):
        self.choice = True
        self.destroy()

    def on_no_button_click(self):
        self.choice = False
        self.destroy()

    def get_answer(self):
        self.wait_window(self)
        return self.choice