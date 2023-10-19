from function.app import App

CONFIGURATION = {
    "font":("Calibri", 11),
    "tax rates": 1.1,
    "n_building": 5,
    "n_drawing": 5,
    "n_items": 3
}

if __name__ == '__main__':
    app = App(CONFIGURATION)
    app.mainloop()