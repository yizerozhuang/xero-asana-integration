from app import App
from login import Login

from config import CONFIGURATION


if __name__ == '__main__':
    user = [""]
    login = [False]
    login_app = Login(CONFIGURATION, user, login)
    login_app.mainloop()
    if login[0]:
        app = App(CONFIGURATION, user[0])
        app.mainloop()
