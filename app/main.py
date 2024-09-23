from app import App
from login import Login

from config import CONFIGURATION


if __name__ == '__main__':
    user = [""]
    login = [False]
    user_email = [""]
    admin = [False]
    login_app = Login(CONFIGURATION, user, user_email, login, admin)
    login_app.mainloop()
    if login[0]:
        app = App(CONFIGURATION, user[0], user_email[0], admin[0])
        app.mainloop()
