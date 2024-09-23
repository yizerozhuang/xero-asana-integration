import json
import asana
class User():
    def __init__(self, name, password, admin, asana_email):
        self.name = name
        self.password = password
        self.admin = admin
        self.asana_email = asana_email

        self.asana_gid = self.get_asana_gid()

    # def get_asana_gid(self):
    #
    #

    @staticmethod
    def get_user_from_json(json_dir):
        user_list = json.load(json_dir)
        for user in user_list:
            return User(user['name'], user['password'], user[''])
