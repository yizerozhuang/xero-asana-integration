# import tkinter as tk
# from tkinter import messagebox
#
# import os
# import json
#
# class ScopeList(list):
#     def __init__(self, app, service, extra):
#         super().__init__()
#         self.app = app
#         self.service = service
#         self.extra = extra
#     def __setitem__(self, key, value):
#         super(ScopeList, self).__setitem__(key, value)
#
#     def __delitem__(self, value):
#         super(ScopeList, self).__delitem__(value)
#
#     def __add__(self, value):
#         super(ScopeList, self).__add__(value)
#
#     def __iadd__(self, value):
#         super(ScopeList, self).__iadd__(value)
#
#     def append(self, value):
#         super(ScopeList, self).append(value)
#         self._append_value()
#
#     def remove(self, value):
#         super(ScopeList, self).remove(value)
#
#     def _append_value(self):
#         scope_dir = os.path.join(self.app.conf["database_dir"], "scope_of_work.json")
#         scope = self.app.data["Fee Proposal"]["Scope"]
#         item = self.app.fee_proposal_page.append_context[self.service][self.extra]["Item"].get()
#         scope[service][extra].append(
#             {
#                 "Include":tk.BooleanVar(value=True),
#                 "Item":tk.StringVar(value=item)
#             }
#         )
#         tk.Entry(self.scope_frames[service].winfo_children()[index], width=100,
#                  textvariable=scope[service][extra][-1]["Item"],
#                  font=self.conf["font"]).grid(row=len(scope[service][extra]), column=1)
#
#         tk.Checkbutton(self.scope_frames[service].winfo_children()[index],
#                        variable=scope[service][extra][-1]["Include"]).grid(row=len(scope[service][extra]), column=0)
#
#         if self.append_context[service][extra]["Add"].get():
#             scope_data = json.load(open(scope_dir))
#             scope_data[service][extra].append(item)
#             with open(scope_dir, "w") as f:
#                 json_object = json.dumps(scope_data, indent=4)
#                 f.write(json_object)
#             # self.app.cur.execute(f"""
#             #     INSERT INTO service_scope(service_type, extra, context)
#             #     VALUES
#             #     ('{service}','{extra}','{item}')
#             #     """)
#             # self.app.conn.commit()
#         self.append_context[service][extra]["Item"].set("")
#         self.append_context[service][extra]["Add"].set(False)
