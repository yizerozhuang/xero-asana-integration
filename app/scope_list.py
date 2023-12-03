import tkinter as tk

import os
import json

class ScopeList(list):
    def __init__(self, app, service, extra):
        super().__init__()
        self.app = app
        self.service = service
        self.extra = extra
        self.page = self.app.fee_proposal_page
    def __setitem__(self, key, value):
        super(ScopeList, self).__setitem__(key, value)

    def __delitem__(self, value):
        super(ScopeList, self).__delitem__(value)

    def __add__(self, value):
        super(ScopeList, self).__add__(value)

    def __iadd__(self, value):
        super(ScopeList, self).__iadd__(value)

    def append(self, value):
        super(ScopeList, self).append(value)
        self._append_value()

    def remove(self, value):
        super(ScopeList, self).remove(value)

    def _append_value(self):
        scope_dir = os.path.join(self.app.conf["database_dir"], "scope_of_work.json")
        scope = self.app.data["Fee Proposal"]["Scope"]
        item = self.page.append_context[self.service][self.extra]["Item"].get()
        scope[self.service][self.extra].append(
            {
                "Include":tk.BooleanVar(value=True),
                "Item":tk.StringVar(value=item)
            }
        )
        tk.Entry(self.page.scope_frames[self.service][self.extra], width=100,
                 textvariable=scope[self.service][self.extra][-1]["Item"],
                 font=self.app.conf["font"]).grid(row=len(scope[self.service][self.extra]), column=1)

        tk.Checkbutton(self.page.scope_frames[self.service][self.extra],
                       variable=scope[self.service][self.extra][-1]["Include"]).grid(row=len(scope[self.service][self.extra]), column=0)

        if self.page.append_context[self.service][self.extra]["Add"].get():
            scope_data = json.load(open(scope_dir))
            scope_data[self.service][self.extra].append(item)
            with open(scope_dir, "w") as f:
                json_object = json.dumps(scope_data, indent=4)
                f.write(json_object)
        self.page.append_context[self.service][self.extra]["Item"].set("")
        self.page.append_context[self.service][self.extra]["Add"].set(False)
