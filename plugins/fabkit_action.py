import pcbnew
from PyQt5.QtWidgets import QApplication
from .gui import *

class FabKitAction(pcbnew.ActionPlugin):
    def defaults(self):
        self.name = "uBitFabKit"
        self.category = "Automation"
        self.description = "Plugin for creating BOM"
        self.show_toolbar_button = True # Optional, defaults to False
        self.icon_file_name = "../resources/icon.png" # Optional

    def Run(self):
        app = QApplication(['','--no-sandbox'])
        form = HomeWindow()
        form.show()
        app.exec_()
