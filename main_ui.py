import src.GUI as GUI

import sys
from PyQt5 import QtWidgets

import logging
import logging.config

# Uncomment to disable logging outside of this application.
logging.config.dictConfig({
    'version': 1,
    'disable_existing_loggers': True,
})

logging.basicConfig(filename="sparx_automation.log", filemode="w", format='%(asctime)s [%(name)s.%(funcName)s] - %(levelname)s - %(message)s',
                    level=logging.DEBUG)


def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)

app = QtWidgets.QApplication(sys.argv)

ui = GUI.GuiController()

ui.pre_load_forms()
sys.excepthook = except_hook
ui.show()
sys.exit(app.exec_())