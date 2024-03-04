#PART OF CODE WAS COPIED FROM QGIS PLUGIN GEOSCAN FOREST (SOURCE - https://github.com/geoscan/geoscan_forest
def check_deps():
    from qgis.utils import iface
    import pathlib
    import os.path
    import subprocess

    from PyQt5.QtWidgets import QMessageBox
    plugin_dir = pathlib.Path(__file__).parent.parent

    requirements = ['pdfminer.six pdfminer', 'openpyxl openpyxl', 'pandas pandas']
    
    for dep in requirements:
                dep, import_tag = dep.strip().split()
                try:
                    mod = __import__(import_tag)
                    print("Модуль " + dep + " вже встановлено! ", mod)
                except ImportError as e:
                    print("Модуль {} не доступний, встановлюю...".format(dep))

                    reply = QMessageBox.question(iface.mainWindow(), 'Module install',
                                                 'Для роботи плагіна потрібно встановити модуль:' + dep + '. Продовжити?',
                                                 QMessageBox.Yes, QMessageBox.No)
                    if reply == QMessageBox.Yes:
                        result = subprocess.run(["python3", '-m', 'pip', 'install', dep])

                        if result.returncode == 0:
                            print("Модуль {} успішно встановлено!".format(dep))
                        else:
                            print("Модуль {} не вдалося встановити!".format(dep))
                    else:
                        return 1
    return 0
