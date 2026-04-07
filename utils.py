import os
import sys

def resource_path(relative_path):
    """Получить абсолютный путь к ресурсу, работает для разработки и для PyInstaller"""
    try:

        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def rebuild(a, b):
    """Заменяет пустые значения на пробел"""
    if (str(a) == "" or len(str(a)) == str(a).count(' ')):
        return ' '
    else:
        return b

