# flake8: noqa
from clientes import clientes
from estoques import estoques
from produtos import produtos
from vendas import vendas
import time

import ctypes
import threading


MB_OK = 0x0
TIMEOUT = 5000


def show_message_box():
    ctypes.windll.user32.MessageBoxW(0, "Envio de arquivos McCain realizado!", "Automação McCain", MB_OK)


def envio_automatico():
    clientes()
    time.sleep(5)
    estoques()
    time.sleep(5)
    produtos()
    time.sleep(5)
    vendas()
    time.sleep(5)
   
    t = threading.Thread(target=show_message_box)
    t.start

    time.sleep(TIMEOUT / 1000.0)

    hwnd = ctypes.windll.user32.FindWindowW(None, "Automação McCain")
    if hwnd != 0:
        ctypes.windll.user32.SendMessageW(hwnd, 0x0010, 0, 0)
    
    
if __name__ == "__main__":
    envio_automatico()