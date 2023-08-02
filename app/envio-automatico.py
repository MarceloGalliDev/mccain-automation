# flake8: noqa
from clientes import clientes
from estoques import estoques
from produtos import produtos
from vendas import vendas
from clientes_estado import clientes_estado 
from vendas_estado import vendas_estado
from estoques_estado import estoques_estado


import time

import ctypes
import threading


MB_OK = 0x0
TIMEOUT = 7000


def show_message_box():
    ctypes.windll.user32.MessageBoxW(0, "Envio de arquivos McCain realizado!", "Automação McCain", MB_OK)


def envio_automatico():
    clientes()
    estoques()
    produtos()
    vendas()
    clientes_estado()
    vendas_estado()
    estoques_estado()
   
    t = threading.Thread(target=show_message_box)
    t.start

    time.sleep(TIMEOUT / 1000.0)

    hwnd = ctypes.windll.user32.FindWindowW(None, "Automação McCain")
    if hwnd != 0:
        ctypes.windll.user32.SendMessageW(hwnd, 0x0010, 0, 0)
    
    
if __name__ == "__main__":
    envio_automatico()