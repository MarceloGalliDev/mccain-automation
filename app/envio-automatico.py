from clientes import clientes
from estoques import estoques
from produtos import produtos
from vendas import vendas
import time
import ctypes

def envio_automatico():
    clientes()
    time.sleep(5)
    estoques()
    time.sleep(5)
    produtos()
    time.sleep(5)
    vendas()
    time.sleep(5)

    ctypes.windll.user32.MessageBoxW(0, "Envio de arquivos McCain realizado!", "Automação McCain", 1)

if __name__ == "__main__":
    envio_automatico()