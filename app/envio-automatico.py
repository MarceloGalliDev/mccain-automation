from clientes import clientes
from estoques import estoques
from produtos import produtos
from vendas import vendas
import time
import ctypes

def envio_automatico():
    clientes()
    time.sleep(3)
    estoques()
    time.sleep(3)
    produtos()
    time.sleep(3)
    vendas()
    time.sleep(3)
    
    ctypes.windll.user32.MessageBoxW(0, "Envio de arquivos McCain realizado!", "Automação McCain", 1)

if __name__ == "__main__":
    envio_automatico()