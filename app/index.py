#carregamento e importações
from sqlalchemy import create_engine
from dotenv import load_dotenv
import os
import logging

load_dotenv()

def log_produtos():
    for arquivo in os.listdir("C:/Users/Windows/Documents/Python/mccain-automation/app/app/log"):
        if arquivo.endswith('.log'):
            logging.info('Arquivo iniciado')
    logging.basicConfig(
        filename='log/data.log',
        level=logging.INFO,
        format='%(asctime)s %(message)s',
        datefmt='%d/%m/%Y %I:%M:%S %p -',
)

if __name__ == '__main__':
    log_produtos()
