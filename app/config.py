#carregamento e importações
from sqlalchemy import create_engine, exc
from dotenv import load_dotenv
import os
import logging

load_dotenv()

def log_data():        
    for arquivo in os.listdir("C:/Users/Windows/Documents/Python/mccain-automation/app/log"):
        if arquivo.endswith('.log'):
            logging.info('Arquivo iniciado')
    logging.basicConfig(
        filename='log/data.log',
        level=logging.INFO,
        format='%(asctime)s %(message)s',
        datefmt='%d/%m/%Y %I:%M:%S %p -',
)
    
def get_db_engine():
    log_data()
    try:
        db_url = os.getenv('URL')
        engine = create_engine(db_url)
        # Test connection
        with engine.connect() as connection:
            logging.info('Conexão estabelecida!')
            pass
        logging.info('Banco de dados conectado!')
        return engine
    except exc.SQLAlchemyError as e:
        logging.info(f"Error: {e}")
        return None


