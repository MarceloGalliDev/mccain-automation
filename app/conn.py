import os
import logging
from dotenv import load_dotenv
from sqlalchemy import create_engine, exc

load_dotenv()

def get_db_engine():
    try:
        db_url = os.getenv('URL')
        engine = create_engine(db_url)
        # Test the connection
        connection = engine.connect()
        connection.close()
        logging.info('Banco de dados conectado!')
        return engine
    except exc.SQLAlchemyError as e:
        logging.info(f"Error: {e}")
        return None

