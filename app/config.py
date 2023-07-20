#!/usr/bin/env python3
#conexão e configurações
from sqlalchemy import create_engine
from dotenv import load_dotenv
import os
import logging

load_dotenv()

class ConnectionData():
    DB_CONFIG = {
        'drivername': os.getenv('DRIVERNAME'),
        'host': os.getenv('HOST'),
        'port': os.getenv('PORT'),
        'database': os.getenv('DATABASE'),
        'username': os.getenv('USERNAME'),
        'password': os.getenv('PASSWORD'),
    }

    FTP_CONFIG = {
        'server-ftp': os.getenv('SERVER-FTP'),
        'user-ftp': os.getenv('USER-FTP'),
        'password-ftp': os.getenv('PASSWORD-FTP'),
        'path_clientes': os.getenv('PATH-CLIENTES'),
        'path_estoque': os.getenv('PATH-ESTOQUE'),
        'path_produto': os.getenv('PATH-PRODUTO'),
        'path_vendas': os.getenv('PATH-VENDAS'),
    }

def conn_engine(config):
    db_url = "{drivername}://{username}:{password}@{host}:{port}/{database}".format(**config)
    engine = create_engine(db_url)
    logging.info('Banco de dados conectado!')
    return engine

    
