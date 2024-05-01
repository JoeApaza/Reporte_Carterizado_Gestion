import logging
import requests
#import cx_Oracle
import oracledb
import polars as pl

#funcion para conectarme  base de datos
def get_connection(user_db,password_db,dsn_db):
    logging.info(f'Iniciando proceso de conexion a la base de datos {dsn_db}')
    try:
        conexion= oracledb.connect(
            user=user_db,
            password=password_db,
            dsn=dsn_db
        )
        logging.info(f'Conexion exitosa a la base de datos {dsn_db}')
        return conexion
    except Exception as ex:
        logging.error(ex)

#funcion para cerrar la conexion a base de datos
def close_connection_db(conexion):
    logging.info('Iniciando proceso para cerrar conexion a base de datos')
    try:
        cierre_conexion= conexion.close()
        logging.info('Se cerro conexion de manera exitosa')
        return cierre_conexion
    except Exception as ex:
        logging.error(ex)

dtypes = {
        "NRO_CUENTA":pl.Utf8,
        "GRUPO_FAC":pl.Utf8,
        "CICLO":pl.Utf8,
        "RAZON_SOCIAL":pl.Utf8,
        "RUC":pl.Utf8,
        "ORIGEN":pl.Utf8,
        "CUENTA_LARGA":pl.Utf8,
        "DOCUMENTO":pl.Utf8,
        "EMISION":pl.Datetime(time_unit='us', time_zone=None),
        "VENCIMIENTO":pl.Datetime(time_unit='us', time_zone=None),
        "MONEDA":pl.Utf8,
        "CAMBIO":pl.Float64,
        "MONTO_FAC":pl.Float64,
        "SALDO":pl.Float64,
        "SALDO_SOLES":pl.Float64,
        "DEUDA_SOLES":pl.Float64,
        "SALDO_A_FAVOR":pl.Float64,
        "EJECUTIVO":pl.Utf8,
        "SUPERVISOR":pl.Utf8,
        "SECTOR":pl.Utf8,
        "REGION_CAE":pl.Utf8,
        "GRUPO":pl.Utf8,
        "DIRECCION":pl.Utf8,
        "SEGMENTO":pl.Utf8,
        "WHITELIST":pl.Utf8,
        "GESTOR":pl.Utf8,
        "NOMBRE_CARTERA":pl.Utf8,
        "DIAS":pl.Utf8,
        "TRAMO_VENCIMIENTO":pl.Utf8,
        "TOP":pl.Utf8,
        "PLAZO":pl.Utf8,
        "PLAZO_ESPECIAL":pl.Utf8,
        "OPERADORES":pl.Utf8,
        "GESTOR_DETALLE":pl.Utf8,
        "GESTOR_COBRANZAS":pl.Utf8
}

def read_database_db(sql_query,source_connection,dtypes):
        logging.info('Iniciando proceso para guardar la informacion en un dataframe')
        df_polars = pl.read_database(sql_query, source_connection,batch_size=0,schema_overrides=dtypes)
        logging.info('Se guardo la informacion en un dataframe')
        return df_polars


def leer_sql(archivo_sql):
    logging.info('Se inicia la funcion de leer contenido de archivo sql')
    with open(archivo_sql, 'r',encoding='utf-8') as archivo:
        x=archivo.read()
        logging.info('Se ha leido todo el contenido del archivo sql')
        return x
    
