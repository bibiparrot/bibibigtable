import pandas as pd
from loguru import logger


def excel_to_csv():
    excel_file = 'super_large_excel.xlsx'

    # Load the Excel file into a Pandas DataFrame
    # Use 'chunksize' parameter to process the file in smaller chunks if it's too large to fit into memory
    chunk_size = 100000  # Adjust the chunk size according to your system's memory capacity
    reader = pd.read_excel(excel_file, chunksize=chunk_size)

    # CSV output file path
    csv_output = 'output.csv'

    # Iterate through each chunk and append it to the CSV file
    for i, chunk in enumerate(reader):
        if i == 0:
            chunk.to_csv(csv_output, index=False, mode='w')  # Write header for the first chunk
        else:
            chunk.to_csv(csv_output, index=False, header=False, mode='a')  # Append chunks without header


def read_large_excel(io, sheet_name="Sheet1"):
    import openpyxl
    workbook = openpyxl.load_workbook(io, read_only=True)
    worksheet = workbook[sheet_name]
    rows = []
    for row in worksheet.iter_rows(values_only=True):
        rows.append(row)
    dataDF = pd.DataFrame(rows[1:], columns=rows[0])
    return dataDF


def read_sql_to_hdf(sql, db, hdf_file, hdf_key, chunk_size=None):
    if chunk_size is None:
        chunk_size = 1024 * 4
    iterator = pd.read_sql(sql, db, chunksize=chunk_size)

    store = pd.HDFStore(hdf_file)
    cols_to_index = None

    for i, chunk in enumerate(iterator):
        logger.info(f"Processing chunk {i}")
        store.append(hdf_key, chunk, data_columns=cols_to_index, index=False)

    store.create_table_index(hdf_key, columns=cols_to_index, optlevel=9, kind='full')
    store.close()


def read_sql_to_csv(sql, db, csv_file, chunk_size=None):
    if chunk_size is None:
        chunk_size = 1024 * 4
    iterator = pd.read_sql(sql, con=db, chunksize=chunk_size)

    for i, chunk in enumerate(iterator):
        logger.info(f"Processing chunk {i}")
        if i == 0:
            chunk.to_csv(csv_file, mode='w', encoding='utf8', header=True, index=False)
        else:
            chunk.to_csv(csv_file, mode='a', encoding='utf8', header=False, index=False)


def oracle_to_csv():
    import cx_Oracle
    # get a connection to the Oracle database
    dsn_tns = cx_Oracle.makedsn('ORACLE_HOSTNAME', 'PORT', service_name='SERVICE_NAME')
    db = cx_Oracle.connect(user='USERNAME', password='PASSWORD', dsn=dsn_tns)

    # specify the SQL query to execute
    sql = '''
    
    '''

    # use Pandas' chunksize parameter to read the data in chunks
    chunk_size = 5000
    iterator = pd.read_sql(sql, con=db, chunksize=chunk_size)

    # iterate over the chunks and perform your calculations or manipulations
    for i, chunk in enumerate(iterator):
        print(f"Processing chunk {i}")

        # write the processed data to a local CSV file
        if i == 0:
            chunk.to_csv('output.csv', mode='w', header=True, index=False)
        else:
            chunk.to_csv('output.csv', mode='a', header=False, index=False)

########################################################################################################################
