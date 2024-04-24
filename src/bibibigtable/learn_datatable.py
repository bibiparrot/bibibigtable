from pathlib import Path

import datatable
import pandas as pd
from loguru import logger

from bibibigtable.bigtable import read_large_excel

xlsx_path = Path(r'E:\FBXYFXYJ\gp23nscrfs-python\dataCache\20240331-新sprd.xlsx')
logger.info('read large excel [datatable] ..')
dt = datatable.fread(xlsx_path)

logger.info('readed large excel, to pandas')
df = dt.to_pandas()

logger.info('finished large excel [datatable], to pandas')


logger.info('read large excel [read_large_excel]..')
df = read_large_excel(xlsx_path)
logger.info('finished large excel [read_large_excel], to pandas')


logger.info('read large excel [pandas]..')
df = pd.read_excel(xlsx_path)
logger.info('finished large excel [pandas], to pandas')