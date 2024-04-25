from pathlib import Path
import pandas as pd

from loguru import logger
xlsx_path = Path(r'E:\FBXYFXYJ\gp23nscrfs-python\dataCache\20240331-新sprd.xlsx')



logger.info('read large excel [python_calamine]..')
calamine_df = pd.read_excel(xlsx_path, engine="calamine")
print(calamine_df.info())
logger.info('finished large excel [python_calamine], to pandas')


logger.info('read large excel [python_calamine]..')
from pandas import read_excel
from python_calamine.pandas import pandas_monkeypatch
pandas_monkeypatch()
calamine_df = read_excel(xlsx_path, engine="calamine")
print(calamine_df.info())
logger.info('finished large excel [python_calamine], to pandas')


import modin.config as modin_cfg
modin_cfg.Engine.put("dask")  # Modin will use Dask

import datatable
import pandas as pd

from bibibigtable.bigtable import read_large_excel

def read_excel(inputs, **kwargs):
    import pandas as pd

    return from_map(pd.read_excel, inputs, **kwargs)
#
# logger.info('read large excel [modin]..')
# import modin.pandas as mpd
#
# modin_df = mpd.read_excel(xlsx_path)
# # modin_df = mpd.read_excel(xlsx_path, engine='openpyxl')
# logger.info('finished large excel [modin], to pandas')

# logger.info('read large excel [python_calamine]..')
# from pandas import read_excel
# from python_calamine.pandas import pandas_monkeypatch
# pandas_monkeypatch()
# calamine_df = read_excel(xlsx_path, engine="calamine")
# print(calamine_df.info())
# logger.info('finished large excel [python_calamine], to pandas')
#
#
# logger.info('read large excel [read_large_excel]..')
# xl_df = read_large_excel(xlsx_path) # 6 min
# logger.info('finished large excel [read_large_excel], to pandas')

# logger.info('read large excel [dask]..')
# from dask.dataframe.io.io import from_map
# dask_df = read_excel([xlsx_path], engine='openpyxl')
# logger.info('finished large excel [dask], to pandas')



logger.info('read large excel [datatable] ..')
dt = datatable.fread(xlsx_path) # 6 min

logger.info('readed large excel, to pandas')
df = dt.to_pandas()

logger.info('finished large excel [datatable], to pandas')


logger.info('read large excel [pandas]..')
pd_df = pd.read_excel(xlsx_path)
logger.info('finished large excel [pandas], to pandas')

logger.info('read large excel [pandas - openpyxl]..')
pd_openpyxl_df = pd.read_excel(xlsx_path, engine='openpyxl')
logger.info('finished large excel [pandas - openpyxl], to pandas')


logger.info('read large excel [pandas - xlrd]..')
pd_xlrd_df = pd.read_excel(xlsx_path, engine='xlrd')
logger.info('finished large excel [pandas - xlrd], to pandas')


