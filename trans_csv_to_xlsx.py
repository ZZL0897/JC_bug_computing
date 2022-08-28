import pandas as pd
import os
from pathlib import Path

path = r'C:\Users\65487\Desktop\股票代码\20\沪市主板'
file = os.listdir(path)

for f in file:
    f_path = Path(os.path.join(path, f))
    csv = pd.read_csv(f_path, encoding='utf-8')
    csv.to_excel(os.path.join(path, f_path.stem + '.xlsx'), sheet_name=f_path.stem)

new_file = os.listdir(path)

for f in new_file:
    f = Path(f)
    if f.suffix == '.csv':
        os.remove(os.path.join(path, f))

