import pandas as pd
import numpy as np
from constants import *
from openpyxl import load_workbook

def pivot_table():
  df = pd.read_excel(filename, header=1)
  # print(df)
  pivot = pd.pivot_table(df,
    values=['Pesimista Proy.', 'Optimista Proy.'], 
    index=['Oficina', 'Sector', 'Material'], 
    aggfunc={ 'Pesimista Proy.': np.sum, 'Optimista Proy.': np.sum }
  )
  # print(pivot)

  df2 = pd.DataFrame({
    "A": ["foo", "foo", "foo", "foo", "foo","bar", "bar", "bar", "bar"],
    "B": ["one", "one", "one", "two", "two","one", "one", "two", "two"],
    "C": ["small", "large", "large", "small","small", "large", "small", "small","large"],
    "D": [1, 2, 2, 3, 3, 4, 5, 6, 7],
    "E": [2, 4, 5, 5, 6, 6, 8, 9, 9]}
  )

  df1 = pd.DataFrame([['a', 'b'], ['c', 'd']],
    index=['row 1', 'row 2'],
    columns=['col 1', 'col 2']
  )

  with pd.ExcelWriter(filename, mode="a", engine="openpyxl") as writer:
    pivot.to_excel(writer, sheet_name="TD")
