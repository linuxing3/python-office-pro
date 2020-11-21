import pandas as pd

titanic = pd.read_csv("data/titanic.csv")
titanic.head(8)
titanic.to_excel('titanic.xlsx', sheet_name='passengers', index=False)

titanic = pd.read_excel('titanic.xlsx', sheet_name='passengers')
titanic.info()
