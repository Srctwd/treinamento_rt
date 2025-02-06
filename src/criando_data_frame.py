import pandas as pd

df = pd.DataFrame(columns=['PDV','BANDEIRA'])

df = df.rename({"PDV":'PONTO DE VENDA'},axis=1)