import pandas as pd

df = pd.DataFrame(columns=['PDV','BANDEIRA'])

df = df.rename({'BANDEIRA':'REDE'},axis=1)