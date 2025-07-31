import pandas as pd
from itertools import product
df=pd.read_csv(r"f5.csv")
df.head()
values = df['cell_id'].dropna().unique()
ordered_pairs = list(product(values, repeat=2))
pairs_df = pd.DataFrame(ordered_pairs, columns=['Source', 'Target'])
pairs_df.head()
pairs_df.to_csv(r"f5_pairs.csv",index_col=False)