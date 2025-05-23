import pandas as pd

csv_file = "people_coming.csv"
df = pd.read_csv(csv_file)
df.insert(0, "#", range(1, len(df) + 1))
cols = df.columns.tolist()
first_idx = cols.index("First Name")
last_idx = cols.index("Last Name")
cols[first_idx], cols[last_idx] = cols[last_idx], cols[first_idx]
df = df[cols]
df.to_csv("people_coming.csv", index=False)
