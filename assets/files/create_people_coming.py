# pip install pandas

import pandas as pd
from pathlib import Path

FILE_PATH = "~/Downloads/Meeting Registration form (Responses) - Form Responses 1.csv"
df = pd.read_csv(Path(FILE_PATH).expanduser().resolve())
df = df[["Last Name", "First Name", "Affiliation/Institution"]]
df = df.sort_values("Last Name")
df = df.reset_index(drop=True)
df.index += 1
df.to_csv(Path(__file__).parent / "people_coming.csv", index=True, index_label="#")
