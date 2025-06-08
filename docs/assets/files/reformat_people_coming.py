# pip install numbers-parser
# pip install pandas

import pandas as pd
from numbers_parser import Document
from pathlib import Path

FILE_DIR = Path(__file__).parent

doc = Document(Path("/Users/nabil/Downloads/people_coming.numbers"))
sheets = doc.sheets
tables = sheets[0].tables
data = tables[0].rows(values_only=True)
df = pd.DataFrame(data[1:], columns=data[0])
df.insert(0, "#", range(1, len(df) + 1))
cols = df.columns.tolist()
first_idx = cols.index("First Name")
last_idx = cols.index("Last Name")
cols[first_idx], cols[last_idx] = cols[last_idx], cols[first_idx]
df = df[cols]
df.to_csv(FILE_DIR / "people_coming.csv", index=False)
