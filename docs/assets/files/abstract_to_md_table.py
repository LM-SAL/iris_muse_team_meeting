import pandas as pd

csv_file = "./accepted_abstracts.csv"
df = pd.read_csv(csv_file)
markdown_table = df.to_markdown(index=False)
with open("abstracts_table.md", "w") as f:
    f.write(markdown_table)
