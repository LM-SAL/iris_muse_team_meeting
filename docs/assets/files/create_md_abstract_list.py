import pandas as pd
from pathlib import Path

ABSTRACT_FILE = Path(__file__).parent / "accepted_abstracts.csv"

abstracts = pd.read_csv(ABSTRACT_FILE)
abstracts = abstracts.rename(
    columns={
        "First Name": "First Name",
        "Last Name": "Last Name",
        "Abstract title": "Abstract title",
        "List of (co)authors if any:": "List of (co)authors",
    }
)
abstract_data = abstracts.fillna("None").to_dict(orient="records")
template = """
* **Author**: {First Name} {Last Name}

    **Coauthors**: {List of (co)authors}
    <details> <summary> <b>Title</b>: {Abstract title} </summary>
      <b>Abstract</b>: {Abstract (max 300 words)}
    </details>

"""
markdown_text = "\n".join([template.format(**entry) for entry in abstract_data])
output_file = Path(__file__).parent.parent.parent / "abstracts.md"
with open(output_file, "r") as f:
    existing_content = f.read()
split_point = "Accepted Abstracts."
if split_point in existing_content:
    header_part = existing_content.split(split_point)[0] + split_point + "\n\n"
else:
    header_part = "---\ntitle: Abstracts\n---\n\nAccepted Abstracts.\n\n"
with open(output_file, "w") as f:
    f.write(header_part + markdown_text)
