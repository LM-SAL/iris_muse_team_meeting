import pandas as pd
from pathlib import Path
from urllib.parse import quote

FILE_PATH = "~/Downloads/Abstract submission form (Responses) - Form Responses 1.csv"
abstracts = pd.read_csv(Path(FILE_PATH).expanduser().resolve())
abstracts = abstracts.rename(
    columns={
        "First Name": "First Name",
        "Last Name": "Last Name",
        "Abstract title": "Abstract title",
        "List of (co)authors if any:": "List of (co)authors",
    }
)
abstracts = abstracts.drop(columns=["Timestamp", "Email Address"])


# Reformat the abstract text to avoid markdown issues
def custom_indent(text: str, prefix: str = "      ") -> str:
    new_text = []
    for i, line in enumerate(text.splitlines()):
        if i == 0:
            new_text.append(line)
            continue
        new_text.append(prefix + line)
    return "\n".join(new_text)


abstracts["Abstract (max 300 words)"] = abstracts["Abstract (max 300 words)"].apply(
    lambda x: custom_indent(x)
)
abstract_data = (
    abstracts.fillna("None").sort_values(by="Last Name").to_dict(orient="records")
)
abstract_data = [
    {**entry, "Last Name Quote": quote(entry["Last Name"])} for entry in abstract_data
]
# Strip any trailing whitespace from any field
abstract_data = [
    {k: v.strip() if isinstance(v, str) else v for k, v in entry.items()}
    for entry in abstract_data
]
# Below "When" is filled when the schedule is created
template = """
* <p id="{Last Name}">**Author**: {First Name} {Last Name} <a class="headerlink" href="#{Last Name Quote}" title="Permanent link">¶</a>

    **When**: FILL IN

    **Coauthors**: {List of (co)authors}
    <details> <summary> <b>Title</b>: {Abstract title} </summary>
      <b>Abstract</b>
      {Abstract (max 300 words)}
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
