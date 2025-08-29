# pip install pandas openpyxl inflect
import pandas as pd
from pathlib import Path
import inflect
from urllib.parse import quote

p = inflect.engine()

# WARNING is all hardcoded to the sheet given to me.
# If that layout changes, this code will break
FILE_PATH = "~/Downloads/draft_MUSE_IRIS_meeting_Oct2025.xlsx"
OUTPUT_FILE = Path(__file__).parent.parent.parent / "schedule.md"
ABSTRACT_FILE = Path(__file__).parent.parent.parent / "abstracts.md"
COL_OFFSETS = [1, 4, 7, 10]
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday"]
BREAK_EVENTS = {
    "Lunch break",
    "Coffee break",
    "Discussion",
    "Social Dinner",
    "Reception",
    "Close out",
}
NAME_REPLACEMENTS = {
    "Jose Diaz Baso": "Diaz Baso",
    "Kumar Srivastava": "Srivastava",
    "Franco Rappazzo": "Rappazzo",
}
DAY_CONFIG = [
    ("MONDAY", 0, "one"),
    ("TUESDAY", 3, "two"),
    ("WEDNESDAY", 6, "three"),
    ("THURSDAY", 9, "four"),
]
HTML_TABLE_TEMPLATE = """
<style type="text/css">
    .tg  {border-collapse:collapse;border-spacing:0;}
    .tg td{border-color:black;border-style:none;border-width:1px;font-family:Arial, sans-serif;font-size:14px;
    overflow:hidden;padding:10px 5px;word-break:normal;text-align:center;}
    .tg th{border-color:black;border-style:none;border-width:1px;font-family:Arial, sans-serif;font-size:14px;font-weight:normal;overflow:hidden;padding:10px 5px;word-break:normal;}
    .tg .tg-one{background-color:#e49edd;border-color:inherit;text-align:center;vertical-align:top}
    .tg .tg-two{background-color:#60cbf3;border-color:inherit;text-align:center;vertical-align:top}
    .tg .tg-three{background-color:#ffc000;border-color:inherit;text-align:center;vertical-align:top}
    .tg .tg-four{background-color:#fefe00;border-color:inherit;text-align:center;vertical-align:top}
    .tg .tg-five{background-color:#caedfb;border-color:inherit;text-align:center;vertical-align:top}
    .tg .tg-six{background-color:#7ffe00;border-color:inherit;text-align:center;vertical-align:top}
    .tg .tg-seven{background-color:#fae2d5;border-color:inherit;text-align:center;vertical-align:top}
    .tg .tg-eight{background-color:#9ee2a3;border-color:inherit;text-align:center;vertical-align:top}
    .tg .tg-mcqj{border-color:#000000;font-weight:bold;text-align:center;vertical-align:top}
    .tg .tg-73oq{border-color:#000000;text-align:center;vertical-align:top}
    .tg .tg-0lax{text-align:center;vertical-align:top}
    .tg .tg-0pky{border-color:inherit;text-align:center;vertical-align:top}
</style>
<table class="tg">
    <thead>
        <tr>
            <th class="tg-mcqj" colspan="2" style="background-color:red">Monday</th>
            <th class="tg-mcqj" colspan="2" style="background-color:orange">Tuesday</th>
            <th class="tg-mcqj" colspan="2" style="background-color:yellow">Wednesday</th>
            <th class="tg-mcqj" colspan="2" style="background-color:green">Thursday</th>
        </tr>
    </thead>
    <tbody>
        {BODY}
    </tbody>
</table>
<table class="tg">
    <thead>
        <tr>
            <th class="tg-mcqj">Session Type</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td class="tg-one">Chromosphere</td>
        </tr>
        <tr>
            <td class="tg-two">Corona</td>
        </tr>
        <tr>
            <td class="tg-three">Flares &amp; Eruptions</td>
        </tr>
        <tr>
            <td class="tg-four">Global Connections</td>
        </tr>
        <tr>
            <td class="tg-five">MUSE</td>
        </tr>
        <tr>
            <td class="tg-six">Future Capabilities</td>
        </tr>
        <tr>
            <td class="tg-seven">Scene Setting</td>
        </tr>
    </tbody>
</table>
"""
HTML_TABLE_ROW_TEMPLATE = """
<tr>
    <td class="tg-0pky">{MONDAY_1}</td>
    <td class="tg-{SESSION_MONDAY}"><a href="{MONDAY_URL}">{MONDAY_2}</a></td>
    <td class="tg-0pky">{TUESDAY_1}</td>
    <td class="tg-{SESSION_TUESDAY}"><a href="{TUESDAY_URL}">{TUESDAY_2}</a></td>
    <td class="tg-0pky">{WEDNESDAY_1}</td>
    <td class="tg-{SESSION_WEDNESDAY}"><a href="{WEDNESDAY_URL}">{WEDNESDAY_2}</a></td>
    <td class="tg-0pky">{THURSDAY_1}</td>
    <td class="tg-{SESSION_THURSDAY}"><a href="{THURSDAY_URL}">{THURSDAY_2}</a></td>
</tr>
"""


def parse_session(row, idx):
    """
    Turn the numeric session code into a word, or 'zero' if blank/invalid.
    """
    try:
        return p.number_to_words(int(row.iloc[idx]))
    except (ValueError, TypeError):
        return "zero"


def find_when(df, surname):
    # Build one boolean series that matches any of the 4 surname‐columns
    mask = False
    for col in COL_OFFSETS:
        mask |= df.iloc[:, col].str.contains(surname, case=True, na=False)
    match = df[mask]
    if match.empty:
        raise ValueError(f"Surname '{surname}' not found in schedule.")
    # Find which column matched
    col = next(c for c in COL_OFFSETS if match.iloc[0, c] == surname)
    day = DAYS[col // 3]
    time = match.iloc[0, col - 1]
    return f"{day} - {time}"


df = pd.read_excel(Path(FILE_PATH).expanduser().resolve())
df = df.fillna("")
html_rows = []
for _, row in df.iloc[1:].iterrows():
    cells = {}
    for day, base, css in DAY_CONFIG:
        cells[f"{day}_1"] = row.iloc[base]
        cells[f"{day}_2"] = row.iloc[base + 1]
        cells[f"SESSION_{day}"] = parse_session(row, base + 2)
        title = str(row.iloc[base + 1])
        if title in BREAK_EVENTS:
            cells[f"{day}_URL"] = ""
        else:
            # apply any name fixes
            for old, new in NAME_REPLACEMENTS.items():
                title = title.replace(old, new)

            anchor = quote(title)
            cells[f"{day}_URL"] = (
                f"https://lm-sal.github.io/iris_muse_team_meeting/abstracts/#{anchor}"
            )
    html_rows.append(HTML_TABLE_ROW_TEMPLATE.format(**cells))

html_table = HTML_TABLE_TEMPLATE.replace("{BODY}", "\n".join(html_rows))
output_path = Path(OUTPUT_FILE)
output_path.write_text(html_table, encoding="utf-8")

# Now to modify the abstract markdown file with the times
abstract_md = ABSTRACT_FILE.read_text(encoding="utf-8")
new_abstract_md = ""
for entry in abstract_md.split("* "):
    if "**Author**: " in entry:
        name = entry.split("**Author**:")[1].split("<a")[0].strip()
        surname = " ".join(name.split(" ")[1:])
        when = find_when(df, surname) or "FILL IN"
        entry = entry.replace("FILL IN", f"{when}")
        new_abstract_md += "* " + entry
    else:
        new_abstract_md += entry
ABSTRACT_FILE.write_text(new_abstract_md, encoding="utf-8")
