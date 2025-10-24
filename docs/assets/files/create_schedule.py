# pip install pandas openpyxl inflect
import pandas as pd
from pathlib import Path
import inflect
from urllib.parse import quote

p = inflect.engine()

# WARNING is all hardcoded to the sheet given to me.
# If that layout changes, this code will break
FILE_PATH = "~/Dropbox/draft_MUSE_IRIS_meeting_Oct2025.xlsx"
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
    "Chair",
    "End",
}
NAME_REPLACEMENTS = {
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
    .tg {font-family:Tahoma, sans-serif;border-collapse:collapse;border-spacing:0;}
    .tg td{border-style:none;padding:3px 3px;text-align:center;vertical-align:middle}
    .tg .tg-one{background-color:#332288;text-align:center;vertical-align:middle}
    .tg .tg-two{background-color:#0077BB;text-align:center;vertical-align:middle}
    .tg .tg-three{background-color:#88CCEE;text-align:center;vertical-align:middle}
    .tg .tg-four{background-color:#44AA99;text-align:center;vertical-align:middle}
    .tg .tg-five{background-color:#117733;text-align:center;vertical-align:middle}
    .tg .tg-six{background-color:#999933;text-align:center;vertical-align:middle}
    .tg .tg-seven{background-color:#DDCC77;text-align:center;vertical-align:middle}
    .tg .tg-eight{background-color:#EE7733;text-align:center;vertical-align:middle}
    .tg .tg-nine{border-style: double; border-width: thick;border-color: dimgrey;}
    .tg .tg-extra{border-color:#000000;font-weight:bold;text-align:center;vertical-align:middle}
    .md-typeset a{color: white}
    .tg a {
    color: white;
    text-shadow:
        -1px -1px 0 black,
        1px -1px 0 black,
        -1px 1px 0 black,
        1px 1px 0 black;
    }
</style>
<table class="tg">
    <thead>
        <tr>
            <th class="tg-extra" colspan="2" style="color:white;background-color:#CC3311">Monday</th>
            <th class="tg-extra" colspan="2" style="color:white;background-color:#882255">Tuesday</th>
            <th class="tg-extra" colspan="2" style="color:white;background-color:#AA4499">Wednesday</th>
            <th class="tg-extra" colspan="2" style="color:white;background-color:#EE3377">Thursday</th>
            <th class="tg-extra">Session Type</th>
        </tr>
    </thead>
    <tbody>
        {BODY}
    </tbody>
</table>
"""
LEGEND_HTML = """
<table class="tg">
    <tbody>
        <tr><td class="tg-five" style="color:white;">MUSE</td></tr>
        <tr><td class="tg-five" style="color:white;">Monday 9:10</td></tr>
        <tr><td class="tg-three" style="color:white;">Flares &amp; Eruptions</td></tr>
        <tr><td class="tg-three" style="color:white;">Monday 11:55</td></tr>
        <tr><td class="tg-two" style="color:white;">Corona</td></tr>
        <tr><td class="tg-two" style="color:white;">Tuesday 13:10</td></tr>
        <tr><td class="tg-one" style="color:white;">Chromosphere</td></tr>
        <tr><td class="tg-one" style="color:white;">Wednesday 13:20</td></tr>
        <tr><td class="tg-four" style="color:white;">Global Connections</td></tr>
        <tr><td class="tg-four" style="color:white;">Thursday 11:40</td></tr>
        <tr><td class="tg-six" style="color:white;">Future Capabilities</td></tr>
        <tr><td class="tg-six" style="color:white;">Thursday 14:40</td></tr>
        <tr><td class="tg-nine" style="color:Black;">Scene Setting Talks</td></tr>
    </tbody>
</table>
"""

HTML_TABLE_ROW_TEMPLATE = """
<tr>
    <td class="tg">{MONDAY_1}</td>
    <td class="tg-{SESSION_MONDAY}">{MONDAY_URL}</td>
    <td class="tg">{TUESDAY_1}</td>
    <td class="tg-{SESSION_TUESDAY}">{TUESDAY_URL}</td>
    <td class="tg">{WEDNESDAY_1}</td>
    <td class="tg-{SESSION_WEDNESDAY}">{WEDNESDAY_URL}</td>
    <td class="tg">{THURSDAY_1}</td>
    <td class="tg-{SESSION_THURSDAY}">{THURSDAY_URL}</td>
    {LEGEND_CELL}
</tr>
"""

SCENE_SETTING_NAMES = [
    "Reeves",
    "Longcope",
    "Klimchuk",
    "Rempel",
    "de la Cruz Rodriguez",
    "Carlsson",
    "Downs",
]


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
        print(f"Warning: Surname '{surname}' not found in schedule.")
        return ""
    # Find which column matched
    col = next(c for c in COL_OFFSETS if match.iloc[0, c] == surname)
    day = DAYS[col // 3]
    time = match.iloc[0, col - 1]
    time = time.strftime("%H:%M") if time != "" else "FILL IN"
    return f"{day} - {time}"


df = pd.read_excel(Path(FILE_PATH).expanduser().resolve(), sheet_name=2)
df = df.fillna("")
html_rows = []

rows = df.iloc[0:]
row_span = len(rows)
for i, (_, row) in enumerate(rows.iterrows()):
    cells = {}
    for day, base, css in DAY_CONFIG:
        cells[f"{day}_1"] = (
            row.iloc[base].strftime("%H:%M")
            if row.iloc[base] not in ("", "Chair")
            else row.iloc[base]
        )
        cells[f"{day}_2"] = row.iloc[base + 1]
        title = str(row.iloc[base + 1])
        session = parse_session(row, base + 2)
        session_mod = None
        # Scene setting special cases
        if any(name in title for name in SCENE_SETTING_NAMES):
            if title == "Rempel" and row.iloc[base].strftime("%H:%M") == "15:35":
                pass
            elif title == "Reeves" and row.iloc[base] == "":
                pass
            else:
                session_mod = "nine"
        if session_mod:
            session = session + f" tg-{session_mod}"
        cells[f"SESSION_{day}"] = session
        if title in BREAK_EVENTS or session == "zero":
            cells[f"{day}_URL"] = title
        else:
            # Apply any name fixes
            if title in NAME_REPLACEMENTS:
                title = NAME_REPLACEMENTS[title]
            anchor = quote(title)
            # Create the URL anchor
            cells[f"{day}_URL"] = (
                f'<a href="https://lm-sal.github.io/iris_muse_team_meeting/abstracts/#{anchor}">{title}</a>'
            )
    # Only add the legend cell to the first row; rowspan covers the rest
    if i == 0:
        cells["LEGEND_CELL"] = (
            f'<td class="tg" rowspan="{row_span}" style="vertical-align:top">{LEGEND_HTML}</td>'
        )
    else:
        cells["LEGEND_CELL"] = ""
    html_rows.append(HTML_TABLE_ROW_TEMPLATE.format(**cells))

schedule_preamble = f"Last updated on {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')} and subject to change.\n This is subject to revision if the US government shutdown ends before the meeting. \n\n"
html_table = HTML_TABLE_TEMPLATE.replace("{BODY}", "\n".join(html_rows))
html_table = schedule_preamble + html_table
output_path = Path(OUTPUT_FILE)
output_path.write_text(html_table, encoding="utf-8")

# Now to modify the abstract markdown file with the times
abstract_md = ABSTRACT_FILE.read_text(encoding="utf-8")
new_abstract_md = ""
# Need to drop the chair rows or they will match twice
df_no_chairs = df[~df.iloc[:, 0].str.contains("Chair", na=False)]
for entry in abstract_md.split("* "):
    if "**Author**: " in entry:
        name = entry.split("**Author**:")[1].split("<a")[0].strip()
        surname = " ".join(name.split(" ")[1:])
        when = find_when(df_no_chairs, surname) or "FILL IN"
        entry = entry.replace("FILL IN", f"{when}")
        new_abstract_md += "* " + entry
    else:
        new_abstract_md += entry
ABSTRACT_FILE.write_text(new_abstract_md, encoding="utf-8")
