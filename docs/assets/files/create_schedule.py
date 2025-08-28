# pip install pandas openpyxl inflect
import pandas as pd
from pathlib import Path
import inflect

# WARNING is all hardcoded to the sheet given to me.
# If that layout changes, this code will break
FILE_PATH = "~/Downloads/draft_MUSE_IRIS_meeting_Oct2025.xlsx"
OUTPUT_FILE = Path(__file__).parent.parent.parent / "schedule.md"
p = inflect.engine()
df = pd.read_excel(Path(FILE_PATH).expanduser().resolve())
df = df.fillna("")
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
<table class="tg"><thead>
  <tr>
    <th class="tg-mcqj" colspan="2" style="background-color:red">Monday</th>
    <th class="tg-mcqj" colspan="2" style="background-color:orange">Tuesday</th>
    <th class="tg-mcqj" colspan="2" style="background-color:yellow">Wednesday</th>
    <th class="tg-mcqj" colspan="2" style="background-color:green">Thursday</th>
  </tr></thead>
<tbody>
{BODY}
</tbody>
</table>
<br>
<table class="tg"><thead>
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
    <td class="tg-{SESSION_MONDAY}">{MONDAY_2}</td>
    <td class="tg-0pky">{TUESDAY_1}</td>
    <td class="tg-{SESSION_TUESDAY}">{TUESDAY_2}</td>
    <td class="tg-0pky">{WEDNESDAY_1}</td>
    <td class="tg-{SESSION_WEDNESDAY}">{WEDNESDAY_2}</td>
    <td class="tg-0pky">{THURSDAY_1}</td>
    <td class="tg-{SESSION_THURSDAY}">{THURSDAY_2}</td>
  </tr>
"""
html_rows = []
for idx, row in df.iterrows():
    # Bypass the header row
    if idx == 0:
        continue
    # Each triplet in the row is a time slot, speaker surname, and session type
    # We have to guard against having empty types
    try:
        session_monday = p.number_to_words(int(row.iloc[2]))
    except ValueError:
        session_monday = "zero"
    try:
        session_tuesday = p.number_to_words(int(row.iloc[5]))
    except ValueError:
        session_tuesday = "zero"
    try:
        session_wednesday = p.number_to_words(int(row.iloc[8]))
    except ValueError:
        session_wednesday = "zero"
    try:
        session_thursday = p.number_to_words(int(row.iloc[11]))
    except ValueError:
        session_thursday = "zero"
    html_row = HTML_TABLE_ROW_TEMPLATE.format(
        MONDAY_1=row.iloc[0],
        MONDAY_2=row.iloc[1],
        SESSION_MONDAY=session_monday,
        TUESDAY_1=row.iloc[3],
        TUESDAY_2=row.iloc[4],
        SESSION_TUESDAY=session_tuesday,
        WEDNESDAY_1=row.iloc[6],
        WEDNESDAY_2=row.iloc[7],
        SESSION_WEDNESDAY=session_wednesday,
        THURSDAY_1=row.iloc[9],
        THURSDAY_2=row.iloc[10],
        SESSION_THURSDAY=session_thursday,
    )
    html_rows.append(html_row)
HTML_TABLE_TEMPLATE = HTML_TABLE_TEMPLATE.replace("{BODY}", "\n".join(html_rows))
output_path = Path(OUTPUT_FILE)
output_path.write_text(HTML_TABLE_TEMPLATE, encoding="utf-8")

# Now to modify the abstract markdown file with the times
ABSTRACT_FILE = Path(__file__).parent.parent.parent / "abstracts.md"
abstract_md = ABSTRACT_FILE.read_text(encoding="utf-8")
new_abstract_md = ""
for entry in abstract_md.split("* "):
    if "**Author**: " in entry:
        # Have to try and account for those with multiple surnames
        surname = entry.split("**Author**: ")[1].split("\n")[0].strip().split(" ")[1:]
        surname = " ".join(surname)
        # Find the name in the dataframe
        match = df[
            (df.iloc[:, 1].str.contains(surname, case=True))
            | (df.iloc[:, 4].str.contains(surname, case=True))
            | (df.iloc[:, 7].str.contains(surname, case=True))
            | (df.iloc[:, 10].str.contains(surname, case=True))
        ]
        if not match.empty:
            # Get column idx which matches the surname
            col_idx = None
            # Assume they only speak once
            for col in [1, 4, 7, 10]:
                if match.iloc[0, col] == surname:
                    col_idx = col - 1
                    break
            if col_idx is None:
                raise ValueError(
                    f"No matching column index found for author surname {surname}, full name is {entry}"
                )
            day = ["Monday", "Tuesday", "Wednesday", "Thursday"][col_idx // 3]
            day_and_time = f"{day} - {match.iloc[0, col_idx]}"
            new_entry = entry.replace("**When**: FILL IN", f"**When**: {day_and_time}")
            new_abstract_md += "* " + new_entry
    else:
        new_abstract_md += entry
ABSTRACT_FILE.write_text(new_abstract_md, encoding="utf-8")
