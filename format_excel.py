
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, date

FILE = "youtube_tracker_redesigned.xlsx"

NAVY = "172B4D"
BLUE = "3678B8"
PALE = "EAF1F8"
WHITE = "FFFFFF"
DARK = "273444"
GREY = "F4F6F8"
MUTED = "718096"
LINE = "E2E8F0"

wb = load_workbook(FILE)

# Remove the previous dashboard
if "Dashboard" in wb.sheetnames:
    del wb["Dashboard"]

# Remove the accidental empty default sheet
if "Sheet" in wb.sheetnames:
    old = wb["Sheet"]
    if old.max_row == 1 and old.max_column == 1 and old["A1"].value is None:
        del wb["Sheet"]

library = wb["Video Library"]
directory = wb["Channel Directory"]

videos = []
for row in library.iter_rows(min_row=2, values_only=True):
    if row and row[1]:
        videos.append(row)

channels = []
for row in directory.iter_rows(min_row=2, values_only=True):
    if row and row[1]:
        channels.append(row)

def get_sort_date(row):
    published = row[5]
    try:
        return datetime.fromisoformat(
            str(published).replace("Z", "+00:00")
        )
    except (ValueError, TypeError):
        return datetime.min

videos.sort(key=get_sort_date, reverse=True)

# Create dashboard
ws = wb.create_sheet("Dashboard", 0)
ws.sheet_view.showGridLines = False
ws.sheet_view.zoomScale = 85

# Wider, balanced page layout
widths = {
    "A": 30, "B": 42, "C": 18, "D": 18,
    "E": 24, "F": 16, "G": 20, "H": 20
}
for col, width in widths.items():
    ws.column_dimensions[col].width = width

# Title
ws.merge_cells("A1:H1")
ws["A1"] = "YOUTUBE UPLOAD TRACKER"
ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
ws["A1"].font = Font(
    name="Aptos Display", size=22, bold=True, color=WHITE
)
ws["A1"].alignment = Alignment(
    horizontal="left", vertical="center", indent=1
)
ws.row_dimensions[1].height = 58

ws.merge_cells("A2:H2")
ws["A2"] = "PERSONAL CHANNEL MONITORING  /  OVERVIEW"
ws["A2"].font = Font(name="Aptos", size=9, color=MUTED)
ws["A2"].alignment = Alignment(
    horizontal="left", vertical="center", indent=2
)
ws.row_dimensions[2].height = 28

# Three summary cards
latest_date = None
if videos:
    latest_date = videos[0][3]

cards = [
    ("A4:B4", "A5:B6", "TOTAL VIDEOS", len(videos)),
    ("D4:E4", "D5:E6", "CHANNELS TRACKED", len(channels)),
    ("G4:H4", "G5:H6", "LATEST UPLOAD",
     latest_date if latest_date else "No uploads")
]

for label_range, value_range, label, value in cards:
    ws.merge_cells(label_range)
    ws.merge_cells(value_range)

    label_cell = ws[label_range.split(":")[0]]
    value_cell = ws[value_range.split(":")[0]]

    label_cell.value = label
    label_cell.fill = PatternFill("solid", fgColor=PALE)
    label_cell.font = Font(
        name="Aptos", size=9, bold=True, color=BLUE
    )
    label_cell.alignment = Alignment(
        horizontal="center", vertical="center"
    )

    value_cell.value = value
    value_cell.fill = PatternFill("solid", fgColor=GREY)
    value_cell.font = Font(
        name="Aptos Display", size=22, bold=True, color=NAVY
    )
    value_cell.alignment = Alignment(
        horizontal="center", vertical="center"
    )

    if isinstance(value, (datetime, date)):
        value_cell.number_format = "dd-mmm-yyyy"

ws.row_dimensions[4].height = 27
ws.row_dimensions[5].height = 30
ws.row_dimensions[6].height = 27

# Recent uploads section
ws.merge_cells("A9:F9")
ws["A9"] = "RECENT UPLOADS"
ws["A9"].fill = PatternFill("solid", fgColor=NAVY)
ws["A9"].font = Font(bold=True, color=WHITE, size=11)
ws["A9"].alignment = Alignment(
    horizontal="left", vertical="center", indent=1
)
ws.row_dimensions[9].height = 32

headers = [
    "CHANNEL", "VIDEO TITLE", "UPLOAD DATE",
    "TIME (IST)", "VIDEO ID", "LINK"
]

for col, value in enumerate(headers, 1):
    cell = ws.cell(10, col, value)
    cell.fill = PatternFill("solid", fgColor=BLUE)
    cell.font = Font(bold=True, color=WHITE, size=9)
    cell.alignment = Alignment(
        horizontal="left", vertical="center"
    )

ws.row_dimensions[10].height = 27

for index, video in enumerate(videos[:5], start=11):
    channel, video_id, title, upload_date, upload_time, published, url = video[:7]

    values = [
        channel, title, upload_date, upload_time,
        video_id, "Watch video"
    ]

    for col, value in enumerate(values, 1):
        cell = ws.cell(index, col, value)
        cell.font = Font(name="Aptos", size=9, color=DARK)
        cell.alignment = Alignment(
            vertical="center",
            wrap_text=(col in [1, 2])
        )

        if index % 2 == 0:
            cell.fill = PatternFill("solid", fgColor=GREY)

        cell.border = Border(
            bottom=Side(style="hair", color=LINE)
        )

    ws.cell(index, 3).number_format = "dd-mmm-yyyy"
    ws.cell(index, 4).number_format = "hh:mm AM/PM"

    if url:
        ws.cell(index, 6).hyperlink = str(url)
        ws.cell(index, 6).style = "Hyperlink"

    ws.row_dimensions[index].height = 34

# Channel summary
ws.merge_cells("A19:C19")
ws["A19"] = "CHANNEL OVERVIEW"
ws["A19"].fill = PatternFill("solid", fgColor=NAVY)
ws["A19"].font = Font(bold=True, color=WHITE, size=11)
ws["A19"].alignment = Alignment(
    horizontal="left", vertical="center", indent=1
)
ws.row_dimensions[19].height = 32

for col, value in enumerate(
    ["CHANNEL NAME", "VIDEOS COLLECTED", "STATUS"], 1
):
    cell = ws.cell(20, col, value)
    cell.fill = PatternFill("solid", fgColor=BLUE)
    cell.font = Font(bold=True, color=WHITE, size=9)
    cell.alignment = Alignment(vertical="center")

ws.row_dimensions[20].height = 27

for index, channel in enumerate(channels, start=21):
    name = channel[0]
    count = channel[2] if len(channel) > 2 else 0
    status = channel[3] if len(channel) > 3 else "Active"

    for col, value in enumerate([name, count, status], 1):
        cell = ws.cell(index, col, value)
        cell.font = Font(name="Aptos", size=9, color=DARK)
        cell.alignment = Alignment(vertical="center")

        if index % 2 == 0:
            cell.fill = PatternFill("solid", fgColor=GREY)

        cell.border = Border(
            bottom=Side(style="hair", color=LINE)
        )

    ws.row_dimensions[index].height = 25

# Keep dashboard simple: no frozen panes, no charts
ws.freeze_panes = None
ws.sheet_view.showGridLines = False
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.page_setup.orientation = "landscape"

# Remove Excel Table objects that are causing repair warnings
for sheet in [library, directory]:
    for table_name in list(sheet.tables.keys()):
        del sheet.tables[table_name]

    # Keep normal filtering without Excel Table objects
    sheet.auto_filter.ref = sheet.dimensions

wb.active = 0
wb.save(FILE)
wb.close()

print("Dashboard redesigned successfully!")
print("Videos:", len(videos))
print("Channels:", len(channels))
print("Updated:", FILE)
