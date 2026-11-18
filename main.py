
import feedparser
import os
from datetime import datetime
from zoneinfo import ZoneInfo
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

CHANNEL_FILE = "channels.txt"
EXCEL_FILE ="youtube_tracker_redesigned.xlsx"

HEADERS = [
    "Channel Name",
    "Video ID",
    "Video Title",
    "Upload Date",
    "Upload Time (IST)",
    "Published (UTC)",
    "Video URL"
]


def get_channels():
    with open(CHANNEL_FILE, "r", encoding="utf-8") as file:
        return [
            line.strip()
            for line in file
            if line.strip() and not line.startswith("#")
        ]


def open_excel():
    if os.path.exists(EXCEL_FILE):
        workbook = load_workbook(EXCEL_FILE)
        sheet = workbook["Video Library"]
    else:
        workbook = Workbook()
        sheet = workbook["Video Library"]
        sheet.title = "Videos"
        sheet.append(HEADERS)

        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(
                "solid", fgColor="1565C0"
            )
            cell.alignment = Alignment(horizontal="center")

        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = "A1:G1"

    return workbook, sheet


def check_videos():
    workbook, sheet = open_excel()

    existing_ids = {
        str(row[1].value)
        for row in sheet.iter_rows(min_row=2, max_col=2)
        if row[1].value
    }

    channels = get_channels()
    new_count = 0

    for channel_id in channels:
        rss_url = (
            "https://www.youtube.com/feeds/videos.xml"
            f"?channel_id={channel_id}"
        )

        feed = feedparser.parse(rss_url)

        if not feed.entries:
            print("No videos found:", channel_id)
            continue

        channel_name = feed.feed.get("title", "Unknown")

        print("\nChecking:", channel_name)

        for video in feed.entries:
            video_id = video.get("yt_videoid")

            if not video_id:
                continue

            if video_id in existing_ids:
                continue

            published = datetime.fromisoformat(
                video.published.replace("Z", "+00:00")
            )

            ist_time = published.astimezone(
                ZoneInfo("Asia/Kolkata")
            )

            sheet.append([
                channel_name,
                video_id,
                video.title,
                ist_time.strftime("%d-%m-%Y"),
                ist_time.strftime("%I:%M %p"),
                video.published,
                video.link
            ])

            sheet.cell(
                row=sheet.max_row,
                column=7
            ).hyperlink = video.link

            existing_ids.add(video_id)
            new_count += 1

            print("NEW VIDEO:", video.title)

    # Adjust Excel column widths
    widths = [30, 22, 65, 16, 20, 30, 55]

    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[
            chr(64 + index)
        ].width = width

    sheet.auto_filter.ref = sheet.dimensions

    workbook.save(EXCEL_FILE)
    workbook.close()

    print("\n------------------------------")
    print("Excel database updated!")
    print("New videos added:", new_count)
    print("Excel file:", EXCEL_FILE)


if __name__ == "__main__":
    check_videos()
