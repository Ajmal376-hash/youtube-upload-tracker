
import feedparser
import os
import time
from datetime import datetime
from zoneinfo import ZoneInfo
from openpyxl import load_workbook
from winotify import Notification

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CHANNEL_FILE = os.path.join(BASE_DIR, "channels.txt")
EXCEL_FILE = EXCEL_FILE = "youtube_tracker_redesigned.xlsx"

CHECK_INTERVAL = 300  # 5 minutes

def notify_user(channel, title):
    toast = Notification(
        app_id="YouTube Tracker",
        title="New YouTube Video!",
        msg=f"{channel}\n{title}",
        duration="long"
    )
    toast.show()


def check_new_videos():

    if not os.path.exists(EXCEL_FILE):
        print("Excel database not found!")
        return

    workbook = load_workbook(EXCEL_FILE)
    sheet = workbook["Video Library"]

    existing_ids = {
        str(row[1].value)
        for row in sheet.iter_rows(min_row=2, max_col=2)
        if row[1].value
    }

    with open(CHANNEL_FILE, "r", encoding="utf-8") as file:
        channels = [
            line.strip()
            for line in file
            if line.strip() and not line.startswith("#")
        ]

    new_count = 0

    for channel_id in channels:

        rss_url = (
            "https://www.youtube.com/feeds/videos.xml"
            f"?channel_id={channel_id}"
        )

        try:
            feed = feedparser.parse(rss_url)

            if not feed.entries:
                print("No feed entries:", channel_id)
                continue

            channel_name = feed.feed.get("title", "Unknown")

            for video in feed.entries:

                video_id = video.get("yt_videoid")

                if not video_id or video_id in existing_ids:
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

                notify_user(channel_name, video.title)

        except Exception as error:
            print("Channel error:", channel_id, error)

    workbook.save(EXCEL_FILE)
    workbook.close()

    print("Check completed.")
    print("New videos:", new_count)


if __name__ == "__main__":

    print("================================")
    print("      YOUTUBE TRACKER")
    print("================================")
    print("Automatic monitoring started.")
    print("Checking interval: 5 minutes")

    while True:

        try:
            print("\nChecking YouTube:", datetime.now())
            check_new_videos()

        except Exception as error:
            print("Monitoring error:", error)

        print("Next check in 5 minutes...")
        time.sleep(CHECK_INTERVAL)
