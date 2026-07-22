# TELEGRAM_AUTO IMPORTS ================= #

from telethon import TelegramClient, events
from telethon.sessions import StringSession
from openpyxl import Workbook

from remedy_bot import (
    register_remedy_handler,
    #remedy_followup_checker,
    summary_checker,
    leaders_summary_checker,
    sync_open_incidents
)

# Optional dashboard integration intentionally excluded from this sanitized sample.
import re
import asyncio
from datetime import datetime, timedelta, timezone
import os


# ================= TELEGRAM ================= #

api_id = int(os.getenv("API_ID"))
api_hash = os.getenv("API_HASH")
SESSION_STRING = os.getenv("SESSION_STRING")

client = TelegramClient(
    StringSession(SESSION_STRING),
    api_id,
    api_hash
)
register_remedy_handler(client)

SOURCE_CHAT = int(os.getenv("SOURCE_CHAT_ID", "0"))
TARGET_CHAT = int(os.getenv("TARGET_CHAT_ID", "0"))

IST = timezone(timedelta(hours=5, minutes=30))

def now_ist():
    return datetime.now(IST)


# ================= CHAT ID CAPTURE ================= #

#seen_chats = set()

#@client.on(events.NewMessage())
#async def capture_chat_ids(event):
#    try:
#        chat = await event.get_chat()
#        sender = await event.get_sender()

#        if hasattr(chat, "id") and chat.id not in seen_chats:
#            seen_chats.add(chat.id)

#            chat_name = getattr(chat, "title", "Private Chat")
#            sender_name = getattr(sender, "first_name", "Unknown")

#            print(f"\n📌 CHAT NAME : {chat_name}")
#            print(f"📌 CHAT ID   : {chat.id}")
#            print(f"👤 SENDER    : {sender_name}")
#            print(f"💬 MESSAGE   : {event.raw_text}")
#            print("-" * 60)

#    except Exception as e:
#        print("Error:", e)


# ================= ZONE MAPPING ================= #

# Replace the sample placeholders with approved configuration values.
# Do not commit real leader names or Telegram user IDs.
ZONE_MAPPING = {
    "AREA_1": {"name": "AREA_1_LEADER", "id": int(os.getenv("AREA_1_LEADER_CHAT_ID", "0"))},
    "AREA_2": {"name": "AREA_2_LEADER", "id": int(os.getenv("AREA_2_LEADER_CHAT_ID", "0"))},
    "AREA_3": {"name": "AREA_3_LEADER", "id": int(os.getenv("AREA_3_LEADER_CHAT_ID", "0"))},
}

# ================= COH MAPPING ================= #

COH_MAPPING = {
    "AREA_1": {"name": "AREA_1_CIRCLE_LEADER", "id": int(os.getenv("AREA_1_CIRCLE_LEADER_CHAT_ID", "0"))},
    "AREA_2": {"name": "AREA_2_CIRCLE_LEADER", "id": int(os.getenv("AREA_2_CIRCLE_LEADER_CHAT_ID", "0"))},
    "AREA_3": {"name": "AREA_3_CIRCLE_LEADER", "id": int(os.getenv("AREA_3_CIRCLE_LEADER_CHAT_ID", "0"))},
}

active_outages = {}


# ================= TIME HELPERS ================= #

def parse_time(text, label):
    for line in text.split("\n"):
        if label in line:
            match = re.search(r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2})", line)
            if match:
                return datetime.strptime(match.group(1), "%d/%m/%Y %H:%M").replace(tzinfo=IST)
    return None


def format_duration(start_time):
    diff = now_ist() - start_time
    if diff.total_seconds() < 0:
        return "0:00 Hrs"
    hrs = int(diff.total_seconds() // 3600)
    mins = int((diff.total_seconds() % 3600) // 60)
    return f"{hrs}:{mins:02d} Hrs"


def format_duration_between(start, end):
    if not start or not end:
        return "0:00 Hrs"
    diff = end - start
    hrs = int(diff.total_seconds() // 3600)
    mins = int((diff.total_seconds() % 3600) // 60)
    return f"{hrs}:{mins:02d} Hrs"


# ================= FOLLOW-UP CHECKER ================= #

async def followup_checker():
    while True:
        now = now_ist()

        for key, data in list(active_outages.items()):
            if now - data["last_alert"] >= timedelta(hours=1):

                duration = format_duration(data["start_time"])

                msg  = f"⏰ Follow-up: {data['severity']} Outage Still Running - {data['zones_text']}\n"
                msg += f"📉 Sites: {data['site_count']}\n"
                msg += f"⏱ Down Since: {data['start_time'].strftime('%d/%m/%Y %H:%M')}\n"
                msg += f"⏳ Duration: {duration}\n"
                msg += f"🛠 Fault: {data['fault']}\n"
                msg += f"{data['tags_text']}"

                # COH tag if outage > 4 hrs
                diff = now - data["start_time"]
                if diff.total_seconds() >= 4 * 3600:
                    coh_set = set()
                    for z in data["zones_text"].split(", "):
                        coh = COH_MAPPING.get(z)
                        if coh:
                            coh_set.add((coh["name"], coh["id"]))

                    if coh_set:
                        coh_tags = [
                            f'<a href="tg://user?id={cid}">{cname}</a>'
                            for cname, cid in coh_set
                        ]
                        msg += "\n\n⚠️ Dear " + " / ".join(coh_tags) + " Sir Kindly intervene as outage running since > 4 Hrs."

                await client.send_message(TARGET_CHAT, msg, parse_mode="html")
                data["last_alert"] = now

        await asyncio.sleep(600)


# ================= MAIN MESSAGE HANDLER ================= #

@client.on(events.NewMessage(chats=SOURCE_CHAT))
async def handler(event):
    text = event.raw_text
    text_upper = text.upper()

    match_type = re.search(r"NOTIFICATION\s*TYPE\s*:\s*(INT|UPD|FIN)", text_upper)
    if not match_type:
        return

    notif_type = match_type.group(1)
    status = "RUNNING" if notif_type in ["INT", "UPD"] else "RESTORED"

    site_match = re.search(r"(\d+)\s*sites", text, re.I)
    site_count = int(site_match.group(1)) if site_match else 0

    impact_text = next((l for l in text_upper.split("\n") if "END USER IMPACT" in l), "")

    zones = [z for z in ZONE_MAPPING if re.search(rf"\b{z}\b", impact_text)]
    if not zones:
        return

    zones_sorted = sorted(set(zones))
    zones_text = ", ".join(zones_sorted)

    tags = [
        f'<a href="tg://user?id={ZONE_MAPPING[z]["id"]}">👤 {ZONE_MAPPING[z]["name"]}</a>'
        for z in zones_sorted
    ]
    tags_text = " | ".join(tags)

    start_time = parse_time(text, "Start Time") or now_ist()
    stop_time  = parse_time(text, "Stop Time")

    fault_match = re.search(r"Fault Cause:(.*)", text)
    fault = fault_match.group(1).strip() if fault_match else "NA"

    severity = (
        "🔴 CRITICAL" if site_count >= 20 else
        "🟠 MAJOR"    if site_count >= 10 else
        "🟡 MINOR"
    )

    key = "|".join(zones_sorted)

    # ── RUNNING ──────────────────────────────────────────── #
    if status == "RUNNING":
        duration = format_duration(start_time)

        if key not in active_outages:
            msg  = f"{severity} Outage Running - {zones_text}\n"
            msg += f"📉 Sites: {site_count}\n"
            msg += f"⏱ Down Since: {start_time.strftime('%d/%m/%Y %H:%M')}\n"
            msg += f"⏳ Duration: {duration}\n"
            msg += f"🛠 Fault: {fault}\n"
            msg += f"{tags_text}"

            await client.send_message(TARGET_CHAT, msg, parse_mode="html")

            active_outages[key] = {
                "start_time": start_time,
                "last_alert": now_ist(),
                "site_count": site_count,
                "fault":      fault,
                "severity":   severity,
                "tags_text":  tags_text,
                "zones_text": zones_text,
            }
        else:
            active_outages[key]["site_count"] = site_count
            active_outages[key]["fault"]      = fault

    # ── RESTORED ─────────────────────────────────────────── #
    else:
        if key in active_outages:
            start_time = active_outages[key]["start_time"]

        duration = format_duration_between(start_time, stop_time)

        msg  = f"✅ Outage Restored - {zones_text}\n"
        msg += f"📈 Sites: {site_count}\n"
        msg += f"⏱ Down: {start_time.strftime('%d/%m/%Y %H:%M')}\n"

        if stop_time:
            msg += f"⏱ Up: {stop_time.strftime('%d/%m/%Y %H:%M')}\n"

        msg += f"⏳ Total Duration: {duration}\n"
        msg += f"🛠 Fault: {fault}\n"
        msg += f"{tags_text}"

        await client.send_message(TARGET_CHAT, msg, parse_mode="html")

        if key in active_outages:
            del active_outages[key]


# ================= AREA MAPPING ================= #

AREA_MAPPING = {
    "AREA_1": "AREA_1",
    "AREA_2": "AREA_2",
    "AREA_3": "AREA_3",
}

# ================= CYCLE RANGE ================= #

from datetime import datetime, timedelta

def get_cycle_range(now):

    if now.day >= 26:

        start = datetime(
            now.year,
            now.month,
            26,
            0,
            0,
            tzinfo=IST
        )

        if now.month == 12:
            end = datetime(
                now.year + 1,
                1,
                25,
                23,
                59,
                59,
                tzinfo=IST
            )
        else:
            end = datetime(
                now.year,
                now.month + 1,
                25,
                23,
                59,
                59,
                tzinfo=IST
            )

    else:

        if now.month == 1:

            start = datetime(
                now.year - 1,
                12,
                26,
                0,
                0,
                tzinfo=IST
            )

        else:

            start = datetime(
                now.year,
                now.month - 1,
                26,
                0,
                0,
                tzinfo=IST
            )

        end = datetime(
            now.year,
            now.month,
            25,
            23,
            59,
            59,
            tzinfo=IST
        )

    return start, end


# ================= REPORT GENERATOR ================= #

async def generate_backup_report():
    print("📊 GENERATING REPORT...")

    wb = Workbook()
    ws = wb.active
    ws.title = "Backup"

    ws.append([
        "Date", "TT", "Description", "Impact", "Start Time", "Stop Time",
        "Zone", "Area", "2G", "4G", "Total Site",
        "Duration Hrs", "Duration Min", "CMO", "CMO Mn", "Fault"
    ])

    rows_added = 0  # ✅ FIXED POSITION

    now = now_ist()
    start_range, end_range = get_cycle_range(now)

    print(f"Cycle: {start_range} → {end_range}")

    async for msg in client.iter_messages(SOURCE_CHAT, reverse=False, limit=8000):
        if not msg.text:
            continue

        msg_time = msg.date.astimezone(IST)

        if msg_time < start_range:
            continue
        if msg_time > end_range:
            continue

        text = msg.text
        u = text.upper()

        # ✅ FIX 1: Only check FIN (remove NOTIFICATION TYPE strict)
        if "FIN" not in u:
            continue

        # ✅ FIX 2: REMOVE strict fault filter OR make optional
        fault_match = re.search(r"Fault Cause:(.*)", text)
        fault = fault_match.group(1).strip() if fault_match else "UNKNOWN"

        # OPTIONAL filter (keep if needed)
        # if "TNG MEDIA" not in fault.upper():
        #     continue

        impact = next((l for l in u.split("\n") if "END USER IMPACT" in l), "")

        zone = None
        for z in ZONE_MAPPING:
            if z in impact:
                zone = z
                break

        if not zone:
            continue

        area = AREA_MAPPING.get(zone, "NA")

        start_time = parse_time(text, "Start Time")
        stop_time  = parse_time(text, "Stop Time")

        if not start_time or not stop_time:
            continue

        g2 = int(re.search(r"2G-(\d+)", text).group(1)) if re.search(r"2G-(\d+)", text) else 0
        g4 = int(re.search(r"4G-(\d+)", text).group(1)) if re.search(r"4G-(\d+)", text) else 0

        site = g2 + g4

        duration_min = int((stop_time - start_time).total_seconds() / 60)
        duration_hr = round(duration_min / 60, 2)

        cmo = (g2 * duration_min * 4) + (g4 * duration_min * 8)
        cmo_mn = round(cmo / 1_000_000, 2)

        tt_match = re.search(r"TT[:\s\-]*([A-Z0-9]+)", text)
        tt = tt_match.group(1) if tt_match else ""

        desc = next((l for l in text.split("\n") if "Description" in l), "")

        ws.append([
            start_time.strftime("%d-%b-%Y"),
            tt, desc, impact,
            start_time.strftime("%d/%m/%Y %H:%M"),
            stop_time.strftime("%d/%m/%Y %H:%M"),
            zone, area, g2, g4, site,
            duration_hr, duration_min,
            int(cmo), cmo_mn, fault
        ])

        rows_added += 1

    print(f"✅ TOTAL ROWS: {rows_added}")

    if rows_added == 0:
        print("⚠️ No data matched filters")

    file = "Daily_Backup_Report.xlsx"
    wb.save(file)

    await client.send_file(TARGET_CHAT, file)
    print("✅ REPORT SENT")

# ================= FETCH DIALOGS ================= #

async def fetch_dialogs(client):

    print("\n📋 AVAILABLE DIALOGS\n")

    async for dialog in client.iter_dialogs():

        print(
            f"{dialog.name} --> {dialog.id}"
        )

# ================= REPORT SCHEDULER ================= #

last_report_date = None

async def report_scheduler():

    global last_report_date

    print("📊 Report Scheduler Started")

    await asyncio.sleep(30)

    while True:

        try:

            now = now_ist()

            if (
                now.hour == 8
                and 0 <= now.minute < 9
                and last_report_date != now.date()
            ):

                print("📊 DAILY REPORT TRIGGERED")

                await generate_backup_report()

                last_report_date = now.date()

                print("✅ DAILY REPORT COMPLETED")

        except Exception as e:

            print(
                f"❌ REPORT SCHEDULER ERROR : {e}"
            )

        await asyncio.sleep(30)
# ================= ENTRY POINT ================= #

async def main():

    print("🚀 LIVE BOT STARTED")

    # ================= FOLLOWUP ================= #

    client.loop.create_task(
        followup_checker()
    )

    # ================= BACKUP REPORT ================= #

    client.loop.create_task(
        report_scheduler()
    )

    # ================= TEST ONLY ================= #
    # Uncomment once to verify report generation
    #
    client.loop.create_task(
        generate_backup_report()
    )

    # ================= REMEDY ================= #

    # client.loop.create_task(
    #     remedy_followup_checker(client)
    # )

    # ================= SUMMARY ================= #

    client.loop.create_task(
        summary_checker(client)
    )

    client.loop.create_task(
        leaders_summary_checker(client)
    )

    # ================= SYNC OPEN INCIDENTS ================= #

    try:

        print("🔄 Syncing Open Incidents...")

        await sync_open_incidents(client)

        print("✅ Open Incidents Synced")

    except Exception as e:

        print(
            f"❌ sync_open_incidents ERROR : {e}"
        )

    # ================= KEEP BOT RUNNING ================= #

    print("✅ BOT READY")

    await client.run_until_disconnected()


# ================= START BOT ================= #

with client:

    try:

        client.loop.run_until_complete(
            main()
        )

    except Exception as e:

        print(
            f"❌ MAIN ERROR : {e}"
        )
