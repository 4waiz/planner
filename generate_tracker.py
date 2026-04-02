import xlsxwriter
from datetime import datetime, timedelta

# --- Task Data ---
raw_tasks = [
    {"id": 1, "category": "Amazon Delivery", "task": "Nilight 50029R 120 Pcs Standard Blade Fuse Assorted Set with 10 Pack 14AWG ATC/ATO Inline Holder", "source": "Amazon", "start_date": "", "due_date": "21 April 2026", "time_zone": "Emirates Time", "status": "Arriving", "priority": "Medium", "notes": "Return eligible through 4 May 2026"},
    {"id": 2, "category": "Amazon Delivery", "task": "TUOFENG 16 AWG Electrical Wire 3m [1.5m Black and 1.5m Red]", "source": "Amazon", "start_date": "", "due_date": "21 April 2026", "time_zone": "Emirates Time", "status": "Arriving", "priority": "Medium", "notes": "Return eligible through 4 May 2026"},
    {"id": 3, "category": "Amazon Delivery", "task": "DROK Buck Converter 12v to 5v Voltage Regulator Board", "source": "Amazon", "start_date": "", "due_date": "21 April 2026", "time_zone": "Emirates Time", "status": "Arriving", "priority": "Medium", "notes": "Return eligible through 4 May 2026"},
    {"id": 4, "category": "Amazon Delivery", "task": "Fancasee Replacement 5.5mm x 2.5mm 90 Degree Right Angle DC Power Male Plug Jack to Bare Wire Cable", "source": "Amazon", "start_date": "", "due_date": "21 April 2026", "time_zone": "Emirates Time", "status": "Arriving", "priority": "Medium", "notes": "Return eligible through 4 May 2026"},
    {"id": 5, "category": "Amazon Delivery", "task": "NVIDIA Jetson Orin Nano Developer Kit", "source": "Amazon", "start_date": "", "due_date": "18 April 2026", "time_zone": "Emirates Time", "status": "Arriving", "priority": "High", "notes": "Return eligible through 2 May 2026"},
    {"id": 6, "category": "Amazon Delivery", "task": "Comidox USB Logic Analyzer Device Set USB Cable 24MHz 8CH UART IIC SPI Debug", "source": "Amazon", "start_date": "", "due_date": "18 April 2026", "time_zone": "Emirates Time", "status": "Arriving", "priority": "High", "notes": "No return date shown in screenshot"},
    {"id": 7, "category": "Project", "task": "Build and integrate everything for GuideBot - BRIDGE", "source": "Project", "start_date": "", "due_date": "21 April 2026", "time_zone": "Emirates Time", "status": "Pending", "priority": "Urgent", "notes": "Final integration deadline"},
    {"id": 8, "category": "Reminder Window", "task": "Emirates reminder period", "source": "Personal", "start_date": "4 May 2026", "due_date": "7 May 2026", "time_zone": "Emirates Time", "status": "Planned", "priority": "Medium", "notes": "Reminder window from 4 May to 7 May"},
]

today = datetime.now()

def parse_date(date_str):
    try:
        return datetime.strptime(date_str, "%d %B %Y")
    except ValueError:
        return None

def get_reminder_level(days_left, priority, status):
    if status == "Completed":
        return "Completed"
    elif days_left is not None and days_left < 0:
        return "Overdue"
    elif priority == "Urgent":
        return "Urgent"
    elif days_left is not None and days_left <= 3:
        return "Soon"
    elif days_left is not None and days_left <= 7:
        return "Upcoming"
    else:
        return "On Track"

# --- Process tasks ---
for task in raw_tasks:
    due = parse_date(task["due_date"])
    if due:
        task["days_left"] = (due - today).days
    else:
        task["days_left"] = None
    task["reminder_level"] = get_reminder_level(task["days_left"], task["priority"], task["status"])

# --- Create Excel workbook ---
workbook = xlsxwriter.Workbook("Delivery_and_Project_Reminder_Tracker.xlsx")
worksheet = workbook.add_worksheet("Tracker")

# --- Formats ---
title_fmt = workbook.add_format({
    "bold": True, "font_size": 18, "font_color": "#1e293b",
    "bottom": 2, "bottom_color": "#6366f1",
})
subtitle_fmt = workbook.add_format({
    "font_size": 11, "font_color": "#64748b", "italic": True,
})
header_fmt = workbook.add_format({
    "bold": True, "font_size": 11, "font_color": "#ffffff",
    "bg_color": "#1e293b", "border": 1, "border_color": "#334155",
    "text_wrap": True, "valign": "vcenter", "align": "center",
})

# Row color formats based on reminder level
row_formats = {
    "Overdue":   workbook.add_format({"bg_color": "#fef2f2", "font_color": "#7f1d1d", "border": 1, "border_color": "#fecaca", "text_wrap": True, "valign": "vcenter"}),
    "Urgent":    workbook.add_format({"bg_color": "#fef2f2", "font_color": "#7f1d1d", "border": 1, "border_color": "#fecaca", "text_wrap": True, "valign": "vcenter"}),
    "Soon":      workbook.add_format({"bg_color": "#fff7ed", "font_color": "#7c2d12", "border": 1, "border_color": "#fed7aa", "text_wrap": True, "valign": "vcenter"}),
    "Upcoming":  workbook.add_format({"bg_color": "#fefce8", "font_color": "#713f12", "border": 1, "border_color": "#fef08a", "text_wrap": True, "valign": "vcenter"}),
    "On Track":  workbook.add_format({"bg_color": "#f0fdf4", "font_color": "#14532d", "border": 1, "border_color": "#bbf7d0", "text_wrap": True, "valign": "vcenter"}),
    "Completed": workbook.add_format({"bg_color": "#ecfdf5", "font_color": "#064e3b", "border": 1, "border_color": "#a7f3d0", "text_wrap": True, "valign": "vcenter"}),
}

days_left_formats = {
    "Overdue":   workbook.add_format({"bg_color": "#ef4444", "font_color": "#ffffff", "bold": True, "border": 1, "align": "center", "valign": "vcenter", "num_format": "0"}),
    "Urgent":    workbook.add_format({"bg_color": "#ef4444", "font_color": "#ffffff", "bold": True, "border": 1, "align": "center", "valign": "vcenter", "num_format": "0"}),
    "Soon":      workbook.add_format({"bg_color": "#f97316", "font_color": "#ffffff", "bold": True, "border": 1, "align": "center", "valign": "vcenter", "num_format": "0"}),
    "Upcoming":  workbook.add_format({"bg_color": "#eab308", "font_color": "#ffffff", "bold": True, "border": 1, "align": "center", "valign": "vcenter", "num_format": "0"}),
    "On Track":  workbook.add_format({"bg_color": "#22c55e", "font_color": "#ffffff", "bold": True, "border": 1, "align": "center", "valign": "vcenter", "num_format": "0"}),
    "Completed": workbook.add_format({"bg_color": "#10b981", "font_color": "#ffffff", "bold": True, "border": 1, "align": "center", "valign": "vcenter", "num_format": "0"}),
}

priority_formats = {
    "Urgent": workbook.add_format({"bg_color": "#dc2626", "font_color": "#ffffff", "bold": True, "border": 1, "align": "center", "valign": "vcenter"}),
    "High":   workbook.add_format({"bg_color": "#f97316", "font_color": "#ffffff", "bold": True, "border": 1, "align": "center", "valign": "vcenter"}),
    "Medium": workbook.add_format({"bg_color": "#3b82f6", "font_color": "#ffffff", "bold": True, "border": 1, "align": "center", "valign": "vcenter"}),
    "Low":    workbook.add_format({"bg_color": "#22c55e", "font_color": "#ffffff", "bold": True, "border": 1, "align": "center", "valign": "vcenter"}),
}

# --- Title ---
worksheet.merge_range("A1:K1", "Delivery and Project Reminder Tracker", title_fmt)
worksheet.merge_range("A2:K2", f"Emirates Time  |  Generated: {today.strftime('%d %B %Y')}", subtitle_fmt)

# --- Headers ---
headers = ["#", "Category", "Task / Item", "Source", "Start Date", "Due / Arrival Date", "Time Zone", "Status", "Priority", "Days Left", "Reminder Level", "Notes"]
col_widths = [4, 18, 50, 10, 14, 20, 16, 12, 12, 12, 16, 40]

for col, (header, width) in enumerate(zip(headers, col_widths)):
    worksheet.write(3, col, header, header_fmt)
    worksheet.set_column(col, col, width)

# --- Data Rows ---
for i, task in enumerate(raw_tasks):
    row = 4 + i
    level = task["reminder_level"]
    fmt = row_formats.get(level, row_formats["On Track"])
    dl_fmt = days_left_formats.get(level, days_left_formats["On Track"])
    p_fmt = priority_formats.get(task["priority"], fmt)

    worksheet.write(row, 0, task["id"], fmt)
    worksheet.write(row, 1, task["category"], fmt)
    worksheet.write(row, 2, task["task"], fmt)
    worksheet.write(row, 3, task["source"], fmt)
    worksheet.write(row, 4, task["start_date"] if task["start_date"] else "-", fmt)
    worksheet.write(row, 5, task["due_date"], fmt)
    worksheet.write(row, 6, task["time_zone"], fmt)
    worksheet.write(row, 7, task["status"], fmt)
    worksheet.write(row, 8, task["priority"], p_fmt)
    worksheet.write(row, 9, task["days_left"] if task["days_left"] is not None else "-", dl_fmt)
    worksheet.write(row, 10, task["reminder_level"], fmt)
    worksheet.write(row, 11, task["notes"], fmt)

    worksheet.set_row(row, 30)

# --- Legend Section ---
legend_row = 4 + len(raw_tasks) + 2
legend_title_fmt = workbook.add_format({"bold": True, "font_size": 13, "font_color": "#1e293b", "bottom": 1})
worksheet.merge_range(legend_row, 0, legend_row, 3, "Color Legend", legend_title_fmt)

legend_items = [
    ("Overdue / Urgent", "#ef4444", "#ffffff"),
    ("Soon (<= 3 Days)", "#f97316", "#ffffff"),
    ("Upcoming (<= 7 Days)", "#eab308", "#ffffff"),
    ("On Track / Ahead", "#22c55e", "#ffffff"),
    ("Completed", "#10b981", "#ffffff"),
]

for j, (label, bg, fg) in enumerate(legend_items):
    r = legend_row + 1 + j
    color_fmt = workbook.add_format({"bg_color": bg, "font_color": fg, "bold": True, "border": 1, "align": "center", "valign": "vcenter"})
    label_fmt = workbook.add_format({"font_size": 11, "valign": "vcenter"})
    worksheet.write(r, 0, "", color_fmt)
    worksheet.merge_range(r, 1, r, 3, label, label_fmt)

# --- Freeze panes & autofilter ---
worksheet.freeze_panes(4, 0)
worksheet.autofilter(3, 0, 3 + len(raw_tasks), len(headers) - 1)

workbook.close()
print(f"Done! Created: Delivery_and_Project_Reminder_Tracker.xlsx")
print(f"Tasks: {len(raw_tasks)} | Date: {today.strftime('%d %B %Y')}")
