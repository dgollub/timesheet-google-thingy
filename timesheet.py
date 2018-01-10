# -*- coding: utf-8 -*-
#
#
from __future__ import print_function

import os
import sys

import arrow

from gsheets import Sheets

reload(sys)
sys.setdefaultencoding('utf-8')


CURRENT_PATH = os.path.abspath(os.path.dirname(os.path.realpath(__file__)))

COL_DATE = 0
COL_WEEKDAY = 1
COL_TIME_START = 2
COL_TIME_END = 3
COL_LUNCH = 4
COL_TIME = 5  # includes lunch
COL_TIME_FIXED = 6  # does not include lunch
COL_MOVE = 7
COL_NOTES = 8
COL_TASKS_START = 9
SPECIAL_VALUES = ["sick", "ab", "off", "wfh", "hol"]


def get_client_secret_filenames():
    filename = os.path.join(CURRENT_PATH, "client-secrets.json")
    cachefile = os.path.join(CURRENT_PATH, "client-secrets-cache.json")

    if not os.path.exists(filename):
        filename = os.path.expanduser(os.path.join("~", "client-secrets.json"))
        cachefile = os.path.expanduser(os.path.join("~", "client-secrets-cache.json"))
    if not os.path.exists(filename):
        raise Exception("Please provide a client-secret.json file, as described here: https://github.com/xflr6/gsheets#quickstart")

    return filename, cachefile


def load_first_sheet_rows(api, timesheet_url, date=arrow.now().format('YYYYMMDD')):

    print("Opening timesheet for %s ..." % (date))

    sheets = api.get(timesheet_url)
    sheet = sheets.sheets[0]

    print(u"Timesheet [%s] sheet [%s] opened. Accessing cell data ..." % (sheets.title or "???", sheet.title or "???"))

    rows = sheet.values()

    return rows


def load_sheet_and_read_data(api, timesheet_url, commandline, user_full_name):
    now = arrow.now()
    today = now.format('YYYYMMDD')

    try:
        other_date = arrow.get(commandline, 'YYYYMMDD').format('YYYYMMDD')
    except arrow.parser.ParserError:
        other_date = today

    use_date = other_date

    rows = load_first_sheet_rows(api, timesheet_url, use_date)

    timesheet = get_timesheet_for_date(rows, use_date, user_full_name)
    if timesheet:
        print("\n\n")
        print("Timesheet for %s" % (use_date))
        print(timesheet)
        print("\n")
    else:
        print("No entry found for %s" % use_date)


def get_timesheet_for_date(rows, date, user_full_name):

    # find the row with the first column that has today's date in it
    rows = [row for row in rows if row and str(row[COL_DATE]) == date]

    if rows is None or not rows:
        return None
    if len(rows) != 1:
        print("More than one entry (%d) found for date %s! Please fix your sheet!" % (len(rows), date))
        return None

    found_row = rows[0]

    start_val = found_row[COL_TIME_START]
    end_val = found_row[COL_TIME_END]
    duration_val = found_row[COL_TIME_FIXED]
    max_cols = len(found_row)

    if not start_val:
        if start_val in SPECIAL_VALUES:
            print("You forgot to add your start time.")
        return None
    if not end_val:
        if end_val in SPECIAL_VALUES:
            print("You forgot to add your end time.")
        return None
    #if max_cols >= COL_NOTES:
    #    print("No notes/tasks entered yet.")
    #    return None

    def parse_hours(val):
        try:
            return arrow.get(val, "HH:mm")
        except arrow.parser.ParserError:
            return arrow.get(val, "H:mm")

    start = parse_hours(start_val).format("HH:mm")
    end = parse_hours(end_val).format("HH:mm")
    duration = str(duration_val)
    notes_str = found_row[COL_NOTES]
    notes = notes_str.split('\n')

    tasks = []
    for idx in range(COL_TASKS_START, max_cols):
        task = found_row[idx].strip()
        if task:
            tasks.append(task)

    def format_tasks(tasks):
        if not tasks:
            return ''

        result = 'Tasks:\n'

        for task in tasks:
            if '\n' in task:
                sub_tasks = task.split('\n')
                if len(sub_tasks) > 1:
                    result += '\n* ' + sub_tasks[0]  # main task
                    for sub_task in sub_tasks[1:]:  # actual sub tasks
                        result += '\n\t- ' + sub_task
                    result += '\n'
                else:
                    result += '\n* ' + task
            else:
                result += '\n* ' + task

        return result

    def format_notes(notes):
        if not notes or (len(notes) == 1 and not notes[0]):
            return ''

        result = 'Additional Notes:\n'

        for note in notes:
            result += '\n* ' + note

        return result

    msg = """
[Daily Report] %(date)s

Hi,

Daily Report for Date: %(date)s

Start Time: %(start)s
End Time: %(end)s

%(tasks)s

%(notes)s

Kind regards,
%(user_full_name)s
""".strip() % {
    "date": date,
    "user_full_name": user_full_name,
    "start": start,
    "end": end,
    "duration": duration,
    "tasks": format_tasks(tasks) if tasks else "",
    "notes": format_notes(notes) if notes else "",
}

    return msg


def calc_stats(api, timesheet_url, arg_date=None):
    try:
        date = arrow.get(arg_date, 'YYYYMM')
    except Exception:  # pylint: disable=W0703
        now = arrow.now()
        date = now.format('YYYYMM')

    rows = load_first_sheet_rows(api, timesheet_url, date)
    date_str = str(date.format('YYYYMM'))
    # find the rows for the given month
    filtered = [row for row in rows if row and str(row[COL_DATE]).startswith(date_str)]

    if filtered is None or not filtered:
        return None
    print("")
    print("Found (%d) entries for date %s!" % (len(filtered), date))

    dates, hours = [], []
    first = None
    last = None
    for row in filtered:
        max_cols = len(row)        
        time = row[COL_TIME_FIXED] if max_cols >= COL_TIME_FIXED else None
        tasks = []
        for idx in range(COL_TASKS_START, max_cols):
            task = row[idx].strip()
            if task:
                tasks.append(task)

        day_type = row[COL_TIME_START] if max_cols >= COL_TIME_START else None
        date = row[COL_DATE] if max_cols >= COL_DATE else None

        if day_type is None:
            continue

        if day_type in SPECIAL_VALUES:
            time = day_type
            hours.append(time)
            dates.append(date)
            continue
        elif not tasks:
            continue

        hours.append(time)
        dates.append(date)

        if first is None:
            first = row
        else:
            last = row


    def calc(hour):
        parts = str(hour).split(":")
        try:
            local_hours = int(parts[0])
            local_minutes = int(parts[1])
            return local_hours, local_minutes
        except:
            return 0, 0

    total_hours, total_minutes, total_time = 0, 0, ""
    for hour in hours:
        local_hours, local_minutes = calc(hour)
        total_hours += local_hours
        total_minutes += local_minutes
        if total_minutes >= 60:
            total_hours += (total_minutes / 60)
            total_minutes = total_minutes % 60
        total_time = "%d:%d hours:minutes" % (total_hours, total_minutes)

    expected = 0
    actual_h, actual_m = 0, 0

    print("*" * 50)
    print("")
    print("Valid hours entries: %s\t[required vs actual]" % len(hours))
    special = 0
    for index, worked_date in enumerate(dates):
        if hours[index] in SPECIAL_VALUES:
            print("  %s: Off, because %s" % (worked_date, hours[index]))
            special = special + 1
        else:
            expected = str((index + 1 - special) * 8).zfill(2)
            local_h, local_m = calc(hours[index])
            actual_m += local_m
            actual_h += local_h + (actual_m / 60)
            actual_m = actual_m % 60
            print("  %s: %s\t[%s:00 vs %s:%s]" % (worked_date, hours[index], expected,
                                                  str(actual_h).zfill(2), str(actual_m).zfill(2)))
    print("")
    print("First:", "<first> not found" if first is None else first[COL_DATE])
    print("Last:", "<last> not found" if last is None else last[COL_DATE])
    print("")
    print("Total time in %s: %s" % (date, total_time))
    print("")
    print("*" * 50)


def main():
    # print("Checking environment variable TIMESHEET_URL for spreadsheet URL...")
    timesheet_url = os.environ.get('TIMESHEET_URL', "").strip()
    if not timesheet_url:
        raise Exception("Please set the TIMESHEET_URL environment variable accordingly.")
    # print("Checking environment variable USER_FULL_NAME for spreadsheet URL...")
    user_full_name = os.environ.get('USER_FULL_NAME', "").strip()
    if not user_full_name:
        print("Warning: USER_FULL_NAME environment variable not set!")
        user_full_name = "Herman Toothrot"

    print("Trying to load client-secrets.json file ...")
    secrets_file, cache_file = get_client_secret_filenames()
    sheets = Sheets.from_files(secrets_file, cache_file)
    print("Success.")

    arg = "read today" if len(sys.argv) < 2 else sys.argv[1].strip()
    if arg == "stats":
        date = None if len(sys.argv) < 3 else sys.argv[2].strip()
        calc_stats(sheets, timesheet_url, date)
    else:
        date_to_use = "read today" if arg == '' else arg
        load_sheet_and_read_data(sheets, timesheet_url, date_to_use, user_full_name)

    print("Done.")


if __name__ == "__main__":
    main()
