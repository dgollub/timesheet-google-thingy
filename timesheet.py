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
SPECIAL_VALUES = ["sick", "ab", "off", "wfh"]


def get_client_secret_filenames():
    filename = os.path.join(CURRENT_PATH, "client-secrets.json")
    cachefile = os.path.join(CURRENT_PATH, "client-secrets-cache.json")

    if not os.path.exists(filename):
        filename = os.path.expanduser(os.path.join("~", "client-secrets.json"))
        cachefile = os.path.expanduser(os.path.join("~", "client-secrets-cache.json"))
    if not os.path.exists(filename):
        raise Exception("Please provide a client-secret.json file, as described here: https://github.com/xflr6/gsheets#quickstart")

    return filename, cachefile


def load_sheet_and_read_data(api, timesheet_url, commandline, user_full_name):
    now = arrow.now()
    today = now.format('YYYYMMDD')

    print("Opening timesheet for %s (%s)..." % (today, commandline))

    sheets = api.get(timesheet_url)
    sheet = sheets.sheets[0]

    print(u"Timesheet [%s] sheet [%s] opened. Accessing cell data ..." % (sheets.title or "???", sheet.title or "???"))

    rows = sheet.values()

    # TODO(dkg): implement proper commandline handling for more than just this default
    timesheet = get_timesheet_for_date(rows, today, user_full_name)
    if timesheet:
        print("\n\n")
        print("Timesheet for %s" % (today))
        print(timesheet)
        print("\n")
    else:
        print("No entry found for %s" % today)


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
SUBJECT: [Daily Report] %(date)s
BODY:
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

    commandline = "read today"
    load_sheet_and_read_data(sheets, timesheet_url, commandline, user_full_name)

    print("Done.")

if __name__ == "__main__":
    main()
