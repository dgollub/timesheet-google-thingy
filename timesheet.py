# -*- coding: utf-8 -*-
#
#
from __future__ import print_function

import csv
import os
import re
import sys

import arrow

from gsheets import Sheets


CURRENT_PATH = os.path.abspath(os.path.dirname(os.path.realpath(__file__)))
DEBUG = os.environ.get('DEBUG', "0") == "1"
AS_CSV = os.environ.get('CSV', "0") == "1"

COL_DATE = 0
COL_WEEKDAY = 1
COL_TIME_START = 2
COL_TIME_END = 3
COL_LUNCH = 4
COL_TIME = 5  # includes lunch
COL_TIME_FIXED = 6  # does not include lunch
COL_MOVE = 7
COL_WORK_FROM_HOME = 8
COL_NOTES = 9
COL_TASKS_START = 10
SPECIAL_VALUES = ["sick", "ab", "off", "wfh", "hol"]

SATURDAY = 5
SUNDAY = 6


def calc(hour, half_it=False, split_char = ":"):
    parts = str(hour).split(split_char)
    try:
        local_hours = int(parts[0])        
        local_minutes = int(parts[1])
        if half_it:
            local_hours = local_hours / 2
            local_minutes = local_minutes / 2
        return local_hours, local_minutes
    except:
        if len(parts) == 1:
            try:
                return int(parts[0]), 0
            except:
                return 0, 0

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
    result_rows = [row for row in rows if row and str(row[COL_DATE]) == date]

    if result_rows is None or not result_rows:
        return None
    if len(result_rows) != 1:
        print("More than one entry (%d) found for date %s! Please fix your sheet!" % (len(result_rows), date))
        return None

    found_row = result_rows[0]
    found_index = rows.index(found_row)

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
    notes_str = found_row[COL_NOTES] if max_cols >= COL_NOTES else ""
    notes = notes_str.split('\n')

    # check the previous Friday entry (if today is not Friday), to see what work from home
    # days were were selected
    weekday = (found_row[COL_WEEKDAY] or "").lower()
    check_start_index = found_index if weekday.startswith("fr") else found_index - 7
    check_row = found_row

    while (check_start_index < found_index):
        check_row = rows[check_start_index]
        if (len(check_row) > COL_WEEKDAY and check_row[COL_WEEKDAY] or "").lower().startswith("fr"):
            break
        check_start_index += 1

    is_same_day = None
    if check_start_index != found_index:
    #     print("HA! GOT PREVS FRIDAY.")
        is_same_day = False
    else:
    #     print("SAME DAY")
        is_same_day = True
    
    wfh = u"" if len(check_row)-1 < COL_WORK_FROM_HOME else check_row[COL_WORK_FROM_HOME] 
    wfh = wfh.replace("Mon", "Monday")
    wfh = wfh.replace("Tue", "Tuesday")
    wfh = wfh.replace("Wed", "Wednesday")
    wfh = wfh.replace("Thu", "Thursday")
    wfh = wfh.replace("Fri", "Friday")
    wfh = wfh.replace(", ", ",").replace(",", " and ")
    wfh_extra = "Next week" if is_same_day else "This week"
    wfh_info = """%s %s""" % (wfh_extra, wfh) if wfh != "" else "all days"
    # 2021-01-04 just make this the default for now
    wfh_info = "at all times, unless mentioned otherwise below"


    # regex: ([a-zA-Z].+-\d+)(.*)((?<=\[).+(?=\]))
    # text:  SCAN-4167 As a developer, I want to update AIScanRobo every week [1h]
    # 3 groups:
    #  SCAN-4167
    #  As a developer, I want to update AIScanRobo every week [
    #  1h
    r = re.compile(r"([a-zA-Z].+-\d+)(.*)((?<=\[).+(?=\]))")

    total_time_minutes_from_tasks = 0
    tasks = []
    for idx in range(COL_TASKS_START, max_cols):
        task = found_row[idx].strip()
        if task:
            t = task.split('\n')[0] if '\n' in task else task
            try:
                g = r.match(t).groups()
            except Exception as ex:
                print("ERROR: %s - %s" % (t, str(ex)))
                continue
            if DEBUG:
                print("task:   %s" % (t))
                print("groups: %s" % len(g))

            [task_number, task_details, task_duration] = g
            hours, half_hours = calc(task_duration.replace("h", ""), split_char=".")
            minutes = (hours * 60) + (6 * half_hours)
            total_time_minutes_from_tasks += minutes
            other_lines = task.split('\n')[1:]
            tasks.append("%s %s\n%s" % (task_number.strip(), task_details[:-2].strip(), '\n'.join(other_lines)))

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
                        result += '\n\t' + sub_task
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

    total_hours = str(int(total_time_minutes_from_tasks / 60)).zfill(2)
    total_minutes = str(total_time_minutes_from_tasks % 60).zfill(2)
    total_duration = "%s:%s" % (total_hours, total_minutes)

    test_duration = duration
    if len(test_duration) <= 4:
        test_duration = "0%s" % duration
    if total_duration != test_duration:
        print("")
        print("")
        print("The task times do not add up! Tasks vs time entered: %s != %s" % (total_duration, test_duration))
        print("")
        print("")
# Time: %(start)s - %(end)s (%(duration)s hours total [%(total_hours)s:%(total_minutes)s])
    msg = """
[Daily Report] %(date)s

WFH: %(wfh_info)s


Hi,

Daily Report for Date: %(date)s


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
    "wfh_info": wfh_info,
    "tasks": format_tasks(tasks) if tasks else "",
    "notes": format_notes(notes) if notes else "",
    "total_hours": total_hours,
    "total_minutes": total_minutes,
}

    print("Total time for all tasks (%s): %s - %s:%s" % (len(tasks), total_time_minutes_from_tasks, total_hours, total_minutes))

    return msg


def _load_sheet_data(api, timesheet_url, arg_date=None):
    try:
        date = arrow.get(arg_date, 'YYYYMM')
    except Exception:  # pylint: disable=W0703
        now = arrow.now()
        date = now.format('YYYYMM')

    rows = load_first_sheet_rows(api, timesheet_url, date)
    date_str = str(date.format('YYYYMM'))

    return (rows, date_str)
    

def export_csv(api, timesheet_url, arg_date):
    rows, date = _load_sheet_data(api, timesheet_url, arg_date)
    filtered = [row for row in rows if row and str(row[COL_DATE]).startswith(date)]

    if filtered is None or not filtered:
        return None
    
    csv_filename = os.path.join(os.getcwd(), "%s.csv" % (arg_date))

    print("")
    print("Found (%d) entries for date %s!" % (len(filtered), date))
    print("Writing to %s" % (csv_filename))

    with open(csv_filename, mode='w') as f:
        f = csv.writer(f, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        # f.writerow(['John Smith', 'Accounting', 'November'])
        f.writerow(["username", "date", "task", "duration", "work_type", "details"])

        def w(task, duration_minutes, details = ""):
            work_type = "Meeting" if "meeting" in details.lower() else "Development" 
            # Needed CSV columns
            # username|date|task|duration|work_type|details
            f.writerow(["daniel", arrow.get(str(date), 'YYYYMMDD').format('YYYY.MM.DD'), task, "%dm" % (duration_minutes), work_type, details])

        # text:  SCAN-4167 As a developer, I want to update AIScanRobo every week [1h]
        # 3 groups:
        #  SCAN-4167
        #  As a developer, I want to update AIScanRobo every week [
        #  1h
        r = re.compile(r"([a-zA-Z].*-\d*[^\sa-zA-Z])(.*)((?<=\[).+(?=\]))")
        for row in filtered:
            max_cols = len(row)        
            time = row[COL_TIME_FIXED] if max_cols >= COL_TIME_FIXED else None
            time_start = row[COL_TIME_START] if max_cols >= COL_TIME_START else None
            time_end = row[COL_TIME_END] if max_cols >= COL_TIME_END else None
            date = row[COL_DATE] if max_cols >= COL_DATE else None

            if time_start is None or time_end is None or date is None:
                continue

            tasks = []
            for idx in range(COL_TASKS_START, max_cols):
                task = row[idx].strip()
                if task:
                    tasks.append(task)
            
            if len(tasks) == 0:
                print("%s: no tasks found! %s" % (date, time_start))
                continue

            print("%s: %d tasks found!" % (date, len(tasks)))

            for task in tasks:
                t = task.split('\n')[0] if '\n' in task else task
                try:
                    g = r.match(t).groups()
                except Exception as ex:
                    print("ERROR: %s - %s" % (t, str(ex)))
                    print("%s: failed parsing - %s::::%s::::%s" % (date, task, t, r.match(t)))
                    continue
                if DEBUG:
                    print("task:   %s" % (t))
                    print("groups: %s" % len(g))
                
                [task_number, task_details, duration] = g
                hours, half_hours = calc(duration.replace("h", ""), split_char=".")
                minutes = (hours * 60) + (6 * half_hours)
                if DEBUG:
                    print("time:   %s, %s $ %s $ %s" % (hours, half_hours, duration, minutes))
                details = "%s %s" % (task_number, task_details[:-1].strip())
                w(task_number, minutes, details.strip())

    print("")
    print("CSV output to: %s" % (csv_filename))


def calc_daily_hours_for_month(api, timesheet_url, arg_date):
    rows, date = _load_sheet_data(api, timesheet_url, arg_date)
    filtered = [row for row in rows if row and str(row[COL_DATE]).startswith(date)]

    if filtered is None or not filtered:
        return None
    print("")
    print("Found (%d) entries for date %s!" % (len(filtered), date))

    minutes = 0
    days = 0
    for row in filtered:
        max_cols = len(row)        
        time = row[COL_TIME_FIXED] if max_cols >= COL_TIME_FIXED else None
        time_start = row[COL_TIME_START] if max_cols >= COL_TIME_START else None
        time_end = row[COL_TIME_END] if max_cols >= COL_TIME_END else None
        date = row[COL_DATE] if max_cols >= COL_DATE else None
        worked_at = row[COL_MOVE] if max_cols >= COL_MOVE else None
        notes = row[COL_NOTES] if max_cols >= COL_NOTES else ""

        if time_start is None or time_end is None or date is None:
            continue
        start_hours, start_minutes = calc(time_start)
        end_hours, end_minutes = calc(time_end)

        if start_hours == 0:
            print("%s: Day off because of %s" % (date, "whatever" if time_start == 0 else time_start))
            continue

        extra_info = ""

        the_date = arrow.get(str(date), 'YYYYMMDD')
        if the_date.weekday() in [SATURDAY, SUNDAY]:
            extra_info += " - Weekend work"

        half_day = 'half' in row[COL_WORK_FROM_HOME]
        if half_day:
            extra_info += " - half day PTO"

        if worked_at in ['o', 'O'] or "OFFICE" in notes.upper():
            extra_info += " - Commute to office"

        minutes_day = abs(end_hours - start_hours) * 60
        minutes_day += end_minutes - start_minutes
        minutes += minutes_day

        hours_day = int(minutes_day / 60)
        hours_day_without_lunch = hours_day - 1
        minutes_day = minutes_day % 60
        total_time_for_date = str(hours_day).zfill(2) + ':' + str(minutes_day).zfill(2)

        days += 1

        no_lunch = str(hours_day_without_lunch).zfill(2) + ':' + str(minutes_day).zfill(2)
        print("%s: %s to %s = %s (without lunch: %s)%s" % (date, str(time_start).zfill(2), str(time_end).zfill(2), total_time_for_date, no_lunch, extra_info))

    hours = str(minutes / 60).zfill(2)
    minutes = str(minutes % 60).zfill(2)
    lunch_hours = str(int(float(hours)) - days).zfill(2)
    print("")
    print("Total days worked: %s" % str(days))
    print("Total hours: %s:%s (with 1 hour lunch: %s:%s)" % (hours, minutes, lunch_hours, minutes))
    print("")


def calc_stats(api, timesheet_url, arg_date=None):
    rows, date = _load_sheet_data(api, timesheet_url, arg_date)
    # find the rows for the given month
    filtered = [row for row in rows if row and str(row[COL_DATE]).startswith(date)]

    if filtered is None or not filtered:
        return None
    if not AS_CSV:
        print("")
        print("Found (%d) entries for date %s!" % (len(filtered), date))

    dates, hours = [], []
    half_days = {}
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

        # If it was a half day, meaning I took half a day off, then only count half the time
        half_day = 'half' in row[COL_WORK_FROM_HOME]
        if half_day:
            half_days[date] = time

        hours.append(time)
        dates.append(date)

        if first is None:
            first = row
        else:
            last = row

    total_hours, total_minutes, total_time = 0, 0, ""
    for index, hour in enumerate(hours):
        date = dates[index]
        local_hours, local_minutes = calc(hour, date in half_days)
        total_hours += local_hours
        total_minutes += local_minutes
        if total_minutes >= 60:
            total_hours += (total_minutes / 60)
            total_minutes = total_minutes % 60
        total_time = "%d:%d hours:minutes" % (total_hours, total_minutes)

    expected = 0
    actual_h, actual_m = 0, 0

    if not AS_CSV:
        print("*" * 50)
        print("")
        print("Valid hours entries: %s\t[required vs actual]" % len(hours))

    deduct_work_hours = 0
    work_hours = 0
    work_minutes = 0
    days = 0
    expected_hours_accumulated_total = 0
    for index, worked_date in enumerate(dates):
        days += 1
        if hours[index] in SPECIAL_VALUES:
            if not AS_CSV:
                print("  %s: Off, because %s" % (worked_date, hours[index]))
            else:
                pass
        else:
            half_day = worked_date in half_days
            # each workday has 8 hours of work, but on half days it is only half of 8, aka 4.
            work_hours_for_the_day = 8 if not half_day else 4
            expected_hours_accumulated_total += 8 - (8 - work_hours_for_the_day)
            expected_minutes_accumulated_total = expected_hours_accumulated_total * 60

            # hours[index] is the actual time worked, e.g. 6:30 means 6 hours and 30 minutes
            local_h, local_m = calc(hours[index])

            work_hours += local_h
            work_minutes += local_m

            actual_h = work_hours
            # 330 minutes = 6 hours and 30 minutes
            actual_h += int(work_minutes / 60)
            actual_m = work_minutes % 60
            if AS_CSV:
                print("%s;%s;" % (worked_date, hours[index]))
            else:
                print("  %s: %s\t[%s:00 vs %s:%s] %s" % (worked_date, hours[index], expected_hours_accumulated_total,
                                                    str(actual_h).zfill(2), str(actual_m).zfill(2),
                                                    "Half day" if half_day else ""))

    if not AS_CSV:
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

    print("")
    print("Usage:   python timesheet.py [command|date] [date]")
    print("Example: python timesheet.py stats 202011")
    print("Example: python timesheet.py 20201130")
    print("")
    print("Available commands:")
    print("- stats: show summed up hours and minutes for the given/current month")
    print("         use \"CSV=1 python timesheet.py stats\" to format the output")
    print("         as CSV")
    print("- daily: same as stats, except ready to email to HR")
    print("- csv: task breakdown for the month and time spend on each task")
    print("")
    print("""Tip: use "DEBUG=1 timesheet <parameter>" to enable debug output""")
    print("")

    print("Trying to load client-secrets.json file ...")
    secrets_file, cache_file = get_client_secret_filenames()
    sheets = Sheets.from_files(secrets_file, cache_file, no_webserver=False)
    print("Success.")

    date = None if len(sys.argv) < 3 else sys.argv[2].strip()
    arg = "read today" if len(sys.argv) < 2 else sys.argv[1].strip()

    if arg == "stats":
        calc_stats(sheets, timesheet_url, date or arrow.now().format('YYYYMM'))
    elif arg == "daily":
        calc_daily_hours_for_month(sheets, timesheet_url, date or arrow.now().format('YYYYMM'))
    elif arg == "csv":
        export_csv(sheets, timesheet_url, date or arrow.now().format('YYYYMM'))
    else:
        date_to_use = "read today" if arg == '' else arg
        load_sheet_and_read_data(sheets, timesheet_url, date_to_use, user_full_name)

    print("Done.")


if __name__ == "__main__":
    main()
