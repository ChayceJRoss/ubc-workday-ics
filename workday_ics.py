import sys
import pandas as pd
from datetime import datetime

def get_events(df):
    events = []
    today = datetime.now()
    ical_format = "%Y%m%dT%H%M%S"
    for index, row in df.iterrows():
        print(row)
        course = row["Course Listing"]
        dates = row["Meeting Patterns"]
        if type(dates) == float:
            continue
        dates = str(dates)
        date_lines = dates.split("\n")
        date_separators = map(lambda section : section.split("|"), filter(lambda line : len(line) > 0, date_lines))
        for section in date_separators:
            start_end = section[0].strip().split(" ")
            start_date = start_end[0]
            end_date = start_end[2]
            end_date = pd.to_datetime(end_date, format="%Y-%m-%d")
            days = section[1].strip().split(" ")
            times = "".join(filter(lambda c: c != '.', section[2])).strip().split(" - ")
            start_date_beg = start_date + " " + times[0]
            start_date_end = start_date + " " + times[1]
            datetimes_start = pd.to_datetime(start_date_beg, format="%Y-%m-%d %I:%M %p")
            datetimes_end = pd.to_datetime(start_date_end, format="%Y-%m-%d %I:%M %p")
            valid_days = ["MO", "TU", "WE", "TH", "FR", "SA", "SU"]
            days_formatted = ",".join(filter(lambda da: da in valid_days, map(lambda day : day.upper()[0:2], days)))
            event = f"""BEGIN:VEVENT
DTSTAMP:{today.strftime("%Y%m%dT%H%M%SZ")} 
DTSTART;TZID=America/Vancouver:{datetimes_start.strftime(ical_format)}
DTEND;TZID=America/Vancouver:{datetimes_end.strftime(ical_format)}
RRULE:FREQ=WEEKLY;BYDAY={days_formatted};UNTIL={end_date.strftime("%Y%m%dT%H%M%SZ")}
SUMMARY:{course}
END:VEVENT"""
            events.append(event)
    return events


def get_ics(events):
    today = datetime.now()
    events_string = '\n'.join(events)
    today_edited = today.strftime("%Y%m%dT%H%M%SZ")
    return f"""BEGIN:VCALENDAR
PRODID:-//Microsoft Corporation//Outlook 19.0 MIMEDIR//EN
VERSION:2.0
X-WR-TIMEZONE:America/Vancouver
METHOD:PUBLISH
BEGIN:VTIMEZONE
TZID:America/Vancouver
X-LIC-LOCATION:America/Vancouver
BEGIN:DAYLIGHT
TZOFFSETFROM:-0800
TZOFFSETTO:-0700
TZNAME:PDT
DTSTART:19700308T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:-0700
TZOFFSETTO:-0800
TZNAME:PST
DTSTART:19701101T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
END:STANDARD
END:VTIMEZONE
{events_string}
END:VCALENDAR"""

def main():
    source = sys.argv[1]
    destination = sys.argv[2]

    df = pd.read_excel(source, dtype=str)
    print(df)
    events = get_events(df)
    print(events)
    ics_string = get_ics(events)
    with open(destination, "w", encoding="utf-8") as f: 
        f.write(ics_string)

if __name__ == "__main__":
    main()
