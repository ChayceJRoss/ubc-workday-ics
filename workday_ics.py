import sys
import pandas as pd
from datetime import datetime
import click

class IcsRRule:
    def __init__(self, freq: str, byday: str, until: str):
        self.freq = freq
        self.byday = byday
        self.until = until
    def ToString(self)->str:
        return f"FREQ={self.freq};BYDAY={self.byday};UNTIL={self.until}"

class IcsVAlarm:
    def __init__(self, action: str, trigger: str, descrption: str):
        self.action = action
        self.trigger = trigger
        self.description = descrption
    def ToString(self)->str:
        return f"""BEGIN:VALARM
ACTION:{self.action}
TRIGGER;VALUE=DURATION:-PT{self.trigger}M  
DESCRIPTION:{self.description}
END:VALARM
"""

class IcsEvent:
    def __init__(self, dtstamp: str, dtstart: str, dtend: str, rrule: IcsRRule, summary: str, location: str, valarms: list[IcsVAlarm] = []):
        self.dtstamp  = dtstamp
        self.dtstart  = dtstart
        self.dtend = dtend
        self.rrule = rrule
        self.summary = summary
        self.location = location
        self.valarms = valarms
    def ToString(self)->str:
        valarms = ""
        for valarm in self.valarms:
            valarms += (valarm.ToString())
        return f"""BEGIN:VEVENT
DTSTAMP:{self.dtstamp}
DTSTART;TZID=America/Vancouver:{self.dtstart}
DTEND;TZID=America/Vancouver:{self.dtend}
RRULE:{self.rrule.ToString()}
SUMMARY:{self.summary}
LOCATION:{self.location}
{valarms}
END:VEVENT
"""


# TODO Make look nice, add instructions for downloading the right calendar file and add author tag.
def get_events(df: pd.DataFrame) -> list[IcsEvent]:
    events = []
    today = datetime.now()
    ical_format = "%Y%m%dT%H%M%S"
    for _, row in df.iterrows():
        course = row["Course Listing"]
        meeting_patterns = row["Meeting Patterns"]
        if type(meeting_patterns) == float:
            continue
        meeting_patterns = str(meeting_patterns)
        date_lines = meeting_patterns.split("\n")
        date_separators: map[list[str]] = map(lambda section : section.split("|"), filter(lambda line : len(line) > 0, date_lines))
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
            location = section[3]
            event = IcsEvent(today.strftime("%Y%m%dT%H%M%SZ"), 
                             datetimes_start.strftime(ical_format), 
                             datetimes_end.strftime(ical_format),
                             IcsRRule("WEEKLY", days_formatted, end_date.strftime("%Y%m%dT%H%M%SZ")),
                             course, 
                             location) 
            events.append(event)
    return events

def get_alarms(events: list[IcsEvent], reminder: str):
    times = reminder.split(",")
    if reminder == "":
        return
    alarms = []
    for time in times:
        alarm = IcsVAlarm("DISPLAY",time, f"T-minus {time} minutes")
        alarms.append(alarm)
    for event in events:
        event.valarms = alarms


def get_ics(events: list[IcsEvent], author: str):
    today = datetime.now()
    events_string = ''
    for event in events:
        events_string += event.ToString()
    today_edited = today.strftime("%Y%m%dT%H%M%SZ")
    return f"""BEGIN:VCALENDAR
PRODID:{author}
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

@click.command()
@click.argument("source", type=click.Path(exists=True) )
@click.argument("destination", type=click.File("wb")) 
@click.option("--author", default="DEFAULT", help="Name of author for ical file.")
@click.option("--reminder", default="", help="Set a reminder the specified minutes before event, seperate values by comma")
def main(source, destination, author, reminder):
    print(source)
    df = pd.read_excel(source, dtype=str)
    events = get_events(df)
    get_alarms(events, reminder)
    ics_string = get_ics(events, author)
    click.echo(ics_string)
    destination.write(ics_string.encode(encoding="utf-8"))

if __name__ == "__main__":
    main()
