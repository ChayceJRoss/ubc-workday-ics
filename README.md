# UBC Workday Excel to ICS CLI Tool

## Installation

To install run:

```bash
python3.XX -m venv venv
```

If on Windows run:

```powershell
.\venv\activate
```

If on Linux run:

```bash
source venv/bin/activate
```

Finally run

```bash
pip install -r requirements.txt
```

## Example

This script takes two filenames, the first is the workday excel file and the
second will be name of the .ics file (should end in .ics)

e.g.

```bash
python workday_ics.py Current_Schedule.xlsx calendar.ics
```

## Help

Use:

```bash
python workday_ics.py --help
```
