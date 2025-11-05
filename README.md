# UBC Workday Excel to ICS CLI Tool

This Python tool converts your Workday class schedule into a .ics file you can use in Apple Calendar, Proton Calendar, or any other iCalendar compatible online calendaring service.

## Installation

Clone this repo:

```
git clone https://github.com/terraceonhigh/ubc-workday-ics.git
```

To install run:

```
python3.XX -m venv venv
```

If on Windows run:

```
.\venv\Scripts\activate
```

If on Linux or Mac run:

```
source venv/bin/activate
```

Finally run

```
pip install -r requirements.txt
```
### Troubleshooting

Your system might throw out an error saying that Numpy failed to install:

>       The Meson build system
      Version: 1.2.99
      Source dir: /tmp/pip-install-zf4z_60j/numpy_f2c88ea52d214bcd9f45d07e251d0aa6
      Build dir: /tmp/pip-install-zf4z_60j/numpy_f2c88ea52d214bcd9f45d07e251d0aa6/.mesonpy-8b8es6f9
      Build type: native build
      Project name: NumPy
      Project version: 2.0.0
      C compiler for the host machine: cc (gcc 15.2.1 "cc (GCC) 15.2.1 20250808 (Red Hat 15.2.1-1)")
      C linker for the host machine: cc ld.bfd 2.44-6
      
      ../meson.build:1:0: ERROR: Unknown compiler(s): [['c++'], ['g++'], ['clang++'], ['nvc++'], ['pgc++'], ['icpc'], ['icpx']]
      The following exception(s) were encountered:
      Running `c++ --version` gave "[Errno 2] No such file or directory: 'c++'"
      Running `g++ --version` gave "[Errno 2] No such file or directory: 'g++'"
      Running `clang++ --version` gave "[Errno 2] No such file or directory: 'clang++'"
      Running `nvc++ --version` gave "[Errno 2] No such file or directory: 'nvc++'"
      Running `pgc++ --version` gave "[Errno 2] No such file or directory: 'pgc++'"
      Running `icpc --version` gave "[Errno 2] No such file or directory: 'icpc'"
      Running `icpx --version` gave "[Errno 2] No such file or directory: 'icpx'"


This Helps:

https://github.com/grpc/grpc/issues/24556#issuecomment-727745067

## Example

This script takes two filenames, the first is the workday excel file and the
second will be name of the .ics file (should end in .ics)

e.g.

1. Navigate to [UBC Workday](myworkday.ubc.ca) and login.
2. Go to the "Academics" section then click the "Registration" tab.
3. Finally click the settings wheel in the "Current Schedule" block and download
   the excel file.
4. Copy and paste it into the cloned directory.
5. Run:

```
python workday_ics.py Current_Schedule.xlsx calendar.ics --reminder 30,10
```

6. I woud reccomend making a new calendar in your calendar software incase this
   does not work. Then upload the ics file to there.

## Options

```
--help
--reminder (comma seperated, minute values for setting reminders)
```
