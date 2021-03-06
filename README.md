[
![PyPI](https://img.shields.io/pypi/v/sp_ask_report_chats_per_school.svg)
![PyPI](https://img.shields.io/pypi/pyversions/sp_ask_report_chats_per_school.svg)
![PyPI](https://img.shields.io/github/license/guinslym/sp_ask_report_chats_per_school.svg)
](https://pypi.org/project/sp_ask_report_chats_per_school/)
[![TravisCI](https://travis-ci.org/guinslym/sp_ask_report_chats_per_school.svg?branch=master)](https://travis-ci.org/guinslym/sp_ask_report_chats_per_school)

# SP Ask Running Total hourly report
An helper that create and Excel document representing the number of chats per day. Which are on the hourly Spreadsheet

### Installation

    $ pip install sp_ask_running_total_hourly_report
    $ or with poetry
    $ poetry add sp_ask_running_total_hourly_report

## Requirement

lh3api is a requirement for this package[ lh3api configuation files](https://gitlab.com/libraryh3lp/libraryh3lp-sdk-python/) need to be installed first.


In ~/.lh3/config::

    [default]
    server = libraryh3lp.com
    timezone = UTC
    salt = "you should probably change this"

The `salt` is used when generating system-level utility accounts.
This is not something you do often.  If your `salt` is unique, your
passwords will be unique.

In ~/.lh3/credentials::

    [default]
    username = <ADMIN_USER>
    password = <ADMIN_PASS>

    [test]
    username = <TEST_USER>
    password = <TEST_PASS>


### Utilisation
```python
from sp_ask_running_total_hourly_report import create_report

# generating stats for a given month (february)
create_report(2019, 2)

#generating stats for all months
for month_number in range(1, 13):
    create_report(2019, month_number)
```
Will create a file **'2-February.xlsx'** on the current directory

### Screenshot
This is a screenshot. We don't provide our hourly stats to the public. The number below were updated to fake our real numbers.
![screenshot of app](screenshots/screenshot3.png)


### Todo
1.  Test
2.  Refactor
3.  Add docstring
4.  Add a Makefile
5.  Add show terminal graph for a given month
6.  Create a .rst file and add metadata for 
