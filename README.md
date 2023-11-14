
# Vacation list generator
This code generates Excel based vacation list for employees at an organisation (team, department, company) to enter their vacations, day off, leave of absence, etc.

## Limitation

The Sweden holiday calendar is used by default. There is currently no support to change to other countries but please make a feature request if needed.

The main challenge is that we need to have a reliable source of holidays for different countries. There is universal Python package or API that can provide trusted and verified list of holidays for a country.

Even in the limitation to Sweden, this tool was desgined to cache 10 years of Swedish holidays (2023-2033) due to the uncertainty if the used API is up and free at any given time. Of course, cached holidays will come in handy not only for this tool but some others. And there is no intension that the tool will survive that long.

## How to use vacation list
By default, a vacation list has 3 tabs in which a tab is correponding to 4 months period.
(This can be configured in later version)

`David`, an imaginary employee working at `Backend team` starts at `Names+Teams - Enter here` tab to enter his name and team. This step can be done by manager at begining of a year.

![Enter names and teams here](images/Enter-names-teams.png)

There are certain colors used in vacation list. For examples:
* <span style="background-color: #ccc;color: black">Sat</span> and <span style="background-color: #ccc;color: black">Sun</span> as weekend.
* <span style="background-color: #FF0000;">Wed</span> indicates that Wednesday is a public holiday.
* Half working day (not public holiday) is marked as pink
* <span style="background-color: #00ff00;">V</span> or <span style="background-color: #00ff00;">v</span> indicate that this is a vacation day.
* <span style="background-color: #006400;">A</span> or <span style="background-color: #006400;">a</span> indicate that this is an approved vacation day by manager.
* <span style="background-color: #0000ff;">E</span> or <span style="background-color: #0000ff;">e</span> indicate that an employee takes off for education.
* <span style="background-color: #800080;">P</span> or <span style="background-color: #800080;">p</span> indicate that an employee takes off for parental leave.

* Other reason (<span style="background-color: #ffff00;">O</span> or <span style="background-color: #ffff00;">o</span> ) is marked as **yellow** 

![Alt text](images/Undersand-the-colors.png)

Now, `David` can enter his day off through out the year.

![Alt text](images/David-enters-his-day-off.png)

# How to generate vacation list generator
## Installation
````
$ python3 -m venv env
$ source env/bin/activate
$ pip install -r requirements.txt

````
## Run

Generate vacation list for 2025

````
$ python main 2025

````
The `2025-vacationlist.xlsx` will be outputed under `output` folder.

To lock all uninvited areas by a password, for example `Yhlm=1`. It is optional to specify password as the default password is `12345`

````
$ python main 2025 --password Yhlm=1

````

By default the vacation list is created with 3 sheets:
January-April, May-August, and September-Decemnber
To change this behaviour, please specify `--periods` parameter. 

For example to create 3 sheets: January-April, May - August and July-December

````
$ python main 2025 --password Yhlm=1 --periods 1-4, 5-8, 9-12

````

And to create only 2 sheets (January-June and July-December)

````
$ python main 2025 --password Yhlm=1 --periods 1-6, 7-12

````

And to create 12 sheets for 12 months

````
$ python main 2025 --password Yhlm=1 --periods 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 

````