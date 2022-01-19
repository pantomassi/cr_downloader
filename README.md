# cr_downloader

# Goal
Automatization of report submission on a webpage connected to Oracle R12 database. One Job Number = one request. No API available, so used selenium to submit the requests. Not the most convenient way to do this... but it works.

# Considering three scenarios of data provided by the user:
1) Only single, non-related Job Numbers in each cell of column A
2) Only ranges of consecutive Job Numbers (e.g. No.11-No.21, then No.30-No.45) in rows of columns B/C
3) Mix of 1 & 2

![excel_data_for_submission](https://github.com/pantomassi/cr_downloader/blob/master/projects_list_example.PNG)

# How to use it
Run the script, choose excel list with Job Numbers to submit, wait for selenium to submit the data.
![choose_data](https://github.com/pantomassi/cr_downloader/blob/master/choose_jobs.png)

# Why it's tricky
* Need to switch between windows, as parent page opens multiple child windows.
* All pages load slowly, in most cases used selenium methods of waiting for DOM elements. Couldn't figure it out in two places where NO wait method worked, so simple timeout had to be applied. The waiting time used is a compromise between slow loading and being sure that everything is filled out correctly, even with a very poor connection/slow machine.
* The first Job submission is different than the ones that follow- it forces a child window pop-up, so couldn't use the same fill-out methods as for following Job numbers.
* The input fields (two of them) where Job Number is to be submitted have an unexpected reload. After one is filled out, the whole frame (or sth else?) is reloaded, so the other field apparently disappears from DOM - simple timeout had to be used.


# What can be improved:
* Code repetition - part with gathering data from excel
* Add error logger
* Figure out how to get rid of timeouts and replace them with more universal waiting methods
* Take into account that chromedriver.exe used for the development will become obsolete/incompatible against new Chrome versions. Could possibly detect what's the current version of browser & driver, then download new version of the latter if it proves to be obsolete. If not, it's up to the user to manually download a new chromedriver
