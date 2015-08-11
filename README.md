# Pick 3 Crawler
    This program was built to crawl through a series of data files, copy each one to a specific directory, run a report generation program through automation, and find matching patterns in the generated reports. 
    
    Please note that this is a very purpose-built program designed to work with uniquely formatted data file and a pre-existing report generation program.
    
## Setup
1. Install Python 2.7.10 (32 bit).
        
    This program was written with the 32 bit version of Python 2.7.10 as opposed to the 64 bit version because of the automation used to run Report3.exe (which itself is a 32 bit program). While it *may* work just fine with the 64 bit version of Python, pywinauto will throw warnings.

2. Install [pyWin32 extentions](http://sourceforge.net/projects/pywin32/files/pywin32/). (Build 219 > pywin32-219.win32-py2.7.exe)
3. Download the [latest version of pywinauto](https://github.com/pywinauto/pywinauto/releases/download/0.5.1/pywinauto-0.5.1.zip)
4. Unpack and run `python setup.py install`
5. Run `python main.py`

## Credits
Author: Nicholas Mastracchio

Date: 8/10/2015