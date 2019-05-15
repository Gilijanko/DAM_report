# DAM_report V1.00
CWE and CEE Electricity day-ahead prices e-mail report

What it does:
The program gathers the following prices:
OKTE (Slovakian DAM)
OPCOM (Romanian DAM)
OTE (Czech DAM)
HUPX (Hungarian DAM)
IBEX (Bulgarian DAM)
CROPEX (Croatian DAM)
BSP (Slovenian DAM)
EPEX DE/FR/AT (German/French/Austrian DAM)
SEEPEX (Serbian DAM)

Sends out an Hourly Chart and Baseload Peakload prices with overall volume

How it works:
The program made out from two python executables.
1. Downloads excels from various websites with Google Chrome
2. Process these excels and sends out an e-mail with Outlook. (I am using Outlook 2013) 

How to setup before first use:
1. Download Chromedriver to C:/webdrivers/chromedriver for Selenium to work with. You can edit the path inside python file.
2. Set default downloadlink: 'C:/Users/**whatever**/Downloads/' This is the place where Chrome downloads files from the browser by default.

Things before use:
1. This program will not work before CET 13:30 everyday. 
This is because the publications on webpages differ and finished around 13:30 CET every day.
2. Sometimes webpages not loading with selenium due to network error. You have to run as many times it stucks on different pages.

Error handling:
If the datahandler (2nd file) not working check the downloaded files. It should contain prices and volumes.
Whenever a webpages changes the program have to change accordingly. I try to keep it up to date.
