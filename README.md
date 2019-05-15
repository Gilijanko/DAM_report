# DAM_report V1.00
CWE and CEE Electricity day-ahead prices e-mail report

What it does: The program gathers the following prices: OKTE (Slovakian DAM) OPCOM (Romanian DAM) OTE (Czech DAM) HUPX (Hungarian DAM) IBEX (Bulgarian DAM) CROPEX (Croatian DAM) BSP (Slovenian DAM) EPEX DE/FR/AT (German/French/Austrian DAM) SEEPEX (Serbian DAM)

Sends out an Hourly Chart and Baseload Peakload prices with overall volume

How it works: The program made out of two python executables.

Downloads excel from various websites with Google Chrome
Process these excels and sends out an e-mail with Outlook. (I am using Outlook 2013)
How to set up before first use:

Download Chromedriver to C:/webdrivers/chromedriver for Selenium to work with. You can edit the path inside python file.
Set default download link: 'C:/Users/whatever/Downloads/' This is the place where Chrome downloads files from the browser by default.
Things before use:

This program will not work before CET 13:30 every day. This is because the publications on webpages differ and finished around 13:30 CET every day.
Sometimes webpages not loading with selenium due to a network error. You have to run as many times it stucks on different pages.
Error handling: If the data handler (2nd file) not working check the downloaded files. It should contain prices and volumes. Whenever a webpage changes the program has to change accordingly. I try to keep it up to date.
