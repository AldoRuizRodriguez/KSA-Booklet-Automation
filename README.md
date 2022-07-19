# KSA-Booklet-Automation

## Description

This Project is used to make automatic all the process to build the reports of KSA

## Steps of the automation

### Download of the reports from the IBIS 
Hence the this platform generates each report individually, the module of selenium is used to interact with the web-site through Python coding.
Firstly, the script login to the page using your credentials, then selecting each report to be generated, and finally making the download in a .xls(excel) file.
### Formatting of each report
The Downloaded reports are modified and cleansed with pandas module.

## Requisites
To use Selenium a special dirver must be downloaded, with the capacity to control the web-site with the selenium module and run automatically.
The link to download the driver is the following. 

https://www.selenium.dev/documentation/webdriver/getting_started/install_drivers/

The modules needed are:
* numpy
* pandas
* selenium
* time

## Files

* Weekly Reports: This Script made in python the Downloads and give the correct format to the reports
* Monthly Reports: This Script made in python the Downloads of each report(to the present day the formatting to this reports is in progress)
* List of Reports: This file shows the Status of each report, giving information if the one Specific report has passed correctly the process to be downloaded, if has been any issues with the Automatic Download or presents some problems at the time of giving format.
* Unpivoting Monthly Reports: This script is used to unpivot some of the Monthly Reports, since when uploading to Google Data Studio,the data have to be with a Standarized format, and not as it  downloaded from IBIS server

