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
To use Selenium a special dirver must be downloaded, with the capacity to be controlled with the selenium module and run automatically.
The link to download the driver is the following. 

https://www.selenium.dev/documentation/webdriver/getting_started/install_drivers/

The modules needed are:
* numpy
* pandas
* selenium
* time
