# This is the Python Power Meter Program


## v1.0 
- Works with a nice gui and gives good results. For the next version add better excel with perhaps prices worked out. This version only works with csv files, in the future see what can be done with excel

## v2.1 
- This is where we added the field to make the tariffs available, they are not yet linked to anything however

## v2.2 
- This is where we were able to implement a button that is used to get the information of the field

## v2.3 
- This is the version where the fields now have the default set to 2.0 and the function that tests them works sort of. It adds the values for now. 

## v2.4 
- In this version, we created the excel spreadsheet in a rough format. FOr future versions:
- add a three extra tariffs in gui: Max demand, standard?, winter?
- add the full package to the excel output file
- add a loading bar when generating results

## v2.5.1 
- Features added since v2.4:
- Gui moved around a bit and added fields for max demand, ECB and NEF levies.
- This is reflected in the excel output 

## v2.5.2 
Features added:
- Added a network charge field for the gui8
- Added the functionality that sums all the sub-totals. 
- The Spreadsheet still does not do current and previous as that will be alot of effort
- Missing nice formatting also 

## v2.6
- This version creates a very complete excel spreadsheet.
- The only thing missing is the month and previous formulas. This will probably be entered manually unless specifically requested.

## v3.1
- This version includes the addition of the Declared Demand field as well as the subsequent excel results.
- In addition, the including Vat field has been added to the excel document.
- Next Version plans: Add the net metering fields as well as their appropriate tarriffs

## v3.2
- This version is complete with the Net metering functionality

## Notes as of 23-03-2025
External libraries are required for this program to work. To install them type the following in the terminal/powershell:

1. pip install numpy
2. pip install pandas
3. pip install openpyxl 

Additionally this program only works for certain comma delimited data of the following format: 

*"","18-Jun-19","14:00","14:30","19.263000","0.000000","0.000000","19.757000","Phase Failure"*

not this format:

*06/07/2020 09:30;1 890;2 820;70;230;3 780;5 630;*

However, a quick script also exists to convert an entire .csv file to the first format should it be required. A sample input csv file has been included, namely *"Hillside.csv"*

