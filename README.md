# Monthly Sales Report Generator
- Written by Chew Wei Jie
- This program allows you to automate the process of creating a monthly sales report for hourly and postpaid customers.

## Usage of User Interface
1. Ensure that the columns and values in the raw data spreadsheet match with the chinese/english version in the Language Translation file
2. Open the program by double clicking it. Ignore any anti-virus warnings!
3. First, select the spreadsheet with raw data by clicking *Browse* and navigating using the File Explorer pop-up.
4. A loading bar pop-up will then appear. Wait for the loading bar pop-up to close or for the *Select Worksheet* dropdown field to populate. This may take a while and the program may appear to hang.
5. Once the *Select Worksheet* dropdown field is populated, you can select from the dropdown the worksheet containing raw data. 
5. Select the translation source file in the *Translation Source File* if needes. (If a file containing `language translation` **case insensitively** is present in the current folder or any subfolders, this field will be autofilled.)
6. To save the raw data translated to English, select the *Save Translation of Raw Data* checkbox. A new excel file will be created containing the translated raw data. (Filename `en_worksheetname`)
7. Select Output method (*Add to Current Spreadsheet*/*Create New Spreadsheet*/Both)
    1. *Add to Existing Spreadsheet*: Adds the sales report (as a new worksheet) to the spreadsheet selected in *Step 3*. Users can specify the worksheet name. **Note: If you select this option, please close the spreadsheet file before moving to the next step!**
    2. *Create New Spreadsheet* (Faster method of output, selected by default.): Outputs the sales report into a new spreadsheet. Users can specify filename and location (autofilled with path of spreadsheet selected in *Step 3*) for the new spreadsheet as well as the worksheet name. Note: **Ensure you do not have a file with the same filename as the output, it will be overwritten!**
8. Click the **Submit** button at the bottom of the user interface. A similar loading screen pop-up will appear, and the program may appear to be unresponsive and hang.
9. A pop-up will be displayed once processing is finished. 
10. Click the **Ok** button. Start over from *Step 3* if you would like to process any other files!

## Alternative Usage
- If you only need to save the translated raw data, follow Steps 1 - 5. Then, select the Translate Raw Data Only checkbox. Lastly click the *Submit* button.
- If the raw report has already been translated according to Step 1, follow Steps 1 - 6. Then, select the *Raw Data in English* checkbox and continue from Step 8.

## Running via Command Line Interface
---------------

    # Install libraries
    $> pip install -r requirements.txt

    # Run program with python
    $> python3 runner.py

    # Convert program to an.exe file that does not print to the command-line
    $> pyinstaller --noconsole runner.py

    # Convert program to a single .exe file
    $> pyinstaller ---onefile --noconsole runner.py

    # Run .exe file via the terminal
    ./runner.exe
---------------------------------------------------

## Files
- `monthly_sales_calculations.py` runs all the calculations and backend processing using the `total_sales()` function.
- `excel_form.py` is used to create to front end interface.
- `styles.py` contains the stylings for the front end interface.
- `runner.py` is used to run the program. 
- `helpers.py` contains additional classes for the loading pop-up and File processor for running slower processes/
 