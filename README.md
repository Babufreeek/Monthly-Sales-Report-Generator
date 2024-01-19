# Monthly Sales

## Usage of User Interface
- Ensure that the columns and values in the spreadsheet match their chines
- Open `runner.exe` by double clicking it.
- Firstly, select the spreadsheet you would like to process.
- Next, from the dropdown, select the worksheet with the raw data.
- Select the translation source file. (If a file containing the phrase `language translation` **case insensitively** is present in the current folder or any subfolders, this field will be autofilled with the information for that file.)
- If you would like to save the raw translated output as well, tick the `Save Translations` checkbox. The output will be added as the `worksheet name + "_translated"`
- Select Output method (Add to Current Spreadsheet/New Spreadsheet/Both)
- Click the Submit button.
- A pop-up/command-line message will be displayed once the files have been processed. 
- Click the "Ok" button if you would like to process any other files.

## Running via Command Line Interface
---------------

    # Install libraries
    $> pip install -r requirements.txt

    # Run program
    $> python3 runner.py

    # Convert program to .exe file
    $> pip install pyinstaller
    $> pyinstaller --onefile runner.py
    $> pyinstaller --onefile --noconsole runner.py

    # Run .exe file via the terminal
    ./runner.exe
---------------------------------------------------

## Files
- `monthly_sales_calculations.py` runs all the calculations and backend processing using the `total_sales()` function.
- `excel_form.py` is used to create to front end interface.
- `styles.py` contains the stylings for the front end interface.
- `runner.py` is used to run the program. 
 