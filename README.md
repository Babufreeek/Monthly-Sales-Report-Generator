# Monthly Sales

## Usage of User Interface
1. Ensure that the columns and values in the spreadsheet match with the chinese/english version in the Language Translation file
2. Open `runner.exe` by double clicking it.
3. Firstly, select the spreadsheet you would like to process.
4. Next, from the dropdown, select the worksheet with the raw data.
5. Select the translation source file. (If a file containing the phrase `language translation` **case insensitively** is present in the current folder or any subfolders, this field will be autofilled with the path for that file.)
6. If you would like to save the raw translated output as well, tick the `Save Translations` checkbox. The output will be added as the `worksheet name + "_translated"`
7. Select Output method (Add to Current Spreadsheet/New Spreadsheet/Both)
8. Click the Submit button.
9. A pop-up/command-line message will be displayed once the files have been processed. 
10. Click the "Ok" button if you would like to process any other files. Start over from **Step 3**.

## Alternative Usage
- If you would like to translate the worksheet only, Click the `Translate Worksheet Only` checkbox. Then click Submit.
- If the worksheet has already been translated and the english translation is based on the same Lnaguage Translation file, check the `Worksheet Already Translated` checkbox.

## Running via Command Line Interface
---------------

    # Install libraries
    $> pip install -r requirements.txt

    # Run program with python
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
 