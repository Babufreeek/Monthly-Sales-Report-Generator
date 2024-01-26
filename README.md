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

## Alternative Usage
- If you would like to translate the content only and exit, Click the `Translate Worksheet Only` checkbox. Then click submit.
- If the worksheet has already been translated according to the guidelines, you can check the `Worksheet Already Translated` checkbox.

## **Chinese Translations 
- Listed at the top of `monthly_sales_calculations.py` are the preset chinese characters for variables that we need to use in our calculations.
- If any of the Chinese phrases below has been changed, make sure to update the preset values in `monthly_sales_calculations.py` as well before running the program.
- Alternatively, if you run the program without updating the preset values, you will be prompted to key in the updated column/value name in the terminal for those values that have changed.

### Preset Chinese Characters
```
untranslated_column_names = [
    "项目", # Project ID
    "资源ID", # Resource ID
    "标识", # Resource Name
    "资源类型", # Resource Type
    "数据中心", # Region
    "计费类型", # Billing Method
    "配置", # Configuration
    "订单类型", # Order Type
    "订单起始时间", # Order Start Time
    "订单结束时间", # Order End Time
    "订单原价", # Unit Price, 
    "消费原价", # Usage Amount
]

# Preset column values for calculations
monthly_chinese = (
    untranslated_column_names[5], # Billing Type
    "按月" # Monthly
 )
delete_refund_chinese = (
    untranslated_column_names[7], # Order Type
    "删除退费" # Delete & Refund
)
```

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
 