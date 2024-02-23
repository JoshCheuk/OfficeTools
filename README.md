# **Project Base**

This project is based on python. In order to successfully execute programs in this project, you will need the following python package:

## Pandas

You can download this package by pip:

```
pip install pandas
```

## Openpyxl

You can download this package by pip:

```
pip install openpyxl
```

## Pyqt5

You can download this package by pip:

```
pip install pyqt5
```

For more information on pip, please refer to [Pypi Official Page](https://pypi.org/project/pip/)

# **Using the program**

## Preparation

In order to use the program, the excel files (xls or xlsx) should have the same header, and should at least contain information of:

1. Vendor Name
2. Entry Date
3. Debit Amount
4. Credit Amount
5. Account Name
6. Account Code

The order of the above six columns can be flexible but should be consistent among all the files you expect to operate.

## How to use

After executing the python files, first you will be asked to select the target folder where you store all the excel files you expect to operate. Remember to exclude irrelevant excel files.

Once the selection of target folder completes, the program will automatically read the excel files.

Then, you will be asked to enter the report date. Please enter in the format of *YYYY-MM-DD*, for example, if you are preparing a report closing on December 31, 2023, then you should enter *2023-12-31*.

Next, you will be asked to map different columns to match vendor name, entry date, debit amount, credit amount, account name and account code. The program reads and lists all the available columns. Please enter the column number according to the list from the program.

After mapping the columns, you will be asked to select the accounts that need calculating the age.

For the CLI version, you need to determine the number of accounts you wish to conduct calculation by entering an integer. Then, you need to enter the account codes for these accounts manually.

For the GUI version, a window will pop up and you need to select all the account that need calculating by checking the boxes next to the account codes and names combo. 

Finally, you just need to wait for the generation of the aging reports.

# **Other information**

The calculation runs on a FIFO (first in first out) basis. If there is material irregular transaction, please note the impact of this assumption.

The positive value indicates "debit" side. The negative value indicates "credit" side.

The program classifies ages into these groups:

1. Within 1 month
2. 1-3 months
3. 3-6 months
4. 6-12 months
5. 1-2 years
6. 2-3 years
7. Over 3 years
