# Excel VBA Amortization Table

This project implements an Excel VBA script to generate an amortization table for a loan. The user can input essential loan information such as the annual interest rate, loan term, and loan amount, and the script will create a dynamic table showing payment amounts, interest, principal, and remaining balances for each period of the loan.

## Features

- **User-Friendly Interface**: Utilizes input boxes to gather loan information, ensuring accurate data input.
- **Dynamic Table Creation**: Implements a loop to calculate and populate the amortization table dynamically.
- **Formatting Automation**: Utilizes VBA macros to automate formatting tasks, enhancing data presentation and readability.


## Macros Overview

- `Input_Annual_Interest_Rate_1`: Prompts user to enter the annual interest rate.
- `Input_Number_of_Years_2`: Prompts user to enter the number of years for the loan term.
- `Input_Loan_Amount_3`: Prompts user to enter the loan amount.
- `Clear_Data_4`: Clears data in the amortization table.
- `Input_All_Info_5`: Fills in loan information in designated cells.
- `Amortization_Table_6`: Generates the amortization table.
- `Proper_Case_A1_A4_7`: Changes titles in cells A1 to A4 to proper case.
- `Capitalize_All_A6_E6_8`: Capitalizes column titles in cells A6 to E6.
- `Bold_A6_E6_9`: Makes column titles in cells A6 to E6 bold.
- `Background_Blue_10`: Applies blue background color to column titles in cells A6 to E6.
- `Display_Monthly_Payment_11`: Calculates and displays the monthly payment in cell E1.
- `Total_Interest_12`: Calculates the total interest paid over the course of the loan.
