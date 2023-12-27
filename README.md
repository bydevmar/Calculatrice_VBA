# Calculatrice_VBA

## Description
Calculatrice_VBA is a simple calculator implemented in VBA (Visual Basic for Applications). This calculator allows basic arithmetic operations, including addition, subtraction, multiplication, and division. It also supports percentage calculations.

## Features
- Addition (+)
- Subtraction (-)
- Multiplication (*)
- Division (/)
- Percentage calculation

## Usage
1. Press numeric buttons (0-9) to input numbers.
2. Use the operation buttons (+, -, *, /) to perform arithmetic operations.
3. Press the "=" button to calculate the result.
4. Press the "C" button to clear the current input and result.
5. The "%" button calculates the percentage of the current result.

## Installation
This project doesn't require any installation. You can simply download the VBA code and run it using Microsoft Excel or another VBA-compatible environment.

## How to Run
1. Open Microsoft Excel.
2. Press `Alt` + `F11` to open the VBA editor.
3. Insert a new module and copy-paste the provided VBA code.
4. Close the VBA editor.
5. Add buttons to your Excel sheet and link them to the corresponding macros in the VBA code.

## Code Overview
The code consists of several macros:
- `btn_percent_Click`: Handles percentage calculations.
- `btn_egale_Click`: Performs the selected arithmetic operation.
- `btn_0_Click`: Appends "0" to the input.
- `btn_c_Click`: Clears the input and result.
- `btn_moin_Click`, `btn_plus_Click`, `btn_mult_Click`, `btn_div_Click`: Sets the operation for subtraction, addition, multiplication, and division.
- `btn_cinq_Click`, `btn_deux_Click`, `btn_huit_Click`, `btn_neuf_Click`, `btn_quatre_Click`, `btn_sept_Click`, `btn_six_Click`, `btn_trois_Click`, `btn_un_Click`: Appends the respective number to the input.
- `operationn(op As String)`: Sets the operation and stores the first operand.
- `numero(num As String)`: Appends a number to the input.

Feel free to explore and customize the code based on your preferences and requirements. If you encounter any issues or have suggestions for improvement, please open an issue or submit a pull request. Happy coding!