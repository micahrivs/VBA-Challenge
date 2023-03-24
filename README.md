# VBA-Challenge
VBA Homework - The VBA of Wall Street

## Background
I am excited to take on this homework assignment as I am on a path towards becoming a skilled programmer and Excel expert. The use of VBA scripting to analyze stock market data is a new challenge for me. However, I am eager to apply my knowledge and skills to this task. The ability to use VBA scripting to analyze stock market data is a valuable skill that will allow me to make informed financial decisions. I understand that this assignment may require me to spend significant time and effort, but I am willing to do so to ensure my success. Moreover, I believe that this homework will be a great opportunity for me to grow and develop as a programmer and Excel user.

### Stock market analyst


## Instructions

To complete this task, we'll start by creating a VBA macro that loops through all the stocks' data for one year and outputs the required information. Here are the steps we can follow:
*Create a new workbook and open the Visual Basic Editor (VBE).
*Insert a new module by clicking on Insert -> Module.
*Define the necessary variables for the macro: Ticker, YearlyChange, PercentChange, TotalVolume, LastRow, YearOpen, YearClose.
*Loop through all the rows of data and extract the necessary information:
*Check if we are still in the same Ticker, and if not, store the new Ticker and YearOpen information.
*If we are still in the same Ticker, then update the TotalVolume information.
*When we reach the last row of a Ticker, store the YearClose information and calculate the YearlyChange and PercentChange.
*Output the Ticker, YearlyChange, PercentChange, and TotalVolume information in the Summary Table.
*Apply conditional formatting to the YearlyChange column to highlight positive change in green and negative change in red.
*Format the PercentChange column to display as a percentage.
*Finally, add some formatting and headers to the Summary Table to make it more readable.

### CHALLENGES

1. To provide a comprehensive analysis of stock performance, my solution will not only display the current stock prices but also offer additional metrics such as the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". This functionality will enable users to gain a deeper understanding of the stock's historical trends and identify the stocks that have performed the best and worst over a specific period. Overall, my solution will provide users with a more complete picture of stock performance, empowering them to make informed investment decisions.


2. To enable the VBA script to run on every worksheet, i.e., for every year, with a single execution, you can make the necessary adjustments to the code. This can be achieved by creating a new subroutine, say "Main," as the entry point of the code and adding a loop that iterates through all the worksheets in the workbook. For each worksheet, the existing code can be modified to perform the required calculations or actions. Once the changes are made, executing the "Main" subroutine will apply the modified code to every worksheet in the workbook, providing the desired results.







