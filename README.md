# VBA-challenge
Main bones of the code was taken from the credit_charges daily task.

Mulitple Worksheets
The code to run over multiple worksheets was from Stack Overflow.
https://stackoverflow.com/questions/43738802/how-to-apply-vba-code-to-all-worksheets-in-the-workbook
Dim ws As Worksheet
'Walk through each worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate
'Add your code here
Next ws

Percent Change
Got the formula from here: https://www.bls.gov/cpi/factsheets/calculating-percent-changes.htm#:~:text=To%20find%20the%20percent%20change,multiply%20the%20result%20by%20100.
Rounded it down from here: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/round-function
I could not figure out how to successfully add the "%" symbol to the answer

Open_Stock and Close_Stock
I wanted to add an inner for loop to look for different variables for Open_Stock. However, after several visits with the tutor I couldn't get it to work successfully so I removed it and just have the loop from the credit_charges assignment. Whne I had the for loop for Open_Stock I would get an Error 13, so I just removed it.

Cell Color
I got the basic code for the If loop for color from here: https://www.mrexcel.com/board/threads/vba-code-to-make-negative-numbers-red-and-positive-numbers-black.1077036/

Max Year_Change, Percent_Change, and Total_Volume
Couldn't figure this out.


