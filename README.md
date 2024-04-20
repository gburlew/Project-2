# Where I got my code, and my process:

Requirements 1 and 2:
I referenced a lot of my code from the credit_charges class activity, where I had also used xpert to get help with looping my code correctly.
I used xpert to help me create this function: "Range("o" & outputrow).Value = yearlychange / count" which references this one "yearlychange = yearlychange + (Cells(i, 3).Value - Cells(i, 6).Value)". I was able to come up with the latter function on my own, but I had trouble trying to calculate the total yearly change without dividing it by all of the cells in the worksheet. Therefore, it had me create the "count" variable, located after the first nested if statements.
Through trial and error, xpert also helped me figure out where to correctly initialize certain variables in the code, helped me add another "if" statement regarding the count, and had me create the variables openingprice, and quarterlychange (when i had originally had a full calculation in the code instead of making it more efficient with creating a variable).
Requirement 3: 
Xpert gave me the correct calculation to find the percent change between two values; also gave me the idea for "old value" and "new value" variables.
While it gave me a function for percentchange in our chat logs, I would have been able to think of the percentchange variable myself.
Xpert helped me convert ws.Columns("p2:P").Value.NumberFormat = "0.00%" into ws.Range("P2:P" & ws.Cells(ws.Rows.Count, "P").End(xlUp).Row).Numberformat = "0.00%".
It also helped with finding the correct location of where to calculate certain variables and where to correctly put a formatting function.
Requirement 4:
Xpert also helped me add & outputrow so that all data would be put into the output columns.
I had xpert help me troubleshoot where to calculate the total stock volume in the loop. Part of the issue was that "totalstockv" was actually supposed to be defined as a double, and not as a long. My way of calculation ended up being right the first time before I had asked xpert for help, but I had been putting it in the wrong location of the loop and had not defined it as the correct datatype, resulting in extremely confusing output values.

Conditional formatting requirement: I was unsure whether I was supposed to use the conditional formatting button that was already in excel, or if I was supposed to do conditional formatting using vba coding, so I did the latter. If this is not what I was supposed to do, I will fix it.
Looping through worksheets requirement: Although I knew the way to initially tell my code to loop through more than one worksheet, I  had trouble getting it not to loop through the same worksheet over and over again. After asking xpert for help and through a lot of trial and error, I realized that the issue was how I was adding my cell references. I hadn’t realized that I was supposed to add “ws.” In front of all of my cell references and ranges, which resulted in the code continuing to divide by itself and replace its own outputs. Xpert showed me how to add “ws.” to my cell references, and then my code was able to run through the entire workbook smoothly.

Other notes: 
I don't know if it matters, but in my code, I had named the variable that was supposed to represent the quarterly change as "yearly change", as I had originally misread the instructions. After testing, it all works correctly but I wanted to add that note to avoid any confusion.
Along with that, there is a lot of space between the raw data and the output results; you may have to scroll to the side to see everything.
