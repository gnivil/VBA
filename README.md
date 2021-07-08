# VBA-challenge

This code creates a summary table which breaks down the following for each stock...

    Ticker Symbol
        Displayed by pulling the symbol from Column A using a for loop
    
    Yearly Change
        Calculated by using a nested conditional formula within our for loop, comparing the value of the stock at the beginning of the year vs. the value of the stock at the end of the year

    Yerly Percent Change
        Calculated by dividing the calculated Yearly Change by the original value of the stock at the beginning of the year

    Total Stock Volume
        Calculated by creating a running total of the volume for each ticker


Conflicts
    My code has bugs in the for loop that I could not correct. Some possible culprits I investigated were...

    The value that closes out my for loop is 'Cells (i,1).End(xlDown)
    My variables may be incorrectly labeled
    The order of my conditionals within my for loop may be incorrect
    The for loop for cycling through each worksheet is not properly labeled

