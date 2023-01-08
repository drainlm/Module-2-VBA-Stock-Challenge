# Module-2-VBA-Stock-Challenge
Module 2 Stock Challenge

I have written a VBA code that contains three for loops: one loops through each worksheet and the other two each produce a summary table. 

I have done so by beginning with a For Each loop so that the script will run on each worksheet in the workbook. 

I then declared a lot of variables. I came to regret my choice of using Consolodated because I learned that's a tricky word for me to type accurately, but nevertheless I clung to that decision anyway. 

I then created the first summary table, which produces the yearly change, percentage change, and total stock volume for each ticker symbol for one year (each worksheet containing a separate year). I created a for i for loop to pull the given data and output the desired data. This entailed ensuring that each ticker name pulled properly and the correct associated values printed into the summary table. I also used color conditional formatting to highlight the yearly change in red for negative and green for positive change. 

I then created the second summary table, which produces the greatest percent change increase, the greatest percent change decrease, and the greatest total volume for each year using the data from the first summary table. This was the trickiest thing to try to accomplish of anything. And I can only hope that this counts as the conditional formatting that needs to be applied for the percent change column (as referenced in the grading requirements). I created a for j loop to pull from summary table one's columns for percent change and stock volume. After many struggles and errors, I was able to find success by reorganizing the way I wrote my code multiple times. Like I clung to the word "consolodated", I tried to cling to a decision to use the max/min function before finally giving up in defeat. I then clung to a decision to print within each of the three If statements because I thought my approach was clever and elemenated one line of code for each. It was not and I learned that the hard way. 
    
I also included an auto format piece so that if the script is working on a fresh excel, then everythiing lines up correctly. 


Additional things that helped me as I trudged through this assignment: 
I produced a second module to serve as a reset button and am still patting myself on the back for doing so from the beginning. 

Sub Reset_Button()
For Each ws In Worksheets
    ws.Range("I:Q").ClearContents
    ws.Range("I:Q").Interior.ColorIndex = 0
Next ws
End Sub

I also made use of just putting a formula in to find the min/max for the summary table two values so I could check my work. I figured that out kind of late, but I'm glad I did.


Anyway, that's what I did. Thanks for reading. 
