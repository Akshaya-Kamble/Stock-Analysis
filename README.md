# Stock-Analysis
Stock Analysis using VBA

# Overview of Project
To refactor the All stock Analysis VBA code in Module 2 to loop through all the data one time (Rather than 12 times) in order to collect 
the same information and to determine whether refactoring the code successfully made the VBA script run faster. Using refactoring we are going to help Steve
include the entire stock market over the last few years.The VBA code will include inputbox for user to enter the year, adding timer to calculate the run time of code,
use of multiple arrays referencing the same index, for loops, if conditions,static formatting and conditional formatting.


# Results:
## Comparing the stocks performance between 2017 and 2018
The total volume for stocks have been fairly same in 2017 and 2018 but there is a huge difference in the return ratios.
Majorty of the 2018 stocks did not perform well with only 2 positive stock return, Where as the 2017 stocks did great and a majority of stocks had a 
positive return. So the 2 stocks that did well in 2017 did not have a positive return in 2018
Eg : Refer Images in resource folder with name 2017 and 2018 comparison.png
 
## Comparing the run time for the 2017 and 2018 stocks using original code and refactored code.
Comparing the time taken to run the code for 2017 and 2018 in original code and refactored code significant different.
The refactored code gives us the same output in less time compared to the original code.

For 2017 the original code run time was 1.19 seconds and the refactored code run time was 0.18 seconds. Similarly for 
2018 the original code run time was 1.14 seconds and the refactore code run time was 0.19 seconds.
Eg: Refer to the below names of screen shots in the resource folder for the recorded runtime for 2017 and 2018 using original and refactored code
	1. VBA_Challenge_2017.png
	2. VBA_Challenge_2018.png
	3. Module 2 timer 2017.png
	4. Module 2 timer 2017.png

This proves that refactored code is an efficient way to reuse code and it also reduces the run time by avoiding unnecessary
repetation of code.
In the previous code without refactoring we have three if conditions inside the for loop,and the for loop runs till the row end.
So all the if statements also run till rowend and this multiple execution of similar code has increased the run time.
In the refactored code , there are more arrays addded referencing the same index once and this is making the code run faster.  


# Summary: 
## What are the advantages or disadvantages of refactoring code?
### Advantages
1. Refactoring takes fewer steps, using less memory or improving the logic of the code to make it easier for future users to read.
1. Refactoring leads to better quality code. 
2. Makes code easy to read and understand.
3. Improves design of code.
4. Helps programming faster.
5. Removes unnecessary repetation of code.

### Disadvantages
1. It takes a long time to code when the application is big.
2. It requires lots of regression testing and debugging at every level of code.


## How do these pros and cons apply to refactoring the original VBA script?
### Pros
1. The code was very readable and there were multiple explanations for the code in comments
2. The refactored code ran faster than the original VBA code, screen shots of images avaiable in resource folder.
3. The use of Arrays indexing the same reference at one time made the code run faster as repetation of steps were minimized.

### Cons
1. Since the code is being refactored , testing is required at small stages making sure that the code is looping well, and this is time consuming.
2. Refactoring the entire code is also like working on the same code again in a different way. This would reduce the code runtime and increase the efficiency of any programm but wold increase the manual work time.
