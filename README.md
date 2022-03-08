‘Stock Market Analysis:

#Stock Analysis

## Overview of the project

### For this project I created a spreadsheet that gave Steve the ability to quickly analyze the market return of twelve stocks related to green energy.  His parents were interested in investing in a green energy company called DAQO Energy.  Steve wanted to help his parents make the right decision of what to invest their money in.  He needed more information about the stocks and with the click of two buttons, he was able to easily see which stocks had a positive and negative performance during a specific year.  Steve loved how much time he was able to save while analyzing stock market information but he is concerned about diversification.  He has asked to expand the current dataset to include all of the stock market over the last few years to be able to identify other potential investments.

### In order for Steve to be able to use the code with significantly more data it will need to be adjusted to make it run faster and more efficiently. 

 
## Results

### The first analysis was for DAQO (ticker: DQ), since that is the stock his parents mentioned investing in.  We compiled the year, total daily volume, and return for DAQO by using the following:

#### 1. Looped through all of the rows of data by using code: For i = 2 To RowCount and took the following actions:

##### a. Checked if the information was specific to DQ by setting the ticker to “DQ” and  

###### i. Added the volume to the total volume calculation (If Cells(i,1).Value = “DQ” Then totalVolume = totalVolume + Cells(i,8).Value.

###### ii. Checked to see if the open price was the first price listed for DQ and if so, set it as the starting price (If Cells(i-1,1).Value <> “DQ” And Cells(i,1.Value = “DQ” Then startingPrice = Cells(i,6).Value.

###### iii. Checked to see if the close price was the last price listed for DQ and if so, set it as the ending price(If Cells(i+1,1).Value <> “DQ” And Cells(i,1.Value = “DQ” Then endingPrice = Cells(i,6).Value.

#### 2. The return was calculated by using the following code: Cells(4,3).Value = (endingPrice/StartingPrice)-1.  The results for DQ 2018 information are shown below:

<img width="196" alt="DQ Analysis" src="https://user-images.githubusercontent.com/99366022/157327701-d1454880-2a99-4db6-a5b1-7db3bbe1cd83.png">

### As you can see from the analysis, DQ lost 63% in 2018.

### To get a better understanding of the sector and identify other potential investments several green energy stocks were added to the analysis.  The following pseudocode was used:

#### 1a. Created a tickerIndex variable and set it equal to zero before looping over the rows of data.  This variable will be used across three output arrays pulling data.

#### 1b. Created three output arrays: tickerVolume As Long data type, tickerStartingPrice As Single data type, and tickerEndingPrice As Single data type.

#### 2a. Created a For loop to initialize the tickerVolume equal to zero.

#### 2b. Created a For loop to go through all of the rows of data.

#### 3a. Created a nested For loop that increases the current tickerVolume (stock ticker) variable and adds the ticker volume for the current stock ticker, using the tickerIndex variable as the index.

#### 3b. Created an If Then statement to check if the current row is the first row for the specific ticker and if it is to assign the open price as the tickerStartingPrice variable.

#### 3c. Created an If Then statement to check if the current row is the last row for the specific ticker and if it is to assign the close price as the tickerEndingPrice variable.

#### 3d. Create code that increases the tickerIndex if the next row’s ticker doesn’t match a previous ticker.

#### 4. Created a For loop to loop through the arrays (tickers, tickerVolumes, tickerStartingPrices, tickerEndingPrices) to output the ticker, Total Daily Volume, and Return for each unique ticker in the tickerIndex in the data set.  Code was re-used from the DQ analysis along with new code being created to automate the year information was pulled for and to ask a question for input of the year to use for the data pull.

### Shown below are highlights of the code described above:

<img width="549" alt="VBA_Challenge_Refactored_Code" src="https://user-images.githubusercontent.com/99366022/157327843-336fefb6-31c2-47e9-bdd5-d275bc808bf4.png">

### Results:

#### The results for 2017 show that DQ, ENPH, FLSR, and SEDG all had returns over 100% for the year, with three of the four significantly over 100% returns.  The results also show all but one of the stocks had gains.  The results for 2017 as well as the amount of time it took to run the analysis are shown below:

<img width="447" alt="VBA_Challenge_2017 - Final" src="https://user-images.githubusercontent.com/99366022/157327930-a12307ae-1934-4978-b758-1bdba29d512c.png">

#### The results for 2018 tell a much different story.  In 2018 only two of the stocks had a positive return, ENPH had a return of 81.9% and RUN had an 84.0% return.  DQ at -62.6% had the highest negative return of all the stocks analyzed for 2018.  The results for 2018 as well as the amount of time it took to run the analysis are shown below:

<img width="438" alt="VBA_Challenge_2018 - Final" src="https://user-images.githubusercontent.com/99366022/157327980-aaa4e8bb-b1a5-4c01-a9c8-affa70dcb385.png">

## Summary

### The advantages of refactoring the code made the code run faster.  Since Steve wants to use the code for a much larger set of data the efficiencies gained will be beneficial.  The only disadvantage I can see to refactoring the code is the amount of time that was spent in completing the work.  Otherwise, I don’t see any.  You can see the difference in performance between the original and refactored code time:

<img width="332" alt="2017 Run times" src="https://user-images.githubusercontent.com/99366022/157328027-c393a9ea-3d06-4c6c-bf5d-d3a4253a42bd.png">

<img width="335" alt="2018 Run times" src="https://user-images.githubusercontent.com/99366022/157328087-8fe2023a-6a82-49f6-b4a5-50be43da00b4.png">

### The pros of refactoring the original VBA script are the current code is more detailed than the original code.  The current code also automated a few pieces of the original code that was specific to a spreadsheet and/or year.  The current code will run either spreadsheet and year without having to manually make changes to the code.  The refactored code has more flexibility, runs faster, and it can handle larger sets of data.  I haven’t found any cons for refactoring the original script.
