# Green stock analysis
performing analysis on green stocks data using excel and VBA
# overview of the project:
The project analyzes the green stocks data to invest based on total daily volumes and return.
## purpose:
The purpose of this challenge is to refactor the VBA code to loop through all the data one time to collect information from 2017 and 2018. 
# Results:
## Analysis:
Before refactoring the code, I downloaded the challenge_starter_code.vbs file, neatly written and the steps I needed to add, tickerindex, create output arrays, use nested for loops, and if statements. By writing the code and activating the worksheet, we have refactored the code, and we can now measure the performance of both years and check the script run time. 
Below is the refactored code is written in the file.

 '1a) Create a ticker Index
 tickerIndex = 0

'1b) Create three output arrays
Dim TickerVolumes(12) As Long
Dim TickerStartingPrices(12) As Single
Dim TickerEndingPrices(12) As Single

''2a) Create a for loop to initialize the tickerVolumes to zero.
For i = 0 To 11
TickerVolumes(i) = 0
   
Next i
        
''2b) Loop over all the rows in the spreadsheet.
  For i = 2 To RowCount

 '3a) Increase volume for current ticker
     TickerVolumes(tickerIndex) = TickerVolumes(tickerIndex) + Cells(i, 8).Value

    
 '3b) Check if the current row is the first row with the selected tickerIndex.
   
 'If  Then
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    TickerStartingPrices(tickerIndex) = Cells(i, 6).Value
       
  End If
    
'3c) check if the current row is the last row with the selected ticker
'If the next row’s ticker doesn’t match, increase the tickerIndex.
    'If  Then
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
     TickerEndingPrices(tickerIndex) = Cells(i, 6).Value
   
   End If

'3d Increase the tickerIndex.
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
    tickerIndex = tickerIndex + 1
    
End If
    

Next i
'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For i = 0 To 11
    
Worksheets("All Stocks Analysis").Activate
Cells(4 + i, 1).Value = tickers(i)
Cells(4 + i, 2).Value = TickerVolumes(i)
Cells(4 + i, 3).Value = TickerEndingPrices(i) / TickerStartingPrices(i) - 1
        
Next i

### compare stock performance b/w 2017 and 2018:
Based on the returns of 2017 except TERP every other stock has done reasonably well. But in 2018, it only did well. Suppose we consider that only 2 of the 12 stocks, ENPH AND RUN, have positive results in both years. Based on the performance of ENPH AND RUN, he can suggest his parents invest in these stocks. 

<img width="458" alt="Screen Shot 2022-03-12 at 11 08 16 AM" src="https://user-images.githubusercontent.com/100738688/158027710-8c95f316-1729-43cd-91c4-c186a1dd5c32.png">

<img width="418" alt="Screen Shot 2022-03-12 at 11 06 02 AM" src="https://user-images.githubusercontent.com/100738688/158027713-1369d609-0af4-41b9-b2b7-776313718fb9.png">





### execution time:

Refactoring the code can succeed as the execution time has improved for 2017 and 2018. if we compare the execution time of the original run time to refactored run time, the refactored run time is fast.

<img width="668" alt="Screen Shot 2022-03-12 at 12 01 02 PM" src="https://user-images.githubusercontent.com/100738688/158027720-1d1895da-0ad9-41f6-b532-89ecd8bc238d.png">

<img width="356" alt="Screen Shot 2022-03-12 at 12 01 43 PM" src="https://user-images.githubusercontent.com/100738688/158027725-6bf42d4f-3960-4ac2-882a-c016c2bf746b.png">

<img width="484" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/100738688/158027734-08707e0d-3752-47ab-91fe-9c5212556324.png">

<img width="554" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/100738688/158027738-2ba33111-a6bb-4f8e-9084-0a7992acaedf.png">




# summary:
## Advantages :
The considerable advantage of refactoring the code decreases the run time in executing the project. Decreased execution is always efficient when analyzing thousands of rows of data.
## Disadvantages:
When refactoring code, it's better to save your original code and any changes you make to the script, as errors can destroy an already working code. While refactoring code, I got many errors while using loops and if codes.
## pros and cons:
The refactored code looks clean and organized, so other users can easily understand. Reducing the number of steps can process the data faster and decrease the run time.
While refactoring the code, the most significant disadvantage I faced was including a new code as I was getting too many errors spent a lot of time debugging it. You should always pay attention in the process of refactoring the codes; you may face difficulty later. Refactoring can be a disadvantage for larger files.


   
   
    

       
            



 
 
       
      











