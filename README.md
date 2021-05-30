# VBA_Challenge
Assignment 2

## Purpose of the Analysis
  
  The purpose of this analysis was to analyze overall stock performance in years 2017 and 2018. 
    
    
 ## Results - Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the             execution times of the original script and the refactored script.

 

  The stock performance for 2017 had more positive returns than in 2018. TERP was the only low stock in 2017. 
  ￼![image](https://user-images.githubusercontent.com/83786920/120090317-7514e900-c0cf-11eb-9bc3-eea19d42a3e0.png) 
  
  The stock performance in 2018 had a very big decrease with only two positive returns for ENPH and RUN. 
  ￼![image](https://user-images.githubusercontent.com/83786920/120090441-91fdec00-c0d0-11eb-9455-922fc4d9c518.png)  
  
  The execution time on the original script for 2017 was .6640625 seconds compared to the refactored script of .09375 seconds. We   can see there is a significant drop in the execution time for the refactored script.
  ￼![image](https://user-images.githubusercontent.com/83786920/120090622-fbcac580-c0d1-11eb-9129-74b2134b55ff.png)
  
  The execution time on the original script for 2018 was .65625 seconds compared to the refactored script of .6875 seconds. We     note there was a slight increase of seconds in the refactored script. The execution time for 2018 ran faster on the original     script.
  ￼![image](https://user-images.githubusercontent.com/83786920/120090671-4c422300-c0d2-11eb-8d08-722dca3671ec.png)
  
  ## Summary: In a summary statement, address the following questions.
  
  **What are the advantages or disadvantages of refactoring code?**
  
  One of the advantages of refactoring a code is that it saves you a lot of time. A second advantage is that you can use the code   from a previous module and make minimal changes without having to write the entire code again. A third advantage would be         producing a faster execution time, however, that is not always the case as I noted above. This would fall under a disadvantage   of refactoring code.
  
  **How do these pros and cons apply to refactoring the original VBA script?**
  
  These pros apply to refactoring the original VBA script because we are essentially running a similar analysis using the same     data. The comments and code styles that were used in the AllStocksAnalysisrefactored , were also used in the original VBA         script. The difference in the both scripts is the tickerIndex and the output arrays: 
                
                tickerIndex = 0
                Dim tickerVolumes(12) As Long
                Dim tickerStartingPrices(12) As Single
                Dim tickerEndingPrices(12) As Single 
                
   
 
      
