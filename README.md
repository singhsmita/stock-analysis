**Stock Market Analysis**

  

# Overview of project

Steve aims to analyze the returns of various stocks in the stock market over the last few years.In order to help Steve to analyze multiple stocks and see their return over the years it's preferable to write a flexible macro which would enable him to do so at the click of a button.

## Purpose

The purpose of this project is to refactor code used in Module 2 to analyze stock data within our workbook.In order to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.We use the starter code to loop through all the data one time in order to collect the same information as Module 2. Then, we need to determine whether refactoring the code successfully made the VBA script run faster.

  

## Background

I am performing data analysis for green-energy stocks (for the year 2017 and 2018).I am primarily focussing on their yearly return and total daily volume parameters.

For this data analysis I am using Microsoft Visual Basic for Applications (VBA).To determine the results of the above mentioned parameters I have made use of conditional statements, for loops, static and conditional formatting.In order to improve its efficiency I have refactored the code.

Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task.

  
  

# Results

## Analysis of stocks from 2017 and 2018.

The table below displays the analysis from the year 2017 and 2018 for a set of 12 stocks.

  

![](https://lh5.googleusercontent.com/x-AXxL8i98JNjlm-9nmMk3cbvgvrCCqjGm3KYHoRRi7gWyMQ8gaW0VSwPBce-vTA5kkO70Zy9atm4kZVukCxUoCPVhemcfH8O3NhvojsNwakzxZQ9WVA-YBaJsf33mniQF7oDprVihmc4J7j4KKWtkcCRcuXMnPEZrikH1xtp91zprgyQG37e0YBOBnx)

  
  

The table displays the value of -Ticker , Total Daily Volume and Return.

  
  

### Comparison of Stocks between 2017 vs 2018 based on Daily Volume and Return.

  

From the screenshot above we can see that 2017 had a high ratio of positive yearly returns except TERP which had a negative yearly return. Upon analyzing data from 2018 it is seen that the majority of the stocks had negative returns except RUN and ENPH.

  

These results indicate an unstable trend of the stocks as far as we consider the data set for these two years.

  

### Comparison of code

Now,let's compare the refactored code and the original code from Module 2.

  
  

Refactored code

  

![](https://lh6.googleusercontent.com/8lV3DUSPHun1_-IgtTR6BqZOkXWARLEIazxNfWpRuMsrMgdWQH9lOZu3drifnzQSKM1B_3jyvD1Z7WSNMT5SXK9WAad82zVd8GXTYnSigk3kEun5YKwebLO78OF-wEGldBl628AEgFCF70r3TllErC5UXeAqd58CYDU99bkR0Ig0TfXAGWbJFLCz0SC-)

  
  
  
  
  
  
  
  
  
  
  
  
  

Original code

  

![](https://lh3.googleusercontent.com/e5cG6XI5BMknHyKummtE8IoALSmsd4yk2FZzOBzOAHP1lHYBDz-DAmNYD9D7cj3Fa4UnvwjnGL2hRvZoj20ufxDCgVBR9DDv5Nkq3bbXrn9DNWoHrFWpcwohLGMwSPxo-V15h1oSnUc4c-cNWbiHPJem63GdO-q0oVOXT1fSEsIkxjVt8gAtfPbs29ex)

  
  

Upon comparing both the codes we can see that in Original code we had used nested loops whereas in Refactored code we loop through the data one time and collect all the information in arrays.

  

In the refactored code ,the output of the Total Daily Volume and Return is calculated based on the below logics:-

  

-   Create a tickerIndex variable and set it to zero.
    

![](https://lh6.googleusercontent.com/R0FBpqsXJwZ3IsnBWdaJd0b4FN544IBq79RWSMmnp7EcnnAKktJLYd4b-WqebsE-oymwYa1FrvSTcU5f3owJXPgXTXXDECmKcceQUzMAwqd1nyeji-n6g5ag0LcBH3F7PZdFDrPJ0Tdm1CRDLY4Q53Juz794Mfesbwi0CYrha8c5NlwSPZgSWxiXgDpL)

-   Create 3 output arrays for tickerVolumes , tickerStartingPrices , tickerEndingPrices.
    

![](https://lh5.googleusercontent.com/ijqarnx5rXGNaiCQSTOZU_h7grbaqOuGoP25QAcwNptoZlN9hZNzsOcCtsDbHL6ycLT7lB96hRQaxTyvAgRfw92Stnw417Vybyml3VIM2ULPSPt1ICEk6DsjsrxxSa1XLgmCMN9IJnveVKx5R4VwkhNGaNSbJIK16a1sGTw2mdQ9o0zUnAm851p3i-bB)

We create an array tickerVolumes(12) that holds 12 elements, and use our new variable tickerIndex to access the ticker index in order to store the right value for the ticker.

  
  

-   Create a for loop to initialize tickerVolumes to zero![](https://lh3.googleusercontent.com/R5NtkcvH0OdJbKzSyVhERN43em9WVWk3X9s6_vjr8grMOgM9nwGoiqdB4wUQWCWhvLkOUaQTdszA5AVcX8Dw6A_1kPiCnfN3XYzCfYnuPKRsDNgBJDizckRtzZ6ogwGFAavoc-Axv37urMfL0MGX8PNK76aTJgQRuMtaBGwTaEkmswQXOQ11XblM_57W)
    
-   Write if-then loops to calculate tickerStartingPrices and tickerEndingPrices with the selected tickerIndex.
    

![](https://lh6.googleusercontent.com/r1nLG21Cs2oqQyAYK2QMj_Ndb51uugdUw0fy4AD8GH6_DtdUGfKrkYrlnWY-DETWkaApkTcUzAHOiAaqaXJFNfslIl1q3KM1e3TXeVGMrWzA4HRjvk0ayw2aUfkCoG0ZzSb7sNtUulF9gJv7LkSlnHs9m1gKEI-m6q3nFSEBFtIgFpKgqHV8PMbBhA5q)

-   increase the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.
    

![](https://lh6.googleusercontent.com/PFLNJQexdkRRpKhi_g81IhTj9Rz-i83nhDRruunjgCj53ae2lspYiKtMYCsDLMVBdf_bezFXPO5U_OvdJ3ge10IHBpTmCEQOj2Jx5iGhVRiMSHBTWrxu1-3wgi_G4AnQsIp55iy_4D0xyxCIzoy40Nr5n9pOVsYWzaCs-YyoXy3-Q_45kCCrb8qAM-d9)

-   Create another loop to output the ticker, Total Volume and Return through the arrays.
    

![](https://lh5.googleusercontent.com/BANpa-ZBTZ517jv9j1jPFnCOlxjTw8FwMCK7L07Agx6dLPlNy0Eo-7HI8RYFipdDng2hzhNIHxK2-13J-p_hrE6M9K3DGdpH7Nbu7q0_lebxOR74wxQzwl5I25CHIu2BBQM3U3lohgtmZGUubYwoYlLXQQOdWPKMqt87U8gte8Cn4yNg70RiG4eSaq9W)

  
  
  

Both the codes have the same output.However,we see a significant difference in the time elapsed in returning the output of both the codes.

  
  
  
  
  

Time elapsed in original code Time elapsed in Refactored code

  

![](https://lh3.googleusercontent.com/ijGQnM78cAh4xtAM3d1wHt1VyDCHqKY9qaNuJw5KQUmKRB_7j6D8AaWxSa7xCaf3pDk0I25NbsQqUHBwv2XQ-z4DLc9cpMS4SSSvRDm-pp2KbQeyhrBfl6aE53wuRCBz0XCnRuxTsT60FXSezWUtovanD5wZxtlag6KxkFCDsXVxp2y2w3tBcyqEXc6f)

  
  
  

Time elapsed in original code Time elapsed in Refactored code

  

![](https://lh3.googleusercontent.com/D4rqNylrcy5nCt51O5-nXFnkoPPAl05VgEStqmbQQZgRWxnpqqmGYiMDW4UOSCWBefoW-xyf8hB5GXM4RubnIcvCS_3jAmZOQDQ92vr_xRe32_eaViuuASBs-to4IVXHakotQc-7jHomANR7xl3SgS6UMvO58lCJ65QrgDdPLa33bxLkfRd6HCbg_cMT)

  

We can see that the refactored code is almost 6X faster than the original code.

  

# Summary

  

1.  ## What are the advantages or disadvantages of refactoring code?
    

### Advantages

-   The refactored code runs faster than the original code.So,it can analyze large datasets in a shorter time.
    
-   The code is more readable and well structured.This makes error detection easier when the code contains nested conditionals and loops.
    

  

### Disadvantages

-   Sometimes, refactoring someone else’s code can be really tedious as we might not be aware of the background and functionality.
    
-   In the process of refactoring the code we may end up making the code less efficient as it is almost a hit and trail methodology.
    
-   There is also the issue of the code not working once reworked.
    

  

## How do these pros and cons apply to refactoring the original VBA script?

  

Refactoring code was a great way to explore finding alternative methods to a previously successful one.

It also allowed further opportunities to debug different types of coding issues.

  

I realized that refactoring an already working code may or may not work as expected as we start making changes to it.I did come to a point where upon making changes to starter code my output for return and daily volumes were not matching the Module 2 results.But after spending some time on debugging I could fix the issue.

However,on successful refactoring ,the new code ran faster , almost 6 times faster .This Having worked with both, I do like the speed and efficiency of the refactored code, but knowing its more temperamental with respect to data corruption, It may require some more refactoring to work through corruptions.
