# VBA of Wall Street

## Overview of Project

Refactoring of a VBA script.

## Purpose

The purpose of this analysis is to refactor the code for a stock performance evaluation workbook.

## Results

In the original VBA script, each ticker looped through the entire worksheet to get its output. In the refactored VBA script, all of the output was obtained in one pass. By eliminating 10 loops, the script was able to run roughly 6x to 7x faster. 

##### Original Runtime 2017
![Original Runtime 2017](https://github.com/BiscuitButter/stock-analysis/blob/master/Resources/green_stocks_2017.png?raw=true)
##### Refactored Runtime 2017
![Refactored Runtime 2017](https://github.com/BiscuitButter/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png?raw=true)
##### Original Runtime 2018
![Original Runtime 2018](https://github.com/BiscuitButter/stock-analysis/blob/master/Resources/green_stocks_2018.png?raw=true)
##### Refactored Runtime 2018
![Refactored Runtime 2018](https://github.com/BiscuitButter/stock-analysis/blob/master/Resources/VBA_Challenge_2018.png?raw=true)

### Challenges and Difficulties Encountered

This analysis presented several challenges. The first problem I faced was with my tickerIndex. Thinking that I needed to set the tickerVolume variable to 0 after each increase in the tickerIndex, I had defined tickerVolumes(tickerIndex) = 0 in a place that caused my tickerVolumes output to always be 0. After some deep breathing, it occurred to me that tickerVolumes(tickerIndex) starts at 0 by default. Once I removed the line, I saw output in Total Daily Volume.
My second problem was with my output on Return. I had declared my variables for tickerStartingPrices and tickerEndingPrices as Long instead of Single. Thinking that something was wrong in the Loop, I analyzed and rewrote sections of script to no avail. Once again, after some deep breathing, I read the script in its entirety and found my mistake.
The last annoyance that I had was typos in my variables. I found that by declaring my variables with some uppercase letters and typing in all lowercase, I was able to easily catch my mistakes. I can see the value in declaring Option Explicit and may start doing that going forward just to save myself some headache.

## Summary

Refactoring code is tedious. It is very easy to break something that is already working as intended. However, as we have seen in this scenario, the benefits from refining code can be rather immense. 
The reduction in runtime for this code makes it much more ideal for handling larger quantities of data. While a runtime of 1 second doesn't seem very long, its inefficiency could have a much larger impact if we were to analyze 100 or 1000 tickers instead of 12.
