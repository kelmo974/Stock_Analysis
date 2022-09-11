# Stock Analysis with VBA

### Purpose
The main objective of this project was to harness the power of VBA and use that to analyze a dataset of stock tickers from the green energy sector. It was hoped that organizing data and conducting basic analysis of its metrics would provide enough insight to allow the stakeholder to make an informed financial decision. Provided data came from twelve different stocks and their performance metrics from the years 2017 and 2018. The VBA script would attempt to collect info on total volume of trades, starting price, and ending price. These values could be used to determine how much attention this stock had as well as it's annual return. These results could then be compared over two consecutive years. Formatting was to be applied in order to highlight market trends with regards to these equities.

From a programming language standpoint, this exercise served to demonstrate the possible advantages (or disadvantafes) of refactoring code. After the initial successful scanning of market data, the program was tweaked in an effort to leave it more efficient, but yield the same raw output.

### Results

The provided ticker data was successfully scanned by creating a for loop to that established a repeteable path as it pulled the data of interest from the worksheets. The below screenshot of arrays and the loop demonstrate this technique.

![Array_and_main_loop](https://user-images.githubusercontent.com/109499859/189533534-3d4bda09-e207-4cf7-b010-f2f9e632e4a4.PNG)


A functional VBA script was able to collect, sort, and aggregate the data for all twelve green tickers. This was made assigning specific variables so that total volume, starting price, and ending price values were collected from the excel workbook. See the following screenshot demonstrates this bit of language.


![Variables+all row loop](https://user-images.githubusercontent.com/109499859/189533668-bdbf1d3e-8481-4b88-8196-69903ffbde70.PNG)


Once the looping was completed, a series of if statements supported the sorting effort. By telling excel to include check if the next row was of the same ticker or a different ticker, it was able to differentiate between the various tickers and only assign the aggregated values of volume and prices to their respective categories. Those are included below:

![Ifs](https://user-images.githubusercontent.com/109499859/189533950-7cd688f0-fd15-4a97-8011-f588dc68fd55.PNG)

Near the end of the code, it was important to include some formatting to keep the data tidy as well as to highlight positive and negative yields for all tickers in either year. That was achieved by including the following bit of code:

![Formatting](https://user-images.githubusercontent.com/109499859/189534230-6f72e087-d9d6-44b7-8ed2-489ab4d39b6e.PNG)

Aggregrated results revealed that 2017 offered generally positive returns, whereas 2018 was largely a downturn for this collection of green energy stocks. The stakeholder will be able to draw conclusions from this data and offer sound financial advice. 

For the programmer, it was also important to see if this code could be made to run more efficient. So, during the intial and refactoring runs of this script, a timer was linked to a mesage box so that the run time would be displayed after each run.

These four screenshots show the run time difference of both 2017 and 2018 runs.

![Timer_2017_standard](https://user-images.githubusercontent.com/109499859/189534538-6d937da0-9c00-4954-8f1b-faf208e6a93c.PNG)
![Timer_2017_Refactored](https://user-images.githubusercontent.com/109499859/189534539-64ac9281-0155-427c-9d19-57589349e72b.PNG)
![Timer_2018_standard](https://user-images.githubusercontent.com/109499859/189534551-3513af4f-d528-4628-983f-f14f6e3b2734.PNG)
![Timer_2018_Refactored](https://user-images.githubusercontent.com/109499859/189534552-fbf0a717-4b1f-41e1-8824-100bb99ccb51.PNG)

### Summary

VBA allowed for a simple stock analysis that hopefully provided valuable insight for the stakeholder. As long as variables are well defined and assigned the correct data type, issues were minimal. The addition of timers added an additional challenge since simply submitting the original script would not suffice. In order to get faster, a series of trial and error tweaking was carried out. While modifying an already working code, it is difficult not the think about the mantra, "if it isn't broken, don't fix it." 

In this instance, the exercise of refactoring took substantially more time than the 0.6 seconds of run time that it saved. However, there are surely scenarios where saving on computational power and run time would be worth the extra effort of refactoring. At the very least, it is a great learning exercise as it stretches your understanding of the language and may lead to additional discoveries that could be used in different projects. 

Finally, the green ticker stock analysis was an overall success. The requested intel was provided to our stakeholder and refactoring was even beneficial for the runtime.
