# stock_analysis
Analysis stock of green energy companies

## Overview of Project

In this assignment, we created a VBA script to analyze stock results for our friend Steve so he can help his parents choose the best green energy companies for investing. The analysis considers:  

- **Daily Volume**: the total number of shares traded throughout the day to measure how actvively a stock is traded. 

- **Yearly Return**: the percentage difference (increase or decrease) in price from the beginning of the year to the end of the year.

The end product gave Steve the tools he needs to easily run the program on his own and generate results that are quickly and easily understood. 

## Results

### Stock Performance
Twelve different stocks were analzed over two years. The results clearly show that 2017 was a profitable year for investors of green energy but 2018 was not. Only two companies, ENPH and RUN had positive returns both years with ENPH far outperforming their competitors from the beginning of 2017 trough the end of 2018. 

![2017 Chart](https://user-images.githubusercontent.com/90162669/135764716-6891e396-b8ca-400c-afb3-7f2c35cad7e8.png)

![2018 Chart](https://user-images.githubusercontent.com/90162669/135764487-0c62f6ff-10ed-4c02-b97d-0845de5a49f8.png)

### Refactoring
The original code that was designed for this assignment had otwo loops (a loop within a loop) which required the program to run through the data hundreds of times. To improve run time, an addtional array was added to capture total volume, starting  price, and ending price. This allowed for the elimination of the inner loop. The refractured code below only loops through the worksheet twelve times creating a much more efficient code.   

![Refactored Code](https://user-images.githubusercontent.com/90162669/135765206-28d64019-7396-4bf1-a8d1-24cd06aac61c.png)


#### Improved Run Times

At first glance, improving efficiency from 7 seconds to 1 second may seem immaterial; however, the change in performance will be significant as Steve expands his dataset to look at more stock options. 

Run Times Before Refactoring:

![Original 2017 Run Time](https://user-images.githubusercontent.com/90162669/135769332-c01cc841-9caf-46a5-a1de-0fecd0ec4663.png)


![Original 2018 Run time](https://user-images.githubusercontent.com/90162669/135769409-c1af8009-3115-41ba-90ca-3ba768e2510c.png)




Run Times After Refactoring: 

![2017 Run Time](https://user-images.githubusercontent.com/90162669/135764522-d37c4b29-7d4e-4829-afbb-ca13e7918d4f.png)

![2018 Run Time](https://user-images.githubusercontent.com/90162669/135764540-ec812c94-31e9-4755-8c4e-6000ebb3580f.png)


## Summary
"What are the advantages or disadvantages of refactoring code?
"How do these pros and cons apply to refactoring the original VBA script?
"there must be a deteailed statementent of the advantages and disadvantages or refactoring code in general
"There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script

Could use different shades of green to better differentiate the levels or degrees of positve performance. 

