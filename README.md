# All Stocks Analysis using VBA 

## Overview of the project

Using VBA, create a code to run the stocks analysis for several years in order to be able to compare returns on investments for various stocks. 

###Purpose

Since Steve wants to determine which stocks are a better investment for his parents, the search was expanded to include the entire stock market over the last few years. 


##Results

First, let’s take a look at 2017 Stocks Analysis and 2018 Stocks Analysis as presented below.


<img width="273" alt="image" src="https://user-images.githubusercontent.com/107759305/199090741-efac3032-5314-4634-9450-7339cfac608a.png">

<img width="278" alt="image" src="https://user-images.githubusercontent.com/107759305/199090910-1be61193-37aa-4600-960f-fdedae582921.png">


Two graphs present us with a visual on the yearly return on investment for various stocks over the course of two consecutive years.  Formatting that was imbedded in our code helps us visually identify stocks that were positive for the investors. For example, ENPH stock was up by 129% in 2017, however, in 2018 that same stock fell to 82%. The good news, is that it looks like the ENPH stock is still a good investment because it has shown good standing in market volatility. Let’s take a look at another stock for a comparison. VSLR showed negative growth in 2018, however in 2017 it was up by 50%. It tells us, that this stock is not reliable, since, according to the 2 year trend data, shows a dip by almost 50%. Subsequently, VSLR may not be a good investment opportunity according the data. 

Overall, there are two stocks that remained “green” or positive over the course of two consecutive years: ENPH and RUN. Out of those two stocks, if Steve’s parents want to invest into one, I would recommend RUN stock. RUN stock had an increase on return over one year by nearly 80%, and we can conclude that the probability of this stock being successful in year 2019 is greater than in ENPH. 

When addressing the runtimes for All Stocks Analysis in the reformatted version versus the original, it was undoubtedly clear that the refactored version of the module ran faster by nearly 0.207 seconds. See the screenshots below.

![image](https://user-images.githubusercontent.com/107759305/199091062-b516a39b-d457-4179-81a8-27f5ee13fc27.png)

![image](https://user-images.githubusercontent.com/107759305/199091116-a763c372-23f8-412b-8901-b378400244ee.png)

![image](https://user-images.githubusercontent.com/107759305/199091173-d5044231-832c-4d83-957d-9ea20864f99b.png)

![image](https://user-images.githubusercontent.com/107759305/199091219-5027130b-60ff-489e-8d67-39be47cd3cee.png)


##Summary

1. What are the advantages or disadvantages of refactoring code?
2. How do these pros and cons apply to refactoring the original VBA script?

Advantages of the refactoring are imbedded formatting to create better visual for the client and faster run times. Unlike the original VBA script, reformatted code includes conditional formatting which presents the data for the viewer at a glance in color. Most people are visual, and want to be able to see information presented to them in a simple manner so that they can make an informed decision. The speed of run times plays a huge role when you are analyzing large portions of data spanning over multiple years and multiple worksheets. Faster processing speeds will cut down the time for the reviewer to analyze the data outputs. Original VBA script’s run time speed was 0.207 seconds slower than the reformatted version. 
