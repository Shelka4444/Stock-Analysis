# Stock Analysis with VBA

## Overview of Project
This analyzes stock investments using VBA in order to develop a financial strategy for clients.

## Results
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

It is evident from the analyzes shown below that the selected group of stocks fared far better in 2017 than in 2018. Of the twelve separate potential lines of investment, eleven stocks showed a positive rate of return in 2017 versus only two stocks with a positive rate of return the following year. It is therefore recommended to invests in stocks ENPH and RUN as they are the two investments which deliver an above increased rate of positive return from both 2017 and 2018.

Original Code 2017            |  Original Code 2018
:-------------------------:|:-------------------------:
![Original 2017 ](https://github.com/Shelka4444/Stock-Analysis/blob/main/Resources/2017%20Original.png)  |  ![2018](https://github.com/Shelka4444/Stock-Analysis/blob/main/Resources/2018%20Original.png)

Each of the graphic images displayed contain a timer to show the execution times of the code for each year (2017 and 2018). With the refactored script, the run time for 2017 decreased from 0.84375 seconds to 0.1796875 seconds and the run time for 2018 decreased from 0.8007812 seconds to 0.1640625 seconds. While the original script functioned correctly, decreasing the run time with the refactored code will let the system perform at a higher rate of efficiency which will have a large impact on the quality of output for larger data sets than the one demonstrated in this project.

Refactored Code 2017            |  Refactored Code 2018
:-------------------------:|:-------------------------:
![Refactored 2017 ](https://github.com/Shelka4444/Stock-Analysis/blob/main/Resources/2017%20Refactored.png)  |  ![2018](https://github.com/Shelka4444/Stock-Analysis/blob/main/Resources/2018%20Refactored.png)

## Summary

- Refactoring code is a vital part of creating more efficient and logical programs which will benefit both developers and end users. As datasets change and potentially grow exponentially larger, the ability to write cleaner and clearer code will hopefully reduce run time and ergo produce results in a much smaller time frame. The reduction in time and effort to maintain scripts which are vital for a company also reduces overall cost for managing the technical details for such programs. It is imperative, however, to verify that the newly scripted code is indeed a better, faster, cleaner version compared to the previous version else one runs the risk of encountering more problems and headaches which is, needless to say, counterproductive to the purpose of refactored code.

- For this project, the original VBA script ran an analysis for 2017 and 2018 in 0.84375 seconds and 0.8007812 seconds, respectively. The refactored code cut the run time to 0.1796875 seconds for 2017 and 0.1640625 seconds for 2018 to produce the same results (as seen in the images in the previous section). Since the results between the two scripts are identical but the run time is reduced by the refactored VBA script, the refactored script is clearly more efficient than its predecessor.
