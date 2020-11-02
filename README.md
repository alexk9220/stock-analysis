# stock-analysis
Stock Analysis rep

#**Overview of project

Steve needed a workbook to show to his parents so that they can decide on their investments. With the help of VBA we have completed the task for Steve, however, afterwards,  we have also tried to refactor our code to make it work faster. Based on the outcome, we have successfully accomplished our goal.


#**Results 

Due to refactoring, our code now runs almost 5 times faster than it did originally as we can see in the screenshots below(2017 results and 2018 resuts respectively):
**2017
<img width="1440" alt="VBA_Nonref_2017" src="https://user-images.githubusercontent.com/73204192/97829520-5749e780-1c98-11eb-97a6-457d678a02e0.png">
<img width="1440" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/73204192/97829638-9e37dd00-1c98-11eb-9f9d-b526204e8207.png">

**2018
<img width="1440" alt="VBA_NoNref_2018" src="https://user-images.githubusercontent.com/73204192/97829705-cd4e4e80-1c98-11eb-902c-6d816b8f9e95.png">
<img width="1440" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/73204192/97829729-e8b95980-1c98-11eb-841d-553a526988c6.png">

Let's take a look at our code and see if there are any differences that we can notice straight away.
<img width="1440" alt="All_Stocks" src="https://user-images.githubusercontent.com/73204192/97830446-fec81980-1c9a-11eb-8aad-c158d4f6c313.png">
<img width="1440" alt="All_Stocksref" src="https://user-images.githubusercontent.com/73204192/97830475-199a8e00-1c9b-11eb-88fb-ddba5bfe16c1.png">

As we can see in the very begginig of our code, we have only introduced one array in All Stocks Ananlysis while Volumes/Starting/Ending Price have been listed as variables; while refactoring our code we have introduced these variables as arrays too. It is safe to conclude that arrays get processed fatser than variables (in case of using VBA).

Also by putting our year input box closer to the end of or code in case of All Stocks Ananlysis we might have made it run through the entire datasheet before picking the year we need, while we did the exact opposite in case of refactoring our code - put the year input box in the very beginning of our code to make it run only for the year we need.

#**Summary

Based on our analysis we can conclude that replacing some variables with arrays saves some time while coding in VBA. In addition to that, putting input boxes in the very beginning of our code limits the ammount of data the code needs to run through, that saves us more time. While our datasheet was relatively small thus 5 times the speed, it might make a much bigger difference when dealing with larger amounts of data.
