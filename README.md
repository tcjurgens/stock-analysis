# stock-analysis
Check out VBA_Challenge.xlsm to see the VBA and accurate results for this challenge working in action. 
## Project Overview
In this challenge/module we were helping our friend "Steve" analyze the performance 12 different green energy stocks for the years 2017 and 2018. We used Excel worksheets and VBA to analyze and compare Total Daily Volume and yearly Return in a code we built throughout the module.  In this challenge, we were tasked to refactor this code so that the VBA script ran faster.
## Results
![VBA_Challenge_2017](https://user-images.githubusercontent.com/86446641/125168531-ff8f4480-e173-11eb-9da2-b9dcc123292c.png)
<img width="424" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/86446641/125168539-0a49d980-e174-11eb-9655-793e54aee522.png">

Above are saved images of the pop-up messages which are shown after the refactored code (VBA_challenge.xlsm) is run. You can see that the run-times are much less than the original times (roughly .1 seconds when refactored compared to ~.5 of the original).

And for conveinence I will show the outputs of the refactored code below. The resulted outputs are expected. 
<img width="602" alt="Screen Shot 2021-07-10 at 11 49 42 AM" src="https://user-images.githubusercontent.com/86446641/125168729-ec30a900-e174-11eb-8bfe-0f916d2b4d70.png">
<img width="601" alt="Screen Shot 2021-07-10 at 11 49 59 AM" src="https://user-images.githubusercontent.com/86446641/125168746-f5ba1100-e174-11eb-9021-a55e16e0cbfa.png">

## Summary

### Advantages of Refactoring code
An obvious advantage of refactoring code is that when achieved-- it allows for code to be run much more efficiently. 
It seems that the purpose of this assignment was to refactor the code so that it may be both run both easily understood/programmed but especially more quickly run in VBA.  The time advantage gained by refactoring in this example is negligable, but if we were to analyze 100 different stocks over the last 50 years-- that time might really add up. 
### Disadvantages of Refactoring code
The Disadvantage seems somewhat obvious as well. Refactoring code is time consuming and surely not always easy. In this exqample: is it worth it for you to spend a few hours to make a code run faster? Probably not.  
Additionally, how can we know that our refactored code will be 'better' than the original--- how do we know we aren't wasting our time on this project? How will we know that the time spent to refactor will pay itself off?-- especially considering the results shown will be the same.. If it isn't broken, why fix it? 

These questions aren't easy to answer, and should be analyzed on a case by case basis.  In the example, the benefit gained by refactoring the script is negligable. However, other times it may be very advantageous.
