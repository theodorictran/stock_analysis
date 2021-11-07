# stock_analysis

## Overview of the project
Steve Green was interested in a variety of green energy stocks and wanted to see their performance. Passionate about green energy, Steve's parents doubled down and invested all their money into Daqo New Energy Corporation. In efforts to diversify his parents' portfolio and reduce risk, Steve created an Excel file with numerous other green stocks that he wants analyzed. Using Excel and VBA, I was asked to analyze the different stocks and automate this process for Steve.

## Results
![2017_Output](https://user-images.githubusercontent.com/91519293/140657845-2cee6bb9-40ee-4cc0-b0df-2cb112918c0b.png)
![2018_Output](https://user-images.githubusercontent.com/91519293/140657852-ff5aaf05-b8ea-45b6-adc2-982a016fa2b7.png)

Based on the images above, Steve should look into investing ENPH: Enphase Energy Inc and RUN: Sunrun Inc, for they are the only two stocks to give a positive return from 2017 to 2018. To achieve these results, I used Visual Basic for Applications or VBA, a programming language that interacts with Excel. I wrote a script that takes year input from the user and goes to the according sheet in Excel that loops through all the tickers and gets the total volume, starting price, and ending price for the current ticker which then outputs the associated data for the current ticker on the sheet for the user.

![Standard_Code](https://user-images.githubusercontent.com/91519293/140658325-966134b3-3afb-4552-961c-ca084b38f858.png)
![Standard_Code_2](https://user-images.githubusercontent.com/91519293/140658329-8a44863e-b51d-4325-bb7c-35d6327e5991.png)

From this written code, I created three buttons for Steve to press. One button would ask for the year and perform the calculations. Another button would format the table for better visibility and clarity. The third button would clear the sheet clean.

![Workspace](https://user-images.githubusercontent.com/91519293/140658413-de9f5735-8e22-4347-98b0-d00cc168e72c.png)

The major challenge of my work is the performance in time of my first script. Running the analysis script on the year 2017 took about 0.6293945 seconds and running the same script on the year 2018 took 0.6279297 seconds.

![Standard_Time_2017](https://user-images.githubusercontent.com/91519293/140658530-8ff24b61-ff1d-46d2-84ec-71cb688258a6.png)
![Standard_Time_2018](https://user-images.githubusercontent.com/91519293/140658532-664aff2c-1bfb-4798-a1c3-0d1b8b79fc22.png)

In an effort to optimize my code and simplify the process for Steve, I successfully refactored my code.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/91519293/140658587-9fb1532a-dea8-4a7e-a92b-fd8d619c8061.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/91519293/140658588-979ef3c4-bfa5-4ed8-a830-a7239704475f.png)

## Summary

### What are the advantages or disadvantages of refactoring code?
The advantages of refactoring code are the performance increase, simplification of logic, add new features to existing code, and automation. For example, the time-savings from refactoring code could be money saving for a large corporation. Another example, a large script with thousands of lines could be refactored to add comments and change the structure in which helps other analysts or engineers read the code. Although refactoring code has its benefits, the time spent refactoring could be better used elsewhere and the potential of corrupting the existing code is a still a threat. 

### How do these pros and cons apply to refactoring the original VBA script?
In the case of this analysis, refactoring the original VBA script saved time. What originally took about 0.63 seconds takes about 900 hundreths of a second. In addition, for Steve, he would have to go through at least 2 out of 3 buttons that I created for him. One of which would take Steve's input of the year of interest and return the all the tickers and their total daily volume, and return. Another button that Steve would press would format the table to show which stocks were positively performing and negatively performing. Refactoring the code allowed me to integrate the two buttons into one function for Steve which would take overall less time for the computer process. 
