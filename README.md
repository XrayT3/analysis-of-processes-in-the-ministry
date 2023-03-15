# Analysis of processes in the ministry
How many people should be recruited into the new department?

## Probelm
You are a data analyst for a ministry and have been asked to analyze data from a process evaluation of workers at the local level.
The timeframe for an office to respond to a benefit application is typically 30 days. Until now, the Department has mainly monitored compliance with this deadline. 
It has never centrally addressed whether and how benefits are administered procedurally. This has been left to the local level.

**Goal:** In a month we will open recruitment for a team to handle the process centrally. Need to find out how many people will need to be recruited before the start?

**Outputs:**
- Code for data transformation
- Power BI data report and short presentation

## Solution

In the beginning, we have 69 different tables, but fortunately, they have the same format. 
So we can transform them into one big table. It will help us to work with all data at once. 

Let's figure out what columns we need and what information is important to analyze.
First of all, we need **processes**, **activities**, and **name of processes**, because it is the main part of the department job. 
Then we need additional information. They are type of submission, type of application, and role.
For the calculation time difficulty of process we add **time**.
And at the end, we have to add a source of data(region, name of file), so our data can be reproduced.

Now we can use the Phyton script(main.py) to transform data from 69 files into one table(table.xlsx).

Having all data, we can create a Power BI report.

https://user-images.githubusercontent.com/25695606/225366905-27069971-4509-4181-b56c-79c87257bb6c.mp4

