# excel-automated-report-project
## Project overview
- Using excel to create sales reports that utilizes Power Query and VBA Macros to automate report genaration
- Creating sales report by visualizing data using PivotTables and PivotChart
- Electronic store sales data are utilized
## Objective
1. Create steps that is required to turn raw data into usable insight
2. Create a summary on the findings from the sales data
3. Ensure the repeating tasks are automated from month-to-month report
4. Insights from the visualization should help with sales strategies

## Objective 1: Create steps that is required to turn raw data into usable insight
### Step 1: Getting data and uploading it into excel using power query
- The dataset was downloaded from https://www.kaggle.com/sinjoysaha/sales-analysis-dataset in csv file 

![image](https://user-images.githubusercontent.com/53467152/155058763-58a6c049-26a8-4a1a-895a-1ec6f4fb2b63.png)
-Figure 1: The opened csv files downloaded from the link ( dateset for January 2019)

- The dataset then imported into excel workbook using Power Query
- In excel: Data ribbon > Get and Transform data > From Text/CSV and look for the file and load 

![image](https://user-images.githubusercontent.com/53467152/155058918-a9c0dc70-89ef-4bc0-9bb9-a7dc6b420acc.png)
-Figure 2: The loaded dataset into new workbook as table

### Step 2: Apply data cleaning and transformation to the dataset
- Using power query to clean the dataset
- In excel: Query ribbon > Edit or Launch Power Query Editor from the Get data dropdown menu from Data ribbon
- Apply data cleaning steps such as removing blanks, incorrect data entry and data format
- The date format are not the same and some cannot be detected (for this step only it will be address not using the power editor but by using the text-to-column function in Data ribbon and to detect it as MDY date, and for automation purpose it will be recorded in the macros in the following steps)
- Split column to provide more insight (i.g. splitting the address column to obtain city or state)
- Suggestion: If needed, since the state is in abbrieviation, VLOOKUP can be used to switch it into full name
- Adding appropiate column (i.g to create sales column multiple quantity ordered and price)

![image](https://user-images.githubusercontent.com/53467152/155060294-9fadbed0-8096-470e-9ce2-e0449374a978.png)
- Figure 3: The cleaned dataset
## Objective 2: Create a summary on the findings from the sales data
### Step 3: Create PivotTables and Pivot Charts while recording it in Macros
- To record the steps applied just go to Developer ribbon> Code> Record Macro and enter the appriota name and description
- Important: Save it into Personal Macro Workbook to apply it in different workbook later
- using PivotTables to create meaningful relationship between attributes

![image](https://user-images.githubusercontent.com/53467152/155061781-9d354109-3ce6-49d8-811c-e4faf89afd22.png)
- Figure 4: PivotTables examples

- using PivotCharts to visualize for more insight

![image](https://user-images.githubusercontent.com/53467152/155063290-795bae97-c6d1-45b3-804b-dd39df3f7b24.png)
-Figure 5: January Report

- Suggestion: Creating geographical visaulization helps to get insight on how each branch are performing

## Objective 3:  Ensure the repeating tasks are automated from month-to-month report
### Step 5: Using Power Query and VBA Macros to automate report
- The automation reduces the steps required into two major step: cleaning & transformation and Visualization
- Once the loading steps for the new dataset are done, to apply it the steps in the Power Query editor for the previous dataset can be copied.
- In Power Query editor: View > Advanced editor and copied all the syntax except the source and paste it in the editor of the new workbook for new data
- The steps should work given the dataset attributes are the same and have the same header name


![image](https://user-images.githubusercontent.com/53467152/155064999-a7c1cb3e-bfdd-4862-ac51-7ce3e05ee9e7.png)
- Figure 6: Cleaned February dataset

- For the visualization part which is the step where PivotTable and chart are made just apply the macros recorded in step 3
- Go to Developer ribbon > Code > Macros and select the saved macros earlier
- Important: to make the code work it is important that the references in the code are edited to match with the new one (i.e the table name)
- other things should work if nothing else is changed

![image](https://user-images.githubusercontent.com/53467152/155065868-1d9f9b1f-1027-4557-aa37-53c65265666c.png)
-Figure 7: February report

- Repeat steps on new dataset that has similar structure whenever required
- For the steps, it should work on all the dataset in the link given (the link contain monthly dataset for the year 2019)

![image](https://user-images.githubusercontent.com/53467152/155068378-d2d26d60-84f8-450a-b587-3296684788ce.png)
Figure 8: March report

- For this project, January, February and March dataset are the only one used
- Figure 5,7 and 8 are the reports from the said datasets

## Objective 4: Insights from the visualization should help with sales strategies
## Conclusion
- The top perfoming products that generate huge sales are MacBook Pro Laptop, iPhone, Google Phone and ThinkPad Laptop. This trend was consistent from January to March 2019.
- In term of quantity ordered, batteries and USB cables are the products that was sold in higher amount.
- The peak selling hour start around 10 am with a slight dip around 12 pm and rise againat 3pm  until it drops around 8 pm. This information should help to strategize personnel scheduling and stocking timing to better accomodate the prime selling hours. Naturally, more personnel are needed between 10 am to 12 pm and 3 pm to 8 pm. Stocking processes should be done before 10 am and restocking can be done around 12 pm to 3 pm.
- The quantity ordered volume by day in different months does not peak on the same day. Employee need to be trained to adapt to the changes in work-demand.



