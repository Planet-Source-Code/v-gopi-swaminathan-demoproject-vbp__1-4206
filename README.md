<div align="center">

## demoproject\.vbp


</div>

### Description

The project has two forms. Both displays charts using mschart control for a

data from access table. One form displays a chart(2D BAR) for the first record of the table.

The second form displays a chart(2D BAR) for a set of records.

1.Both the charts displays legends.

2.Both the charts displays titles for X and Y axis.

3.Both the charts accepts a title to the chart at the desired location.

4.Both the charts displays data values at the desired location above the

data points of each item.

5.Above all, the second chart(in the second form)displays bars in the charts

dynamically as and when records are entered. i.e., if there are five records,

five sets of data series will be automatically displayed.

6. Column labels and row labels are named in the code itself.

7. Filling colors to each data series is mentioned in the code.

To summarize,

a customized graphical representation of a set of data from an access table

can be created by any user.
 
### More Info
 
We require an access database where a table has to be created.

In my example, i have created an access database called c:\demo\demoproject.mdb.

Note: In my code the path is C:\demo. So, please mention the same path or change

path to your desired directory in the code.

The table name is demoproject. In my example, I am displaying a bar chart for

expenditure on elections by four countries.(Sample only).

I have created five fields namely USA, JAPAN, GERMANY ,INDIA and YEAR.

Four records are created to represent data for the years 1985,1990,1995 and 1999.

Example: USA JAPAN GERMANY INDIA YEAR

200 140  120   80  1985 This is the first record.

Set the

The user should know how to define a database and record set in the code.

Then, the basic property of like enable,visible should be familiar.

Array concepts should be familiar

My code returns two charts. In the form frmdemo1_1, the chart displays a

record for a particular record. In the second form frmdemo1_2, the chart

displays all the records of a table with titles,legends,colours,datavalues etc.,

1. Check the path of the database in the source code. Otherwise error message appear.

2. In MS chart property, blank the row labels and column labels property as code takes

care of it.


<span>             |<span>
---                |---
**Submitted On**   |1999-10-28 17:10:48
**By**             |[V\.Gopi Swaminathan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/v-gopi-swaminathan.md)
**Level**          |Unknown
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD1521\.zip](https://github.com/Planet-Source-Code/v-gopi-swaminathan-demoproject-vbp__1-4206/archive/master.zip)








