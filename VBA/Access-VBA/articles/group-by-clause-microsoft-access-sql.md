---
title: GROUP BY Clause (Microsoft Access SQL)
keywords: jetsql40.chm5277569
f1_keywords:
- jetsql40.chm5277569
ms.prod: access
ms.assetid: fe7d5e27-a47a-1229-232c-cf6a0cbad761
ms.date: 06/08/2017
---


# GROUP BY Clause (Microsoft Access SQL)

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Syntax](#sectionSection0)
[Remarks](#sectionSection1)
[ Example](#sectionSection2)


Combines records with identical values in the specified field list into a single record. A summary value is created for each record if you include an [SQL aggregate function](http://msdn.microsoft.com/library/8866cd71-0216-25b4-6a6a-02cb7acad9a2%28Office.15%29.aspx), such as  **Sum** or **Count**, in the[SELECT](http://msdn.microsoft.com/library/a5c9da94-5f9e-0fc0-767a-4117f38a5ef3%28Office.15%29.aspx) statement.

## Syntax
<a name="sectionSection0"> </a>

SELECT  _fieldlist_ FROM _table_ WHERE _criteria_ [GROUP BY _groupfieldlist_ ]

A SELECT statement containing a GROUP BY clause has these parts:



|**Part**|**Description**|
|:-----|:-----|
| _fieldlist_|The name of the field or fields to be retrieved along with any field-name aliases, SQL aggregate functions, selection predicates ([ALL, DISTINCT, DISTINCTROW, or TOP](all-distinct-distinctrow-top-predicates-microsoft-access-sql.md)), or other SELECT statement options.|
| _table_|The name of the table from which records are retrieved. For more information, see the [FROM](from-clause-microsoft-access-sql.md) clause.|
| _criteria_|Selection criteria. If the statement includes a [WHERE](where-clause-microsoft-access-sql.md) clause, the Microsoft Access database engine groups values after applying the WHERE conditions to the records.|
| _groupfieldlist_|The names of up to 10 fields used to group records. The order of the field names in  _groupfieldlist_ determines the grouping levels from the highest to the lowest level of grouping.|

## Remarks
<a name="sectionSection1"> </a>

GROUP BY is optional.

Summary values are omitted if there is no SQL aggregate function in the SELECT statement.

 **Null** values in GROUP BY fields are grouped and are not omitted. However, **Null** values are not evaluated in any SQL aggregate function.

Use the WHERE clause to exclude rows you do not want grouped, and use the [HAVING](having-clause-microsoft-access-sql.md) clause to filter records after they have been grouped.

Unless it contains Memo or OLE Object data, a field in the GROUP BY field list can refer to any field in any table listed in the FROM clause, even if the field is not included in the SELECT statement, provided the SELECT statement includes at least one SQL aggregate function. The Microsoft® Jet database engine cannot group on Memo or OLE Object fields.

All fields in the SELECT field list must either be included in the GROUP BY clause or be included as arguments to an SQL aggregate function.


## Example
<a name="sectionSection2"> </a>

This example creates a list of unique job titles and the number of employees with each title.

This example calls the EnumFields procedure, which you can find in the SELECT statement example.




```vb
Sub GroupByX1() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
 
    ' For each title, count the number of employees  
    ' with that title.  
    Set rst = dbs.OpenRecordset("SELECT Title, " _ 
        &; "Count([Title]) AS Tally " _ 
        &; "FROM Employees GROUP BY Title;") 
     
    ' Populate the Recordset. 
    rst.MoveLast 
     
    ' Call EnumFields to print the contents of the  
    ' Recordset. Pass the Recordset object and desired 
    ' field width. 
    EnumFields rst, 25 
 
    dbs.Close 
 
End Sub 

```

For each unique job title, this example calculates the number of employees in Washington who have that title.




```vb
Sub GroupByX2() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
     
    ' For each title, count the number of employees  
    ' with that title. Only include employees in the  
    ' Washington region. 
    Set rst = dbs.OpenRecordset("SELECT Title, " _ 
        &; "Count(Title) AS Tally " _ 
        &; "FROM Employees WHERE Region = 'WA' " _ 
        &; "GROUP BY Title;") 
     
    ' Populate the Recordset. 
    rst.MoveLast 
 
    ' Call EnumFields to print the contents of the  
    ' Recordset. Pass the Recordset object and desired 
    ' field width. 
    EnumFields rst, 25 
 
    dbs.Close 
 
End Sub 

```

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

