---
title: First, Last Functions (Microsoft Access SQL)
keywords: jetsql40.chm5278825
f1_keywords:
- jetsql40.chm5278825
ms.prod: access
ms.assetid: 8ea0d390-bb37-003b-fb6c-e15bf2a50718
ms.date: 06/08/2017
---


# First, Last Functions (Microsoft Access SQL)

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Syntax](#sectionSection0)
[Remarks](#sectionSection1)
[ Example](#sectionSection2)
[About the Contributors](#AboutContributors)


Return a field value from the first or last record in the result set returned by a query.

## Syntax
<a name="sectionSection0"> </a>

 **First(** _expr_ **)**

 **Last(** _expr_ **)**

The  _expr_ placeholder represents a string expression identifying the field that contains the data you want to use or an expression that performs a calculation using the data in that field. Operands in _expr_ can include the name of a table field, a constant, or a function (which can be either intrinsic or user-defined but not one of the other SQL aggregate functions).


## Remarks
<a name="sectionSection1"> </a>

The  **First** and **Last** functions are analogous to the **MoveFirst** and **MoveLast** methods of a DAO Recordset object. They simply return the value of a specified field in the first or last record, respectively, of the result set returned by a query. Because records are usually returned in no particular order (unless the query includes an[ORDER BY](order-by-clause-microsoft-access-sql.md) clause), the records returned by these functions will be arbitrary.

 **Link provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community |[About the Contributors](#AboutContributors)


- [Record Order](http://www.utteraccess.com/wiki/index.php/Record_Order)
    

## Example
<a name="sectionSection2"> </a>

This example uses the Employees table to return the values from the LastName field of the first and last records returned from the table.

This example calls the EnumFields procedure, which you can find in the SELECT statement example.




```vb
Sub FirstLastX1() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
     
    ' Return the values from the LastName field of the  
    ' first and last records returned from the table. 
    Set rst = dbs.OpenRecordset("SELECT " _ 
        &; "First(LastName) as First, " _ 
        &; "Last(LastName) as Last FROM Employees;") 
     
    ' Populate the Recordset. 
    rst.MoveLast 
     
    ' Call EnumFields to print the contents of the  
    ' Recordset. Pass the Recordset object and desired 
    ' field width. 
    EnumFields rst, 12 
 
    dbs.Close 
 
End Sub 

```

The next example compares using the  **First** and **Last** functions with simply using the **Min** and **Max** functions to find the earliest and latest birth dates of Employees.




```vb
Sub FirstLastX2() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
     
    ' Find the earliest and latest birth dates of 
    ' Employees. 
    Set rst = dbs.OpenRecordset("SELECT " _ 
        &; "First(BirthDate) as FirstBD, " _ 
        &; "Last(BirthDate) as LastBD FROM Employees;") 
     
    ' Populate the Recordset. 
    rst.MoveLast 
     
    ' Call EnumFields to print the contents of the  
    ' Recordset. Pass the Recordset object and desired 
    ' field width. 
    EnumFields rst, 12 
     
    Debug.Print 
 
    ' Find the earliest and latest birth dates of 
    ' Employees. 
    Set rst = dbs.OpenRecordset("SELECT " _ 
        &; "Min(BirthDate) as MinBD," _ 
        &; "Max(BirthDate) as MaxBD FROM Employees;") 
     
    ' Populate the Recordset. 
    rst.MoveLast 
     
    ' Call EnumFields to print the contents of the  
    ' Recordset. Pass the Recordset object and desired 
    ' field width. 
    EnumFields rst, 12 
 
    dbs.Close 
 
End Sub 

```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

