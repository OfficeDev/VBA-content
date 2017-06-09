---
title: StDev, StDevP Functions (Microsoft Access SQL)
keywords: jetsql40.chm5278827
f1_keywords:
- jetsql40.chm5278827
ms.prod: access
ms.assetid: 880875e9-75bc-da59-5554-810e15ce4d54
ms.date: 06/08/2017
---


# StDev, StDevP Functions (Microsoft Access SQL)

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Syntax](#sectionSection0)
[Remarks](#sectionSection1)
[ Example](#sectionSection2)


Return estimates of the standard deviation for a population or a population sample represented as a set of values contained in a specified field on a query.

## Syntax
<a name="sectionSection0"> </a>

 **StDev(** _expr_ **)**

 **StDevP(** _expr_ **)**

The  _expr_ placeholder represents a string expression identifying the field that contains the numeric data you want to evaluate or an expression that performs a calculation using the data in that field. Operands in _expr_ can include the name of a table field, a constant, or a function (which can be either intrinsic or user-defined but not one of the other SQL aggregate functions).


## Remarks
<a name="sectionSection1"> </a>

The  **StDevP** function evaluates a population, and the **StDev** function evaluates a population sample.

If the underlying query contains fewer than two records (or no records, for the  **StDevP** function), these functions return a **Null** value (which indicates that a standard deviation cannot be calculated).

You can use the  **StDev** and **StDevP** functions in a query expression. You can also use this expression in the **SQL** property of a **QueryDef** object or when creating a **Recordset** object based on an SQL query.


## Example
<a name="sectionSection2"> </a>

This example uses the Orders table to estimate the standard deviation of the freight charges for orders shipped to the United Kingdom.

This example calls the EnumFields procedure, which you can find in the SELECT statement example.




```vb
Sub StDevX() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
 
    ' Calculate the standard deviation of the freight 
    ' charges for orders shipped to the United Kingdom. 
    Set rst = dbs.OpenRecordset("SELECT " _ 
        &; "StDev(Freight) " _ 
        &; "AS [Freight Deviation] FROM Orders " _ 
        &; "WHERE ShipCountry = 'UK';") 
 
    ' Populate the Recordset. 
    rst.MoveLast 
     
    ' Call EnumFields to print the contents of the  
    ' Recordset. Pass the Recordset object and desired 
    ' field width. 
    EnumFields rst, 15 
     
    Debug.Print 
     
    Set rst = dbs.OpenRecordset("SELECT " _ 
        &; "StDevP(Freight) " _ 
        &; "AS [Freight DevP] FROM Orders " _ 
        &; "WHERE ShipCountry = 'UK';") 
 
    ' Populate the Recordset. 
    rst.MoveLast 
 
    ' Call EnumFields to print the contents of the  
    ' Recordset. Pass the Recordset object and desired 
    ' field width. 
    EnumFields rst, 15 
 
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

