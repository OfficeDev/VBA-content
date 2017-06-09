---
title: Sum Function (Microsoft Access SQL)
keywords: jetsql40.chm5278828
f1_keywords:
- jetsql40.chm5278828
ms.prod: access
ms.assetid: 02498420-f177-521c-ef81-e2f7ea02b231
ms.date: 06/08/2017
---


# Sum Function (Microsoft Access SQL)

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Syntax](#sectionSection0)
[Remarks](#sectionSection1)
[ Example](#sectionSection2)


Returns the sum of a set of values contained in a specified field on a query.

## Syntax
<a name="sectionSection0"> </a>

 **Sum(** _expr_ **)**

The  _expr_ placeholder represents a string expression identifying the field that contains the numeric data you want to add or an expression that performs a calculation using the data in that field. Operands in _expr_ can include the name of a table field, a constant, or a function (which can be either intrinsic or user-defined but not one of the other SQL aggregate functions).


## Remarks
<a name="sectionSection1"> </a>

The  **Sum** function totals the values in a field. For example, you could use the **Sum** function to determine the total cost of freight charges.

The  **Sum** function ignores records that contain **Null** fields. The following example shows how you can calculate the sum of the products of UnitPrice and Quantity fields:




```sql
SELECT 
Sum(UnitPrice * Quantity) 
AS [Total Revenue] FROM [Order Details];
```

You can use the  **Sum** function in a query expression. You can also use this expression in the **SQL** property of a **QueryDef** object or when creating a **Recordset** based on an SQL query.


## Example
<a name="sectionSection2"> </a>

This example uses the Orders table to calculate the total sales for orders shipped to the United Kingdom.

This example calls the EnumFields procedure, which you can find in the SELECT statement example.




```vb
Sub SumX() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
 
    ' Calculate the total sales for orders shipped to 
    ' the United Kingdom.   
    Set rst = dbs.OpenRecordset("SELECT" _ 
        &; " Sum(UnitPrice*Quantity)" _ 
        &; " AS [Total UK Sales] FROM Orders" _ 
        &; " INNER JOIN [Order Details] ON" _ 
        &; " Orders.OrderID = [Order Details].OrderID" _ 
        &; " WHERE (ShipCountry = 'UK');") 
 
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

