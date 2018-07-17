---
title: Min, Max Functions (Microsoft Access SQL)
keywords: jetsql40.chm5278826
f1_keywords:
- jetsql40.chm5278826
ms.prod: access
ms.assetid: 5ac77377-1f6a-7b4f-ecbb-5480bc5a3187
ms.date: 06/08/2017
---


# Min, Max Functions (Microsoft Access SQL)

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Syntax](#sectionSection0)
[Remarks](#sectionSection1)
[ Example](#sectionSection2)
[About the Contributors](#AboutContributors)


Return the minimum or maximum of a set of values contained in a specified field on a query.

## Syntax
<a name="sectionSection0"> </a>

 **Min(** _expr_ **)**

 **Max(** _expr_ **)**

The  _expr_ placeholder represents a string expression identifying the field that contains the data you want to evaluate or an expression that performs a calculation using the data in that field. Operands in _expr_ can include the name of a table field, a constant, or a function (which can be either intrinsic or user-defined but not one of the other SQL aggregate functions).


## Remarks
<a name="sectionSection1"> </a>

You can use  **Min** and **Max** to determine the smallest and largest values in a field based on the specified aggregation, or grouping. For example, you could use these functions to return the lowest and highest freight cost. If there is no aggregation specified, then the entire table is used.

You can use  **Min** and **Max** in a query expression and in the **SQL** property of a **QueryDef** object or when creating a **Recordset** object based on an SQL query.

 **Link provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community |[About the Contributors](#AboutContributors)


- [Record Order](http://www.utteraccess.com/wiki/index.php/Record_Order)
    

## Example
<a name="sectionSection2"> </a>

This example uses the Orders table to return the lowest and highest freight charges for orders shipped to the United Kingdom.

This example calls the EnumFields procedure, which you can find in the SELECT statement example.




```vb
Sub MinMaxX() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
     
    ' Return the lowest and highest freight charges for  
    ' orders shipped to the United Kingdom. 
    Set rst = dbs.OpenRecordset("SELECT " _  
        &; "Min(Freight) AS [Low Freight], " _ 
        &; "Max(Freight)AS [High Freight] " _ 
        &; "FROM Orders WHERE ShipCountry = 'UK';") 
     
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

