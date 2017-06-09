---
title: Count Function (Microsoft Access SQL)
keywords: jetsql40.chm5278824
f1_keywords:
- jetsql40.chm5278824
ms.prod: access
ms.assetid: 01743d33-d7de-12b5-eb0f-eb775b0bcffd
ms.date: 06/08/2017
---


# Count Function (Microsoft Access SQL)

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Syntax](#sectionSection0)
[Remarks](#sectionSection1)
[ Example](#sectionSection2)


Calculates the number of records returned by a query.

## Syntax
<a name="sectionSection0"> </a>

 **Count(** _expr_ **)**

The  _expr_ placeholder represents a string expression identifying the field that contains the data you want to count or an expression that performs a calculation using the data in the field. Operands in _expr_ can include the name of a table field or function (which can be either intrinsic or user-defined but not other SQL aggregate functions). You can count any kind of data, including text.


## Remarks
<a name="sectionSection1"> </a>

You can use  **Count** to count the number of records in an underlying query. For example, you could use **Count** to count the number of orders shipped to a particular country or region.

Although  _expr_ can perform a calculation on a field, **Count** simply tallies the number of records. It does not matter what values are stored in the records.

The  **Count** function does not count records that have **Null** fields unless _expr_ is the asterisk (*) wildcard character. If you use an asterisk, **Count** calculates the total number of records, including those that contain **Null** fields. **Count(** * **)** is considerably faster than **Count(** [ _Column Name_ ] **)**. Do not enclose the asterisk in quotation marks (' '). The following example calculates the number of records in the Orders table:




```sql
SELECT Count(*) 
AS TotalOrders FROM Orders;
```

If  _expr_ identifies multiple fields, the **Count** function counts a record only if at least one of the fields is not **Null**. If all of the specified fields are **Null**, the record is not counted. Separate the field names with an ampersand (&;). The following example shows how you can limit the count to records in which either ShippedDate or Freight is not **Null**:




```sql
SELECT 
Count('ShippedDate &; Freight') 
AS [Not Null] FROM Orders;
```

You can use  **Count** in a query expression. You can also use this expression in the **SQL** property of a **QueryDef** object or when creating a **Recordset** object based on an SQL query.


## Example
<a name="sectionSection2"> </a>

This example uses the Orders table to calculate the number of orders shipped to the United Kingdom.

This example calls the EnumFields procedure, which you can find in the SELECT statement example.




```vb
Sub CountX() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
    
    ' Calculate the number of orders shipped  
    ' to the United Kingdom. 
    Set rst = dbs.OpenRecordset("SELECT" _ 
        &; " Count (ShipCountry)" _ 
        &; " AS [UK Orders] FROM Orders" _ 
        &; " WHERE ShipCountry = 'UK';") 
     
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

