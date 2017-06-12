---
title: In Operator (Microsoft Access SQL)
ms.prod: access
ms.assetid: ee4f1d71-82c4-3b0d-94b6-ad3f5a7608b8
ms.date: 06/08/2017
---


# In Operator (Microsoft Access SQL)

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Syntax](#sectionSection0)
[Remarks](#sectionSection1)
[ Example](#sectionSection2)


Determines whether the value of an expression is equal to any of several values in a specified list.

## Syntax
<a name="sectionSection0"> </a>

 _expr_ [ **Not** ] **In(** _value1, value2, …_ **)**


## Remarks
<a name="sectionSection1"> </a>

The  **In** operator syntax has these parts:



|**Part**|**Description**|
|:-----|:-----|
| _expr_|Expression identifying the field that contains the data you want to evaluate.|
| _value1_, _value2_|Expression or list of expressions against which you want to evaluate  _expr_.|
If  _expr_ is found in the list of values _,_ the **In** operator returns **True**; otherwise, it returns **False**. You can include the **Not** logical operator to evaluate the opposite condition (that is, whether _expr_ is not in the list of values).

For example, you can use  **In** to determine which orders are shipped to a set of specified regions:




```sql
SELECT * 
FROM Orders 
WHERE ShipRegion In ('Avon','Glos','Som')
```


## Example
<a name="sectionSection2"> </a>

The following example uses the Orders table in the Northwind.mdb database to create a query that includes all orders shipped to Lancashire and Essex and the dates shipped. 

This example calls the EnumFields procedure, which you can find in the SELECT statement example.




```vb
Sub InX() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
 
    ' Select records from the Orders table that 
    ' have a ShipRegion value of Lancashire or Essex. 
    Set rst = dbs.OpenRecordset("SELECT " _ 
        &; "CustomerID, ShippedDate FROM Orders " _ 
        &; "WHERE ShipRegion In " _ 
        &; "('Lancashire','Essex');") 
     
    ' Populate the Recordset. 
    rst.MoveLast 
     
    ' Call EnumFields to print the contents of 
    ' the Recordset. 
    EnumFields rst, 12 
 
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

