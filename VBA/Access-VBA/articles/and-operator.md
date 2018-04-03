---
title: ..And Operator
keywords: jetsql40.chm5277585
f1_keywords:
- jetsql40.chm5277585
ms.prod: access
ms.assetid: 33a49af8-25f4-b107-e0e2-17c90d80c66a
ms.date: 06/08/2017
---


# ..And Operator

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Syntax](#sectionSection0)
[Remarks](#sectionSection1)
[Example](#sectionSection2)


Determines whether the value of an expression falls within a specified range of values. You can use this operator within SQL statements.

## Syntax
<a name="sectionSection0"> </a>

 _expr_ [ **Not** ] **Between** _value1_ **And** _value2_

The  **Between…And** operator syntax has these parts:



|**Part**|**Description**|
|:-----|:-----|
| _expr_|Expression identifying the field that contains the data you want to evaluate.|
| _value1_, _value2_|Expressions against which you want to evaluate  _expr_.|

## Remarks
<a name="sectionSection1"> </a>

If the value of  _expr_ is between _value1_ and _value2_ (inclusive), the **Between…And** operator returns **True**; otherwise, it returns **False**. You can include the **Not** logical operator to evaluate the opposite condition (that is, whether _expr_ lies outside the range defined by _value1_ and _value2_ ).

You might use  **Between…And** to determine whether the value of a field falls within a specified numeric range. The following example determines whether an order was shipped to a location within a range of postal codes. If the postal code is between 98101 and 98199, the **IIf** function returns "Local". Otherwise, it returns "Nonlocal".




```vb
SELECT IIf(PostalCode Between 98101 And 98199, "Local", "Nonlocal")
FROM Publishers;
```

If  _expr_, _value1_, or _value2_ is **Null**, **Between…And** returns a **Null** value.

Because wildcard characters, such as *, are treated as literals, you cannot use them with the  **Between…And** operator. For example, you cannot use 980* and 989* to find all postal codes that start with 980 to 989. Instead, you have two alternatives for accomplishing this. You can add an expression to the query that takes the left three characters of the text field and use **Between…And** on those characters. Or you can pad the high and low values with extra characters — in this case, 98000 to 98999, or 98000 to 98999 - 9999 if using extended postal codes. (You must omit the - 0000 from the low values because otherwise 98000 is dropped if some postal codes have extended sections and others do not.)


## Example
<a name="sectionSection2"> </a>

This example lists the name and contact of every customer who placed an order in the second quarter of 1995.

This example calls the EnumFields procedure, which you can find in the SELECT statement example.




```vb
Sub SubQueryX() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
    
    ' List the name and contact of every customer  
    ' who placed an order in the second quarter of 
    ' 1995. 
 
    Set rst = dbs.OpenRecordset("SELECT ContactName," _ 
        &; " CompanyName, ContactTitle, Phone" _ 
        &; " FROM Customers" _ 
        &; " WHERE CustomerID" _ 
        &; " IN (SELECT CustomerID FROM Orders" _ 
        &; " WHERE OrderDate Between #04/1/95#" _ 
        &; " And #07/1/95#);") 
     
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

