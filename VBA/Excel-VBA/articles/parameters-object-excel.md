---
title: Parameters Object (Excel)
keywords: vbaxl10.chm524072
f1_keywords:
- vbaxl10.chm524072
ms.prod: excel
api_name:
- Excel.Parameters
ms.assetid: d67147f1-d587-a9e4-ed8e-8a1140e8a868
ms.date: 06/08/2017
---


# Parameters Object (Excel)

A collection of  **[Parameter](parameter-object-excel.md)** objects for the specified query table.


## Remarks

 Each **Parameter** object represents a single query parameter. Every query table contains a **Parameters** collection, but the collection is empty unless the query table is using a parameter query.

You cannot use the  **[Add](parameters-add-method-excel.md)** method on a URL connection query table. For URL connection query tables, Microsoft Excel creates the parameters based on the **[Connection](querytable-connection-property-excel.md)** and **[PostText](querytable-posttext-property-excel.md)** properties.


## Example

Use the  **Parameters** property to return the **Parameters** collection. The following example displays the number of parameters in query table one.


```vb
MsgBox Workbooks(1).ActiveSheet.QueryTables(1).Parameters.Count
```

Use the  **Add** method to create a new parameter for a query table. The following example changes the SQL statement for query table one. The clause "(city=?)" indicates that the query is a parameter query, and the value of city is set to the constant "Oakland."




```sql
Set qt = Sheets("sheet1").QueryTables(1) 
qt.Sql = "SELECT * FROM authors WHERE (city=?)" 
Set param1 = qt.Parameters.Add("City Parameter", _ 
 xlParamTypeVarChar) 
param1.SetParam xlConstant, "Oakland" 
qt.Refresh
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


