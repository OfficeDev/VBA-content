---
title: CalculatedMember Object (Excel)
keywords: vbaxl10.chm685072
f1_keywords:
- vbaxl10.chm685072
ms.prod: excel
api_name:
- Excel.CalculatedMember
ms.assetid: 07a1f8df-107e-a5fd-3d15-dfc92916c4c6
ms.date: 06/08/2017
---


# CalculatedMember Object (Excel)

Represents the calculated fields, calculated items, and named sets for PivotTables with Online Analytical Processing (OLAP) data sources.


## Remarks

Use the  **[Add](calculatedmembers-add-method-excel.md)** method or the[Item](calculatedmembers-item-property-excel.md) property of the **[CalculatedMembers](calculatedmembers-object-excel.md)** collection to return a **CalculatedMember** object.

With a  **CalculatedMember** object you can check the validity of a calculated field or item in a PivotTable using the **[IsValid](calculatedmember-isvalid-property-excel.md)** property.




 **Note**   The **IsValid** property will return **True** if the PivotTable is not currently connected to the data source. Use the **[MakeConnection](pivotcache-makeconnection-method-excel.md)** method before testing the **IsValid** property.


## Example

The following example notifies the user if the calculated member is valid or not. This example assumes a PivotTable exists on the active worksheet that contains either a valid or invalid calculated member.


```vb
Sub CheckValidity() 
 
 Dim pvtTable As PivotTable 
 Dim pvtCache As PivotCache 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 ' Handle run-time error if external source is not an OLEDB data source. 
 On Error GoTo Not_OLEDB 
 
 ' Check connection setting and make connection if necessary. 
 If pvtCache.IsConnected = False Then 
 pvtCache.MakeConnection 
 End If 
 
 ' Check if calculated member is valid. 
 If pvtTable.CalculatedMembers.Item(1).IsValid = True Then 
 MsgBox "The calculated member is valid." 
 Else 
 MsgBox "The calculated member is not valid." 
 End If 
 
End Sub
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

