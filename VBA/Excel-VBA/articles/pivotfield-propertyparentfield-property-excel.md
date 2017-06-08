---
title: PivotField.PropertyParentField Property (Excel)
keywords: vbaxl10.chm240132
f1_keywords:
- vbaxl10.chm240132
ms.prod: excel
api_name:
- Excel.PivotField.PropertyParentField
ms.assetid: 98b4f7e5-0e41-19ea-b6bb-d938e2756f97
ms.date: 06/08/2017
---


# PivotField.PropertyParentField Property (Excel)

Returns a  **PivotField** object representing the field to which the properties in this field pertain.


## Syntax

 _expression_ . **PropertyParentField**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

Valid only for fields that are member property fields.

If the  **[IsMemberProperty](pivotfield-ismemberproperty-property-excel.md)** property is **False** , using the **PropertyParentField** property will return a run-time error.


## Example

This example determines if there are member properties in the fourth field and, if there are, which fields the properties pertain to. Depending on the findings, Excel notifies the user. This example assumes that a PivotTable exists on the active worksheet and that it is based on an Online Analytical Processing (OLAP) data source.


```vb
Sub CheckParentField() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.PivotFields(4) 
 
 ' Check for member properties and notify user. 
 If pvtField.IsMemberProperty = False Then 
 MsgBox "No member properties present." 
 Else 
 MsgBox "The parent field of the members is: " &; _ 
 pvtField.PropertyParentField 
 End If 
 
End Sub
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

