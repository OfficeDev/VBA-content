---
title: PivotField.ChildField Property (Excel)
keywords: vbaxl10.chm240075
f1_keywords:
- vbaxl10.chm240075
ms.prod: excel
api_name:
- Excel.PivotField.ChildField
ms.assetid: 97e246de-208f-5932-a553-525da17b0d4d
ms.date: 06/08/2017
---


# PivotField.ChildField Property (Excel)

Returns a  **[PivotField](pivotfield-object-excel.md)** object that represents the child field for the specified field (if the field is grouped and has a child field). Read-only.


## Syntax

 _expression_ . **ChildField**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

If the specified field has no child field, this property causes an error.

This property is not available for OLAP data sources.


## Example

This example displays the name of the child field for the field named "REGION2."


```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
MsgBox "The name of the child field is " &; _ 
 pvtTable.PivotFields("REGION2").ChildField.Name
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

