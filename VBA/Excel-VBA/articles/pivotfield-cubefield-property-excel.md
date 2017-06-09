---
title: PivotField.CubeField Property (Excel)
keywords: vbaxl10.chm240126
f1_keywords:
- vbaxl10.chm240126
ms.prod: excel
api_name:
- Excel.PivotField.CubeField
ms.assetid: d49d9454-6505-b892-a3c5-32c002326a31
ms.date: 06/08/2017
---


# PivotField.CubeField Property (Excel)

Returns the  **[CubeField](cubefield-object-excel.md)** object from which the specified PivotTable field is descended. Read-only.


## Syntax

 _expression_ . **CubeField**

 _expression_ A variable that represents a **PivotField** object.


## Example

This example creates a list of the cube field names for all the hierarchy fields in the first Online Analytical Processing (OLAP) -based PivotTable report on the first worksheet. This example assumes a PivotTable report exists in the first worksheet.


```vb
Sub UseCubeField() 
 
 Dim objNewSheet As Worksheet 
 Set objNewSheet = Worksheets.Add 
 objNewSheet.Activate 
 intRow = 1 
 
 For Each objPF in _ 
 Worksheets(1).PivotTables(1).PivotFields 
 If objPF.CubeField.CubeFieldType = xlHierarchy Then 
 objNewSheet.Cells(intRow, 1).Value = objPF.Name 
 intRow = intRow + 1 
 End If 
 Next objPF 
 
End Sub
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

