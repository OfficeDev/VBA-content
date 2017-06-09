---
title: Chart.PivotLayout Property (Excel)
keywords: vbaxl10.chm149165
f1_keywords:
- vbaxl10.chm149165
ms.prod: excel
api_name:
- Excel.Chart.PivotLayout
ms.assetid: b621dc49-5321-5426-35cc-386cac251920
ms.date: 06/08/2017
---


# Chart.PivotLayout Property (Excel)

Returns a  **[PivotLayout](pivotlayout-object-excel.md)** object that represents the placement of fields in a PivotTable report and the placement of axes in a PivotChart report. Read-only.


## Syntax

 _expression_ . **PivotLayout**

 _expression_ A variable that represents a **Chart** object.


## Remarks

If the chart you specify isn't a PivotChart report, the value of this property is  **Nothing** .


## Example

This example creates a list of all the PivotTable field names used in the first PivotChart report.


```vb
Set objNewSheet = Worksheets.Add 
objNewSheet.Activate 
intRow = 1 
For Each objPF In _ 
 Charts("Chart1").PivotLayout.PivotFields 
 objNewSheet.Cells(intRow, 1).Value = objPF.Caption 
 intRow = intRow + 1 
Next objPF
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

