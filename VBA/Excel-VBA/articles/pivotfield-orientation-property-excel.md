---
title: PivotField.Orientation Property (Excel)
keywords: vbaxl10.chm240087
f1_keywords:
- vbaxl10.chm240087
ms.prod: excel
api_name:
- Excel.PivotField.Orientation
ms.assetid: 1b3e0867-3a44-a908-ef1b-90ab21653ab9
ms.date: 06/08/2017
---


# PivotField.Orientation Property (Excel)

Returns or sets a  **[XlPivotFieldOrientation](xlpivotfieldorientation-enumeration-excel.md)** value that represents the location of the field in the specified PivotTable report.


## Syntax

 _expression_ . **Orientation**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

For OLAP data sources, setting this property for one field in a hierarchy sets the orientation for the other fields in the same hierarchy. Dimension fields can only be oriented in the row, column, and page field areas of the PivotTable report. Measure fields can only be oriented in the data area. Setting a hierarchy or data field to  **xlHidden** removes the hierarchy or field from the PivotTable report.


## Example

This example displays the orientation for the ORDER_DATE field.


```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
Set pvtField = pvtTable.PivotFields("ORDER_DATE") 
Select Case pvtField.Orientation 
 Case xlHidden 
 MsgBox "Hidden field" 
 Case xlRowField 
 MsgBox "Row field" 
 Case xlColumnField 
 MsgBox "Column field" 
 Case xlPageField 
 MsgBox "Page field" 
 Case xlDataField 
 MsgBox "Data field" 
End Select
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

