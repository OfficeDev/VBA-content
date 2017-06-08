---
title: PivotField.LayoutSubtotalLocation Property (Excel)
keywords: vbaxl10.chm240120
f1_keywords:
- vbaxl10.chm240120
ms.prod: excel
api_name:
- Excel.PivotField.LayoutSubtotalLocation
ms.assetid: 77f250da-7bc3-430d-c6ef-53f79588ecf2
ms.date: 06/08/2017
---


# PivotField.LayoutSubtotalLocation Property (Excel)

Returns or sets the position of the PivotTable field subtotals in relation to (either above or below) the specified field. Read/write  **[XlSubtototalLocationType](xlsubtototallocationtype-enumeration-excel.md)** .


## Syntax

 _expression_ . **LayoutSubtotalLocation**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

You can set this property for any PivotTable field in outline format; however, the formatting appears only if the specified field is a row field other than the innermost (lowest level) row field. For non- OLAP data sources, the value of this property doesn't change when the field is rearranged or when it is added to or from removed from the report.

The  **[LayoutForm](pivotfield-layoutform-property-excel.md)** property determines whether the report appears in table format or in outline format.


## Example

This example displays the state field in the first PivotTable report on the active worksheet in outline format, and it displays the subtotals at the top of the field.


```vb
With ActiveSheet.PivotTables("PivotTable1") _ 
 .PivotFields("state") 
 .LayoutForm = xlOutline 
 .LayoutSubtotalLocation = xlAtTop 
End With
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

