---
title: Range.ShowDetail Property (Excel)
keywords: vbaxl10.chm144196
f1_keywords:
- vbaxl10.chm144196
ms.prod: excel
api_name:
- Excel.Range.ShowDetail
ms.assetid: 1908af55-f61a-2a0f-d828-350e9a680377
ms.date: 06/08/2017
---


# Range.ShowDetail Property (Excel)

 **True** if the outline is expanded for the specified range (so that the detail of the column or row is visible). The specified range must be a single summary column or row in an outline. Read/write **Variant** . For the **PivotItem** object (or the **Range** object if the range is in a PivotTable report), this property is set to **True** if the item is showing detail.


## Syntax

 _expression_ . **ShowDetail**

 _expression_ A variable that represents a **Range** object.


## Remarks

This property isn't available for OLAP data sources.

If the specified range isn't in a PivotTable report, the following statements are true:


- The range must be in a single summary row or column.
    
- This property returns  **False** if _any_ of the children of the row or column are hidden.
    
- Setting this property to  **True** is equivalent to unhiding all the children of the summary row or column.
    
- Setting this property to  **False** is equivalent to hiding all the children of the summary row or column.
    
If the specified range is in a PivotTable report, it's possible to set this property for more than one cell at a time if the range is contiguous. This property can be returned only if the range is a single cell.


## Example

This example shows detail for the summary row of an outline on Sheet1. Before running this example, create a simple outline that contains a single summary row, and then collapse the outline so that only the summary row is showing. Select one of the cells in the summary row, and then run the example.


```vb
Worksheets("Sheet1").Activate 
Set myRange = ActiveCell.CurrentRegion 
lastRow = myRange.Rows.Count 
myRange.Rows(lastRow).ShowDetail = True
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

