---
title: FormatConditions Object (Excel)
keywords: vbaxl10.chm509072
f1_keywords:
- vbaxl10.chm509072
ms.prod: excel
api_name:
- Excel.FormatConditions
ms.assetid: 2486d4b4-605c-76d8-132a-694c0c600a81
ms.date: 06/08/2017
---


# FormatConditions Object (Excel)

Represents the collection of conditional formats for a single range.


## Remarks

 The **FormatConditions** collection can contain multiple conditional formats. Each format is represented by a **[FormatCondition](formatcondition-object-excel.md)** object.

For more information about conditional formats, see the [FormatCondition](formatcondition-object-excel.md) object.

Use the  **FormatConditions** property to return a **FormatConditions** object. Use the **[Add](formatconditions-add-method-excel.md)** method to create a new conditional format, and use the **[Modify](formatcondition-modify-method-excel.md)** method to change an existing conditional format.


## Example

The following example adds a conditional format to cells E1:E10.


```
With Worksheets(1).Range("e1:e10").FormatConditions _ 
 .Add(xlCellValue, xlGreater, "=$a$1") 
 With .Borders 
 .LineStyle = xlContinuous 
 .Weight = xlThin 
 .ColorIndex = 6 
 End With 
 With .Font 
 .Bold = True 
 .ColorIndex = 3 
 End With 
End With
```


## Methods



|**Name**|
|:-----|
|[Add](formatconditions-add-method-excel.md)|
|[AddAboveAverage](formatconditions-addaboveaverage-method-excel.md)|
|[AddColorScale](formatconditions-addcolorscale-method-excel.md)|
|[AddDatabar](formatconditions-adddatabar-method-excel.md)|
|[AddIconSetCondition](formatconditions-addiconsetcondition-method-excel.md)|
|[AddTop10](formatconditions-addtop10-method-excel.md)|
|[AddUniqueValues](formatconditions-adduniquevalues-method-excel.md)|
|[Delete](formatconditions-delete-method-excel.md)|
|[Item](formatconditions-item-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](formatconditions-application-property-excel.md)|
|[Count](formatconditions-count-property-excel.md)|
|[Creator](formatconditions-creator-property-excel.md)|
|[Parent](formatconditions-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
