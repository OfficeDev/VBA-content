---
title: ShapeRange.Align Method (Excel)
keywords: vbaxl10.chm640077
f1_keywords:
- vbaxl10.chm640077
ms.prod: excel
api_name:
- Excel.ShapeRange.Align
ms.assetid: 7a4e6442-6730-ab7d-93b5-4c091ada6b14
ms.date: 06/08/2017
---


# ShapeRange.Align Method (Excel)

Aligns the shapes in the specified range of shapes.


## Syntax

 _expression_ . **Align**( **_AlignCmd_** , **_RelativeTo_** )

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _AlignCmd_|Required| **[MsoAlignCmd](http://msdn.microsoft.com/library/d4c62376-bb88-65e1-8922-ced2e5256ff7%28Office.15%29.aspx)**|Specifies the way the shapes in the specified shape range are to be aligned.|
| _RelativeTo_|Required| **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)**|Not used in Microsoft Excel. Must be  **False** .|

## Example

This example aligns the left edges of all the shapes in the specified range in  `myDocument` with the left edge of the leftmost shape in the range.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.SelectAll 
Selection.ShapeRange.Align msoAlignLefts, False
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

