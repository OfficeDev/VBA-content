---
title: ShapeRange.HorizontalFlip Property (Excel)
keywords: vbaxl10.chm640106
f1_keywords:
- vbaxl10.chm640106
ms.prod: excel
api_name:
- Excel.ShapeRange.HorizontalFlip
ms.assetid: 3b5f3755-987c-cd48-44a2-8be8bdd886dd
ms.date: 06/08/2017
---


# ShapeRange.HorizontalFlip Property (Excel)

 **True** if the specified shape is flipped around the horizontal axis. Read-only **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **HorizontalFlip**

 _expression_ A variable that represents a **ShapeRange** object.


## Example

This example restores each shape on  `myDocument` to its original state if it's been flipped horizontally or vertically.


```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
    If s.HorizontalFlip Then s.Flip msoFlipHorizontal 
    If s.VerticalFlip Then s.Flip msoFlipVertical 
Next
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

