---
title: Shape.VerticalFlip Property (Excel)
keywords: vbaxl10.chm636112
f1_keywords:
- vbaxl10.chm636112
ms.prod: excel
api_name:
- Excel.Shape.VerticalFlip
ms.assetid: 3b50edac-a167-8e07-3286-6ced14bb715d
ms.date: 06/08/2017
---


# Shape.VerticalFlip Property (Excel)

 **True** if the specified shape is flipped around the vertical axis. Read-only **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **VerticalFlip**

 _expression_ A variable that represents a **Shape** object.


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


[Shape Object](shape-object-excel.md)

