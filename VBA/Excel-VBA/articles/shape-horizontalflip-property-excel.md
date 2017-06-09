---
title: Shape.HorizontalFlip Property (Excel)
keywords: vbaxl10.chm636099
f1_keywords:
- vbaxl10.chm636099
ms.prod: excel
api_name:
- Excel.Shape.HorizontalFlip
ms.assetid: e9b64a81-3aef-5d42-0a65-5d03d564a71f
ms.date: 06/08/2017
---


# Shape.HorizontalFlip Property (Excel)

 **True** if the specified shape is flipped around the horizontal axis. Read-only **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **HorizontalFlip**

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

