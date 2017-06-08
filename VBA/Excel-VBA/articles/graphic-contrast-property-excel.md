---
title: Graphic.Contrast Property (Excel)
keywords: vbaxl10.chm694075
f1_keywords:
- vbaxl10.chm694075
ms.prod: excel
api_name:
- Excel.Graphic.Contrast
ms.assetid: 9715ee08-2d9b-1a5c-1fe9-3b5a73991668
ms.date: 06/08/2017
---


# Graphic.Contrast Property (Excel)

Returns or sets the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write  **Single** .


## Syntax

 _expression_ . **Contrast**

 _expression_ An expression that returns a **Graphic** object.


## Example

This example sets the contrast for shape one on  `myDocument`. Shape one must be either a picture or an OLE object.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).PictureFormat.Contrast = 0.8
```


## See also


#### Concepts


[Graphic Object](graphic-object-excel.md)

