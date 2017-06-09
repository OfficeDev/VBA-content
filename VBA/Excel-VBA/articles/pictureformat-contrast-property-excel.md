---
title: PictureFormat.Contrast Property (Excel)
keywords: vbaxl10.chm113004
f1_keywords:
- vbaxl10.chm113004
ms.prod: excel
api_name:
- Excel.PictureFormat.Contrast
ms.assetid: 994cfca5-8ddb-d943-63c8-21abe8508de6
ms.date: 06/08/2017
---


# PictureFormat.Contrast Property (Excel)

Returns or sets the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write  **Single** .


## Syntax

 _expression_ . **Contrast**

 _expression_ An expression that returns a **PictureFormat** object.


## Example

This example sets the contrast for shape one on  `myDocument`. Shape one must be either a picture or an OLE object.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).PictureFormat.Contrast = 0.8
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-excel.md)

