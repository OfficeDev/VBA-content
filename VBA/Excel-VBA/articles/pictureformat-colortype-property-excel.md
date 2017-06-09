---
title: PictureFormat.ColorType Property (Excel)
keywords: vbaxl10.chm113003
f1_keywords:
- vbaxl10.chm113003
ms.prod: excel
api_name:
- Excel.PictureFormat.ColorType
ms.assetid: 6c183163-8fbd-3a0f-b087-05d8d2cdbfd5
ms.date: 06/08/2017
---


# PictureFormat.ColorType Property (Excel)

Returns or sets the type of color transformation applied to the specified picture or OLE object. Read/write .


## Syntax

 _expression_ . **ColorType**

 _expression_ An expression that returns a **PictureFormat** object.


## Example

This example sets the color transformation to grayscale through  **[MsoPictureColorType](http://msdn.microsoft.com/library/d11f2d08-2ac9-6cf4-34b8-7ffaabb5d4ae%28Office.15%29.aspx)** for shape one on `myDocument`. Shape one must be either a picture or an OLE object.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).PictureFormat.ColorType = msoPictureGrayScale
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-excel.md)

