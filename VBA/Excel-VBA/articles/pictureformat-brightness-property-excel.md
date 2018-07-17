---
title: PictureFormat.Brightness Property (Excel)
keywords: vbaxl10.chm113002
f1_keywords:
- vbaxl10.chm113002
ms.prod: excel
api_name:
- Excel.PictureFormat.Brightness
ms.assetid: f17ee171-47da-c982-2f48-9ee333193add
ms.date: 06/08/2017
---


# PictureFormat.Brightness Property (Excel)

Returns or sets the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write  **Single** .


## Syntax

 _expression_ . **Brightness**

 _expression_ A variable that represents a **PictureFormat** object.


## Example

This example sets the brightness for shape one on  `myDocument`. Shape one must be either a picture or an OLE object.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).PictureFormat.Brightness = 0.3
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-excel.md)

