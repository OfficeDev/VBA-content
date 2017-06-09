---
title: ShapeRange.PictureFormat Property (Excel)
keywords: vbaxl10.chm640113
f1_keywords:
- vbaxl10.chm640113
ms.prod: excel
api_name:
- Excel.ShapeRange.PictureFormat
ms.assetid: b7d8ec5c-b0b3-3628-475d-16939c467ad6
ms.date: 06/08/2017
---


# ShapeRange.PictureFormat Property (Excel)

Returns a  **[PictureFormat](pictureformat-object-excel.md)** object that contains picture formatting properties for the specified shape. Applies to a **[ShapeRange](shaperange-object-excel.md)** object that represent pictures or OLE objects. Read-only.


## Syntax

 _expression_ . **PictureFormat**

 _expression_ A variable that represents a **ShapeRange** object.


## Example

This example sets the brightness and contrast for shape one on  `myDocument`. Shape one must be a picture or an OLE object.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).PictureFormat 
 .Brightness = 0.3 
 .Contrast = .75 
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

