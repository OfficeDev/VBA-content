---
title: Shape.PictureFormat Property (Excel)
keywords: vbaxl10.chm636106
f1_keywords:
- vbaxl10.chm636106
ms.prod: excel
api_name:
- Excel.Shape.PictureFormat
ms.assetid: 35a910e8-beac-e4e0-4862-20980d9d633c
ms.date: 06/08/2017
---


# Shape.PictureFormat Property (Excel)

Returns a  **[PictureFormat](pictureformat-object-excel.md)** object that contains picture formatting properties for the specified shape. Applies to a **[Shape](shape-object-excel.md)** object that represent pictures or OLE objects. Read-only.


## Syntax

 _expression_ . **PictureFormat**

 _expression_ A variable that represents a **Shape** object.


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


[Shape Object](shape-object-excel.md)

