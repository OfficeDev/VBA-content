---
title: WrapFormat Object (Word)
keywords: vbawd10.chm2499
f1_keywords:
- vbawd10.chm2499
ms.prod: word
api_name:
- Word.WrapFormat
ms.assetid: 08396db4-f8e0-12fd-2b9f-3a0a61169ac4
ms.date: 06/08/2017
---


# WrapFormat Object (Word)

Represents all the properties for wrapping text around a shape or shape range.


## Remarks

Use the  **WrapFormat** property to return the **WrapFormat** object. The following example adds an oval to the active document and specifies that document text wrap around the left and right sides of the square that circumscribes the oval. There will be a 0.1-inch margin between the document text and the top, bottom, left side, and right side of the square.


```vb
Set myOval = _ 
 ActiveDocument.Shapes.AddShape(msoShapeOval, 36, 36, 100, 35) 
With myOval.WrapFormat 
 .Type = wdWrapSquare 
 .Side = wdWrapBoth 
 .DistanceTop = InchesToPoints(0.1) 
 .DistanceBottom = InchesToPoints(0.1) 
 .DistanceLeft = InchesToPoints(0.1) 
 .DistanceRight = InchesToPoints(0.1) 
End With
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

