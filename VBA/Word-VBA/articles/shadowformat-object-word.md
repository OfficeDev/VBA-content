---
title: ShadowFormat Object (Word)
keywords: vbawd10.chm2508
f1_keywords:
- vbawd10.chm2508
ms.prod: word
api_name:
- Word.ShadowFormat
ms.assetid: 2a179f0b-ec18-c3dd-dd73-51b18f42e0e2
ms.date: 06/08/2017
---


# ShadowFormat Object (Word)

Represents shadow formatting for a shape.


## Remarks

Use the  **Shadow** property to return a **ShadowFormat** object. The following example adds a shadowed rectangle to the active document. The semitransparent, blue shadow is offset 5 points to the right of the rectangle and 3 points above it.


```vb
With ActiveDocument.Shapes _ 
 .AddShape(msoShapeRectangle, 50, 50, 100, 200).Shadow 
 .ForeColor.RGB = RGB(0, 0, 128) 
 .OffsetX = 5 
 .OffsetY = -3 
 .Transparency = 0.5 
 .Visible = True 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


