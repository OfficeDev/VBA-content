---
title: Shape.Line Property (Word)
keywords: vbawd10.chm161480816
f1_keywords:
- vbawd10.chm161480816
ms.prod: word
api_name:
- Word.Shape.Line
ms.assetid: 3bb8d585-8af8-a3fc-f61c-d7bcfe4ffa13
ms.date: 06/08/2017
---


# Shape.Line Property (Word)

Returns a  **LineFormat** object that contains line formatting properties for the specified shape. Read-only.


## Syntax

 _expression_ . **Line**

 _expression_ A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

For a line, the  **LineFormat** object represents the line itself; for a shape with a border, the **LineFormat** object represents the border.


## Example

This example adds a blue dashed line to  _myDocument_ .


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes.AddLine(10, 10, 250, 250).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(50, 0, 128) 
End With
```

This example adds a cross to  _myDocument_ and then sets its border to be 8 points thick and red.




```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes.AddShape(msoShapeCross, 10, 10, 50, 70).Line 
 .Weight = 8 
 .ForeColor.RGB = RGB(255, 0, 0) 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

