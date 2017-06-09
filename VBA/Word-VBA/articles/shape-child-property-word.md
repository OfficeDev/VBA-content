---
title: Shape.Child Property (Word)
keywords: vbawd10.chm161480840
f1_keywords:
- vbawd10.chm161480840
ms.prod: word
api_name:
- Word.Shape.Child
ms.assetid: 86102bd1-3df1-384e-589b-c37ba07b4afe
ms.date: 06/08/2017
---


# Shape.Child Property (Word)

 **True** if the shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only **MsoTriState** .


## Syntax

 _expression_ . **Child**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


## Example

This example selects the first shape in the canvas and, if the selected shape is a child shape, fills the shape with the specified color. This example assumes that the first shape in the active document is a drawing canvas that contains multiple shapes.


```vb
Sub FillChildShape() 
 
 Dim shpCanvasItem As Shape 
 
 'Select the first shape in the drawing canvas 
 Set shpCanvasItem = ActiveDocument.Shapes(1).CanvasItems(1) 
 
 'Fill selected shape if it is a child shape 
 With shpCanvasItem 
 If .Child = msoTrue Then 
 .Fill.ForeColor.RGB = RGB(100, 0, 200) 
 Else 
 MsgBox "This shape is not a child shape." 
 End If 
 End With 
 
End Sub
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

