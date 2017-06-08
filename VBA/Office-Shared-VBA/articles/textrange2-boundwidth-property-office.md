---
title: TextRange2.BoundWidth Property (Office)
ms.prod: office
api_name:
- Office.TextRange2.BoundWidth
ms.assetid: a5668c93-0206-c26f-41bc-771c1ceef7e6
ms.date: 06/08/2017
---


# TextRange2.BoundWidth Property (Office)

Gets the width, in points, of the text bounding box for the specified text. Read-only.


## Syntax

 _expression_. **BoundWidth**

 _expression_ An expression that returns a **TextRange2** object.


### Return Value

Single


## Remarks

The text bounding box is not the same as the  **TextFrame** object. The **TextFrame** object represents the container in which the text can reside. The text bounding box represents the perimeter immediately surrounding the text.


## Example

This example adds a rounded rectangle to slide one with the same dimensions as the text bounding box.


```
With ActivePresentation.Slides(1).Shapes(1) 
 Set txb = .TextFrame.Text 
 Set roundRect = .AddShape(ppShapeRoundRect, _ 
 txb.BoundLeft, txb.BoundTop, txb.BoundWidth, txb.BoundHeight) 
 roundRect.Fill.Transparency = 0.25 
End With 

```


## See also


#### Concepts


[TextRange2 Object](textrange2-object-office.md)
#### Other resources


[TextRange2 Object Members](textrange2-members-office.md)

