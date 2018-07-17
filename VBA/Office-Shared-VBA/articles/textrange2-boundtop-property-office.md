---
title: TextRange2.BoundTop Property (Office)
ms.prod: office
api_name:
- Office.TextRange2.BoundTop
ms.assetid: b225b65e-04a0-1938-9520-ea71eed13b04
ms.date: 06/08/2017
---


# TextRange2.BoundTop Property (Office)

Gets the top coordinate, in points, of the text bounding box for the specified text. Read-only.


## Syntax

 _expression_. **BoundTop**

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

