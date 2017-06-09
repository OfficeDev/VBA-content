---
title: TextRange2.BoundLeft Property (PowerPoint)
ms.assetid: b7adce04-116c-4487-94e7-f895ce7bfc4e
ms.date: 06/08/2017
ms.prod: powerpoint
---


# TextRange2.BoundLeft Property (PowerPoint)

Gets the left coordinate, in points, of the text bounding box for the specified text. Read-only.


## Syntax

 _expression_. **BoundLeft**

 _expression_ An expression that returns a **TextRange2** object.


### Return Value

Single


## Remarks

The text bounding box is not the same as the  **TextFrame** object. The **TextFrame** object represents the container in which the text can reside. The text bounding box represents the perimeter immediately surrounding the text.


## Example

This example adds a rounded rectangle to slide one with the same dimensions as the text bounding box.


```vb
With ActivePresentation.Slides(1).Shapes(1) 
 Set txb = .TextFrame.Text 
 Set roundRect = .AddShape(ppShapeRoundRect, _ 
 txb.BoundLeft, txb.BoundTop, txb.BoundWidth, txb.BoundHeight) 
 roundRect.Fill.Transparency = 0.25 
End With 

```


## See also


#### Other resources


[TextRange2 Object](http://msdn.microsoft.com/library/a6a59c9b-9b64-c1e2-2e98-a1f99025c877%28Office.15%29.aspx)
[TextRange2 Object Members](http://msdn.microsoft.com/library/26daffff-b9ef-fd94-f5b7-ed3a09840cb6%28Office.15%29.aspx)

