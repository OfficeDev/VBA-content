---
title: Shape.Top Property (Publisher)
keywords: vbapb10.chm2228306
f1_keywords:
- vbapb10.chm2228306
ms.prod: publisher
api_name:
- Publisher.Shape.Top
ms.assetid: 76ab84a9-651c-ddc6-6f7f-f98e2b71074f
ms.date: 06/08/2017
---


# Shape.Top Property (Publisher)

Returns or sets a  **Variant** that represents the distance between the top of the page and the top of a shape. Read/write.


## Syntax

 _expression_. **Top**

 _expression_A variable that represents a  **Shape** object.


## Example

This example changes the position, size, and type of shape of the first shape on the first page of the active publication. This example assumes there is at least one shape on the first page of the active publication.


```vb
Sub MoveSizeChangeShape() 
 With ActiveDocument.Pages(1).Shapes(1) 
 .Top = 72 
 .Left = 72 
 .Width = 150 
 .Height = 150 
 .AutoShapeType = msoShape5pointStar 
 End With 
End Sub
```


