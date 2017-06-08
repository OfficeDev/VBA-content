---
title: Shape.MoveOutOfTextFlow Method (Publisher)
keywords: vbapb10.chm2228357
f1_keywords:
- vbapb10.chm2228357
ms.prod: publisher
api_name:
- Publisher.Shape.MoveOutOfTextFlow
ms.assetid: 44411d6b-a627-f0c1-0576-2918f586ff0b
ms.date: 06/08/2017
---


# Shape.MoveOutOfTextFlow Method (Publisher)

Moves a given inline shape out of its containing text range, defined by  ** [TextRange Object](textrange-object-publisher.md)**, and makes the shape fixed.


## Syntax

 _expression_. **MoveOutOfTextFlow**

 _expression_A variable that represents a  **Shape** object.


## Remarks

An automation error is returned if the shape to be moved is not already inline.

After the  **MoveOutOfTextFlow** method is called on an inline shape, the shape will maintain its position on the page, but it will no longer be inline.


## Example

The following example moves the first inline shape contained in a given text range out of the text flow.


```vb
Dim theShape As Shape 
 
Set theShape = ActiveDocument.Pages(2).Shapes(1) _ 
 .TextFrame.TextRange.InlineShapes(1) 
 
theShape.MoveOutOfTextFlow
```


