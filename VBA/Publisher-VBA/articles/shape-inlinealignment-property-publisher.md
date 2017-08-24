---
title: Shape.InlineAlignment Property (Publisher)
keywords: vbapb10.chm5308694
f1_keywords:
- vbapb10.chm5308694
ms.prod: publisher
api_name:
- Publisher.Shape.InlineAlignment
ms.assetid: daef2761-2a93-25da-9c12-1fed0fdd24ab
ms.date: 06/08/2017
---


# Shape.InlineAlignment Property (Publisher)

Returns or sets a  **PbInlineAlignment** constant that indicates whether an inline shape has left, right, or in-text alignment. Read/write.


## Syntax

 _expression_. **InlineAlignment**

 _expression_A variable that represents a  **Shape** object.


## Remarks

The  **InlineAlignment** property value can be one of the **[PbInlineAlignment](pbinlinealignment-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.

An automation error is returned if the shape is not already inline.


## Example

The following example moves the second shape on the second page of the publication into the text flow by using the  **[MoveIntoTextFlow](shape-moveintotextflow-method-publisher.md)** method. The **InlineAlignment** property is then used to align the shape to the right.


```vb
Dim theShape As Shape 
Dim theRange As TextRange 
 
Set theRange = ActiveDocument.Pages(2).Shapes(1).TextFrame.TextRange 
Set theShape = ActiveDocument.Pages(2).Shapes(2) 
 
If Not theShape.IsInline = msoTrue Then 
 theShape.MoveIntoTextFlow Range:=theRange 
 theShape.InlineAlignment = pbInlineAlignmentRight 
End If
```


