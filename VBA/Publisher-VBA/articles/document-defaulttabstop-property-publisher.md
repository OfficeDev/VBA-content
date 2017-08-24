---
title: Document.DefaultTabStop Property (Publisher)
keywords: vbapb10.chm196616
f1_keywords:
- vbapb10.chm196616
ms.prod: publisher
api_name:
- Publisher.Document.DefaultTabStop
ms.assetid: 245ff7a3-9828-5220-b692-2ce6effb9eb6
ms.date: 06/08/2017
---


# Document.DefaultTabStop Property (Publisher)

Returns or sets a  **Variant** corresponding to the default tab stop for all text in the active publication. Valid range is 1 to 1584 points (0.014" to 22"). Once set, numeric values are considered to be in points. **String** values may be in any unit supported by Microsoft Publisher. Point values are always returned. If values are outside the valid range, an error is returned. Read/write.


## Syntax

 _expression_. **DefaultTabStop**

 _expression_A variable that represents a  **Document** object.


### Return Value

Variant


## Remarks

Use the  **[InchesToPoints](application-inchestopoints-method-publisher.md)** method to convert inches to points.


## Example

This example sets the  **DefaultTabStop** property to 72 points for all text in the active publication.


```vb
Sub SetTab() 
 Application.ActiveDocument.DefaultTabStop = 72 
End Sub 
```


