---
title: TextRange.Start Property (Publisher)
keywords: vbapb10.chm5308433
f1_keywords:
- vbapb10.chm5308433
ms.prod: publisher
api_name:
- Publisher.TextRange.Start
ms.assetid: 40604058-7c3e-b4c7-c793-bbf09091b4c1
ms.date: 06/08/2017
---


# TextRange.Start Property (Publisher)

Returns or sets a  **Long** that represents the starting character position of a text range. Read/write.


## Syntax

 _expression_. **Start**

 _expression_A variable that represents a  **TextRange** object.


### Return Value

Long


## Remarks

If this property is set to a value larger than that of the  **End** property, the **End** property is set to the same value as that of the **Start** property.


## Example

This example makes the first 15 characters of the selected text range bold. This example assumes that text is selected in the active publication.


```vb
Sub SetSelectionRange() 
 With Selection 
 With .TextRange 
 .Start = 0 
 .End = 15 
 .Font.Bold = msoTrue 
 End With 
 End With 
End Sub
```


