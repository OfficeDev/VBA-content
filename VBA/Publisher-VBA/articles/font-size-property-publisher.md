---
title: Font.Size Property (Publisher)
keywords: vbapb10.chm5373957
f1_keywords:
- vbapb10.chm5373957
ms.prod: publisher
api_name:
- Publisher.Font.Size
ms.assetid: 485f68fe-c6d7-8288-042e-fc4c35c37b2d
ms.date: 06/08/2017
---


# Font.Size Property (Publisher)

Represents the size of the characters in the text range in points. Read/write.


## Syntax

 _expression_. **Size**

 _expression_An expression that returns a  **Font** object.


### Return Value

Variant


## Example

This example inserts text and then sets the font size of the seventh word of the inserted text to 20 points.


```vb
Sub IncreaseFontSizeOfSelection() 
 With Selection.TextRange 
 .InsertBefore vbLf &; "This is a demonstration of font size." 
 .Words(7).Font.Size = 20 
 End With 
End Sub
```


