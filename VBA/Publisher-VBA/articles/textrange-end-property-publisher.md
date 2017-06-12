---
title: TextRange.End Property (Publisher)
keywords: vbapb10.chm5308434
f1_keywords:
- vbapb10.chm5308434
ms.prod: publisher
api_name:
- Publisher.TextRange.End
ms.assetid: 594cc4b8-d7fb-4b81-4be7-2d416ae513e2
ms.date: 06/08/2017
---


# TextRange.End Property (Publisher)

Sets or returns a  **Long** that represents the ending character position of a selection or text range. Read/write.


## Syntax

 _expression_. **End**

 _expression_A variable that represents an  **TextRange** object.


### Return Value

Long


## Example

This example starts the selection on the fiftieth character of the current text box shape and ends on the one hundred fiftieth character, then makes the text bold.


```vb
Sub test2() 
 With Selection.TextRange 
 .Start = 50 
 .End = 150 
 .Font.Bold = msoTrue 
 End With 
End Sub
```


