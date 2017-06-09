---
title: TextRange.Paste Method (Publisher)
keywords: vbapb10.chm5308482
f1_keywords:
- vbapb10.chm5308482
ms.prod: publisher
api_name:
- Publisher.TextRange.Paste
ms.assetid: dd29c9ab-7f56-3604-3390-8f5a3b97821f
ms.date: 06/08/2017
---


# TextRange.Paste Method (Publisher)

Pastes the text on the Clipboard into the specified text range, and returns a  **[TextRange](textrange-object-publisher.md)** object that represents the pasted text.


## Syntax

 _expression_. **Paste**

 _expression_A variable that represents a  **TextRange** object.


### Return Value

TextRange


## Example

This example deletes the text in shape one on page one in the active publication, places it on the Clipboard, and then pastes it after the first word in shape two on the same page. This example assumes that each shape contains text.


```vb
With ActiveDocument.Pages(1) 
 .Shapes(1).TextFrame.TextRange.Cut 
 .Shapes(2).TextFrame.TextRange. _ 
 Words(1).Paste 
End With 

```


