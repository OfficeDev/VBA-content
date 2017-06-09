---
title: TextFrame.DeleteText Method (Word)
keywords: vbawd10.chm162665370
f1_keywords:
- vbawd10.chm162665370
ms.prod: word
api_name:
- Word.TextFrame.DeleteText
ms.assetid: a5fbf67a-c4d2-9b12-e326-86d63150debc
ms.date: 06/08/2017
---


# TextFrame.DeleteText Method (Word)

Deletes the text from a text frame and all the associated properties of the text, including font attributes.


## Syntax

 _expression_ . **DeleteText**

 _expression_ A variable that represents a **[TextFrame](textframe-object-word.md)** object.


## Example

The following code example deletes the text from the first shape in the active document, if that shape contains text. 


```vb
Public Sub DeleteText_Example() 
 ActiveDocument.Shapes(1).TextFrame.DeleteText 
End Sub
```


## See also


#### Concepts


[TextFrame Object](textframe-object-word.md)

