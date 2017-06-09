---
title: Document.SpellingErrors Property (Word)
keywords: vbawd10.chm158007394
f1_keywords:
- vbawd10.chm158007394
ms.prod: word
api_name:
- Word.Document.SpellingErrors
ms.assetid: c8a987a1-3705-ea0a-103a-99b2f17f5c6b
ms.date: 06/08/2017
---


# Document.SpellingErrors Property (Word)

Returns a  **[ProofreadingErrors](proofreadingerrors-object-word.md)** collection that represents the words identified as spelling errors in the specified document or range. Read-only.


## Syntax

 _expression_ . **SpellingErrors**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example checks the active document for spelling errors and displays the number of errors found.


```vb
myErr = ActiveDocument.SpellingErrors.Count 
If myErr = 0 Then 
 Msgbox "No spelling errors found." 
Else 
 Msgbox myErr &; " spelling errors found." 
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)

