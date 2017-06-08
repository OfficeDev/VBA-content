---
title: Document.Frames Property (Word)
keywords: vbawd10.chm158007319
f1_keywords:
- vbawd10.chm158007319
ms.prod: word
api_name:
- Word.Document.Frames
ms.assetid: 61b7d5dc-6ab4-d29c-6c6e-daac6a2431ed
ms.date: 06/08/2017
---


# Document.Frames Property (Word)

Returns a  **[Frames](frames-object-word.md)** collection that represents all the frames in a document. Read-only.


## Syntax

 _expression_ . **Frames**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example adds a frame around the selection and returns a frame object to the myFrame variable.


```vb
Set myFrame = ActiveDocument.Frames.Add(Range:=Selection.Range)
```


## See also


#### Concepts


[Document Object](document-object-word.md)

