---
title: Document.Comments Property (Word)
keywords: vbawd10.chm158007305
f1_keywords:
- vbawd10.chm158007305
ms.prod: word
api_name:
- Word.Document.Comments
ms.assetid: 1597a002-afa4-743d-60a6-ffd398f2b599
ms.date: 06/08/2017
---


# Document.Comments Property (Word)

Returns a  **[Comments](comments-object-word.md)** collection that represents all the comments in the specified document. Read-only.


## Syntax

 _expression_ . **Comments**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example compares the author name of each comment in the active document with the user name on the  **User Information** tab in the **Options** dialog box ( **Tools** menu). If the names aren't the same, the comment reference mark is formatted to appear in red.


```vb
For Each comm In ActiveDocument.Comments 
 If comm.Author <> Application.UserName Then _ 
 comm.Reference.Font.ColorIndex = wdRed 
Next comm
```


## See also


#### Concepts


[Document Object](document-object-word.md)

