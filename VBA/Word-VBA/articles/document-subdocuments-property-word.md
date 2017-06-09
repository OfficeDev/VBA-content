---
title: Document.Subdocuments Property (Word)
keywords: vbawd10.chm158007341
f1_keywords:
- vbawd10.chm158007341
ms.prod: word
api_name:
- Word.Document.Subdocuments
ms.assetid: 4d0047da-03ef-67da-61ed-8bdbeaa55024
ms.date: 06/08/2017
---


# Document.Subdocuments Property (Word)

Returns a  **[Subdocuments](subdocuments-object-word.md)** collection that represents all the subdocuments in the specified document. Read-only.


## Syntax

 _expression_ . **Subdocuments**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the number of subdocuments embedded in the active document.


```vb
MsgBox ActiveDocument.Subdocuments.Count
```

This example displays the path and file name of each subdocument in the active document.




```vb
For Each subdoc In ActiveDocument.Subdocuments 
 If subdoc.HasFile = True Then 
 MsgBox subdoc.Path &; Application.PathSeparator _ 
 &; subdoc.Name 
 Else 
 MsgBox "This subdocument has not been saved." 
 End If 
Next subdoc
```


## See also


#### Concepts


[Document Object](document-object-word.md)

