---
title: Document.Sections Property (Word)
keywords: vbawd10.chm158007311
f1_keywords:
- vbawd10.chm158007311
ms.prod: word
api_name:
- Word.Document.Sections
ms.assetid: 83c3ec94-b0ef-e8a5-b17a-ad657e7197b2
ms.date: 06/08/2017
---


# Document.Sections Property (Word)

Returns a  **[Section](section-object-word.md)** collection that represents the sections in the specified document. Read-only.


## Syntax

 _expression_ . **Sections**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example sets the page orientation for all the sections in the active document.


```vb
For Each sec In ActiveDocument.Sections 
 sec.PageSetup.Orientation = wdOrientLandscape 
Next sec
```

This example creates a new document then adds some text to the document. It then creates a new section in the document and inserts text into the new section.




```vb
Set myDoc = Documents.Add 
Selection.InsertAfter "This is section 1." 
Set mysec = myDoc.Sections.Add 
mysec.Range.InsertAfter "This is section 2"
```


## See also


#### Concepts


[Document Object](document-object-word.md)

