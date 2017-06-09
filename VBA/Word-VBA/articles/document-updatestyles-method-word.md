---
title: Document.UpdateStyles Method (Word)
keywords: vbawd10.chm158007423
f1_keywords:
- vbawd10.chm158007423
ms.prod: word
api_name:
- Word.Document.UpdateStyles
ms.assetid: fe713979-27e1-c81c-198d-5e25564233c2
ms.date: 06/08/2017
---


# Document.UpdateStyles Method (Word)

Copies all styles from the attached template into the document, overwriting any existing styles in the document that have the same name.


## Syntax

 _expression_ . **UpdateStyles**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example copies the styles from the attached template into each open document, and then it closes each document.


```vb
For Each aDoc In Documents 
 aDoc.UpdateStyles 
 aDoc.Close SaveChanges:=wdSaveChanges 
Next aDoc
```

This example changes the formatting of the Heading 1 style in the template attached to the active document. The  **UpdateStyles** method updates the styles in the active document, including the Heading 1 style.




```vb
Set aDoc = ActiveDocument.AttachedTemplate.OpenAsDocument 
With aDoc.Styles(wdStyleHeading1).Font 
 .Name = "Arial" 
 .Bold = False 
End With 
aDoc.Close SaveChanges:=wdSaveChanges 
ActiveDocument.UpdateStyles
```


## See also


#### Concepts


[Document Object](document-object-word.md)

