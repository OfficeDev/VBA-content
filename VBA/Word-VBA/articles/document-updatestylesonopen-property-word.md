---
title: Document.UpdateStylesOnOpen Property (Word)
keywords: vbawd10.chm158007362
f1_keywords:
- vbawd10.chm158007362
ms.prod: word
api_name:
- Word.Document.UpdateStylesOnOpen
ms.assetid: 7b126a45-2347-8140-25b8-861672dcc8b5
ms.date: 06/08/2017
---


# Document.UpdateStylesOnOpen Property (Word)

 **True** if the styles in the specified document are updated to match the styles in the attached template each time the document is opened. Read/write **Boolean** .


## Syntax

 _expression_ . **UpdateStylesOnOpen**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example enables the option to update document styles for all open documents and then closes the documents. When any of these documents is reopened, changes to the styles in the attached template will automatically appear in the document.


```vb
For Each doc In Documents 
 doc.UpdateStylesOnOpen = True 
 doc.Close SaveChanges:=wdSaveChanges 
Next doc
```

This example disables the option to update document styles so that changes made to the styles in the attached template aren't reflected in Report.doc.




```vb
Documents("Report.doc").UpdateStylesOnOpen = False
```


## See also


#### Concepts


[Document Object](document-object-word.md)

