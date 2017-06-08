---
title: Document.ContentControls Property (Word)
keywords: vbawd10.chm158007804
f1_keywords:
- vbawd10.chm158007804
ms.prod: word
api_name:
- Word.Document.ContentControls
ms.assetid: 86b5af56-3ab4-2440-237e-42af398b260a
ms.date: 06/08/2017
---


# Document.ContentControls Property (Word)

Returns a  **[ContentControls](contentcontrols-object-word.md)** collection that represents all the content controls in a document. Read-only.


## Syntax

 _expression_ . **ContentControls**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Example

The following example inserts a drop-down list content control into the active document.


```vb
Dim objCC As ContentControl 
 
Set objCC = ActiveDocument.ContentControls.Add(wdContentControlDropdownList) 
 
'List entries 
objCC.DropdownListEntries.Add "Cat" 
objCC.DropdownListEntries.Add "Dog" 
objCC.DropdownListEntries.Add "Horse" 
objCC.DropdownListEntries.Add "Monkey" 
objCC.DropdownListEntries.Add "Snake" 
objCC.DropdownListEntries.Add "Other"
```


## See also


#### Concepts


[Document Object](document-object-word.md)

