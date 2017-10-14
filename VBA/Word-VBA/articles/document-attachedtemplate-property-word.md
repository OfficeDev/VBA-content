---
title: Document.AttachedTemplate Property (Word)
keywords: vbawd10.chm158007363
f1_keywords:
- vbawd10.chm158007363
ms.prod: word
api_name:
- Word.Document.AttachedTemplate
ms.assetid: e7489e88-ec82-ff16-558b-1dd5470f83c9
ms.date: 06/08/2017
---


# Document.AttachedTemplate Property (Word)

Returns a  **[Template](template-object-word.md)** object that represents the template attached to the specified document. Read/write **Variant** .


## Syntax

 _expression_ . **AttachedTemplate**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

To set this property, specify either the name of the template or an expression that returns a  **Template** object.


## Example

This example displays the name and path of the template attached to the active document.


```vb
Set myTemplate = ActiveDocument.AttachedTemplate 
MsgBox myTemplate.Path &; Application.PathSeparator _ 
 &; myTemplate.Name
```

This example inserts the contents of the Spike (a built-in AutoText entry) at the beginning of document one.




```vb
Set myRange = Documents(1).Range(0, 0) 
Documents(1).AttachedTemplate.AutoTextEntries("Spike") _ 
 .Insert myRange
```

This example attaches the template "Letter.dot" to the active document.




```vb
ActiveDocument.AttachedTemplate = "C:\Templates\Letter.dot"
```


## See also


#### Concepts


[Document Object](document-object-word.md)

