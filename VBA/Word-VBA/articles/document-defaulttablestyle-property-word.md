---
title: Document.DefaultTableStyle Property (Word)
keywords: vbawd10.chm158007661
f1_keywords:
- vbawd10.chm158007661
ms.prod: word
api_name:
- Word.Document.DefaultTableStyle
ms.assetid: b6782b12-09a6-77b0-a52d-81d4028e7c19
ms.date: 06/08/2017
---


# Document.DefaultTableStyle Property (Word)

Returns a  **Variant** that represents the table style that is applied to all newly created tables in a document. Read-only.


## Syntax

 _expression_ . **DefaultTableStyle**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Example

This example checks to see if the default table style used in the active document is named "Table Normal" and, if it is, changes the default table style to "TableStyle1." This example assumes that you have a table style named "TableStyle1."


```vb
Sub TableDefaultStyle() 
 With ActiveDocument 
 If .DefaultTableStyle = "Table Normal" Then 
 .SetDefaultTableStyle _ 
 Style:="TableStyle1", SetInTemplate:=True 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

