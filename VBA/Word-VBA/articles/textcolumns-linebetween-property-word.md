---
title: TextColumns.LineBetween Property (Word)
keywords: vbawd10.chm158531685
f1_keywords:
- vbawd10.chm158531685
ms.prod: word
api_name:
- Word.TextColumns.LineBetween
ms.assetid: 102b2ff8-b727-32b4-cd2f-9f9d6e0f0385
ms.date: 06/08/2017
---


# TextColumns.LineBetween Property (Word)

 **True** if vertical lines appear between all the columns in the **TextColumns** collection. Read/write **Long** .


## Syntax

 _expression_ . **LineBetween**

 _expression_ An expression that returns a **[TextColumns](textcolumns-objectword.md)** collection object.


## Remarks

The  **LineBetween** property can be **True** , **False** , or **wdUndefined** .


## Example

This example cycles through each section in the active document and displays a message box if the text columns in the section are separated by vertical lines.


```vb
i = 1 
For each s in ActiveDocument.Sections 
 If s.PageSetup.TextColumns.LineBetween = True Then 
 MsgBox "The columns in section " &; i &; " contain lines." 
 End If 
 i = i + 1 
Next s
```


## See also


#### Concepts


[TextColumns Collection Object](textcolumns-objectword.md)

