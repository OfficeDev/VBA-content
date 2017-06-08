---
title: Paragraph.NoLineNumber Property (Word)
keywords: vbawd10.chm156696681
f1_keywords:
- vbawd10.chm156696681
ms.prod: word
api_name:
- Word.Paragraph.NoLineNumber
ms.assetid: f713018a-1024-25fd-7d25-07c278426ba3
ms.date: 06/08/2017
---


# Paragraph.NoLineNumber Property (Word)

 **True** if line numbers are repressed for the specified paragraph. Read/write **Long** .


## Syntax

 _expression_ . **NoLineNumber**

 _expression_ A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

This property can be  **True** , **False** , or **wdUndefined** . Use the **[LineNumbering](pagesetup-linenumbering-property-word.md)** property of the **[PageSetup](pagesetup-object-word.md)** object to set line numbers.


## Example

This example enables line numbering for the active document. The starting number is set to 1, and the numbering is continuous throughout all sections in the document. Line numbering is then repressed for the second paragraph.


```vb
With ActiveDocument.PageSetup.LineNumbering 
 .Active = True 
 .StartingNumber = 1 
 .CountBy = 1 
 .RestartMode = wdRestartContinuous 
End With 
ActiveDocument.Paragraphs(2).NoLineNumber = True
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

