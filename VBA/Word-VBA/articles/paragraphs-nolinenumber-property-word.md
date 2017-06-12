---
title: Paragraphs.NoLineNumber Property (Word)
keywords: vbawd10.chm156762217
f1_keywords:
- vbawd10.chm156762217
ms.prod: word
api_name:
- Word.Paragraphs.NoLineNumber
ms.assetid: d548299c-0f1a-d823-f884-57bb8f9be104
ms.date: 06/08/2017
---


# Paragraphs.NoLineNumber Property (Word)

 **True** if line numbers are repressed for the specified paragraphs. Can be **True** , **False** , or **wdUndefined** . Read/write **Long** .


## Syntax

 _expression_ . **NoLineNumber**

 _expression_ A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

Use the  **[LineNumbering](pagesetup-linenumbering-property-word.md)** property of the **[PageSetup](pagesetup-object-word.md)** object to set line numbers.


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


[Paragraphs Collection Object](paragraphs-object-word.md)

