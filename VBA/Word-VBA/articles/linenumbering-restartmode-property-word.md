---
title: LineNumbering.RestartMode Property (Word)
keywords: vbawd10.chm158466148
f1_keywords:
- vbawd10.chm158466148
ms.prod: word
api_name:
- Word.LineNumbering.RestartMode
ms.assetid: f812d5ab-4921-5d6e-a2f8-51d324c29333
ms.date: 06/08/2017
---


# LineNumbering.RestartMode Property (Word)

Returns or sets the way line numbering runs — that is, whether it starts over at the beginning of a new page or section or runs continuously. Read/write  **WdNumberingRule** .


## Syntax

 _expression_ . **RestartMode**

 _expression_ Required. A variable that represents a **[LineNumbering](linenumbering-object-word.md)** object.


## Remarks

You must be in print layout view to see line numbering.


## Example

This example enables line numbering for the active document. The starting number is set to 1, every tenth line number is shown, and the numbering starts over at the beginning of each section.


```vb
set myDoc = ActiveDocument 
With myDoc.PageSetup.LineNumbering 
 .Active = True 
 .StartingNumber = 1 
 .CountBy = 10 
 .RestartMode = wdRestartSection 
End With
```


## See also


#### Concepts


[LineNumbering Object](linenumbering-object-word.md)

