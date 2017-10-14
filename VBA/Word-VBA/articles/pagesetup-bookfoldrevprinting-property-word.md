---
title: PageSetup.BookFoldRevPrinting Property (Word)
keywords: vbawd10.chm158401736
f1_keywords:
- vbawd10.chm158401736
ms.prod: word
api_name:
- Word.PageSetup.BookFoldRevPrinting
ms.assetid: 3d6db95a-1c2d-424f-f704-ed7d1c05895c
ms.date: 06/08/2017
---


# PageSetup.BookFoldRevPrinting Property (Word)

 **True** for Microsoft Word to reverse the printing order for book fold printing of bidirectional or Asian language documents. Read/write **Boolean** .


## Syntax

 _expression_ . **BookFoldRevPrinting**

 _expression_ An expression that returns a **[PageSetup](pagesetup-object-word.md)** object.


## Example

This example switches from left-to-right book printing to right-to-left book printing for a bidirectional or Asian language document that will print in sixteen-page increments.


```vb
Sub BookletRev() 
 With PageSetup 
 .BookFoldRevPrinting = True 
 .BookFoldPrintingSheets = 16 
 End With 
End Sub
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

