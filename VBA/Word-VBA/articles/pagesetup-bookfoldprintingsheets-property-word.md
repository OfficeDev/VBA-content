---
title: PageSetup.BookFoldPrintingSheets Property (Word)
keywords: vbawd10.chm158401737
f1_keywords:
- vbawd10.chm158401737
ms.prod: word
api_name:
- Word.PageSetup.BookFoldPrintingSheets
ms.assetid: 88431024-42a0-d92e-a62b-eeaedbe0c945
ms.date: 06/08/2017
---


# PageSetup.BookFoldPrintingSheets Property (Word)

Returns or sets a  **Long** which represents the number of pages for each booklet. Read/write **Boolean** .


## Syntax

 _expression_ . **BookFoldPrintingSheets**

 _expression_ An expression that returns a **[PageSetup](pagesetup-object-word.md)** object.


## Example

This example turns the active document into a booklet that will print in sixteen-page booklets.


```vb
Sub Booklet() 
 With PageSetup 
 .BookFoldPrinting = True 
 .BookFoldPrintingSheets = 16 
 End With 
End Sub
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

