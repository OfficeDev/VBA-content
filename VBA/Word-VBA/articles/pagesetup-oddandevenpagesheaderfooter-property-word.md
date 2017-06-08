---
title: PageSetup.OddAndEvenPagesHeaderFooter Property (Word)
keywords: vbawd10.chm158400627
f1_keywords:
- vbawd10.chm158400627
ms.prod: word
api_name:
- Word.PageSetup.OddAndEvenPagesHeaderFooter
ms.assetid: 82b6d6f1-30fe-2946-241a-cdb0077cabf6
ms.date: 06/08/2017
---


# PageSetup.OddAndEvenPagesHeaderFooter Property (Word)

 **True** if the specified **PageSetup** object has different headers and footers for odd-numbered and even-numbered pages. Read/write **Long** .


## Syntax

 _expression_ . **OddAndEvenPagesHeaderFooter**

 _expression_ An expression that returns a **[PageSetup](pagesetup-object-word.md)** object.


## Remarks

The  **OddAndEvenPagesHeaderFooter** property can be **True** , **False** , or **wdUndefined** .


## Example

This example creates different headers and footers for odd-numbered and even-numbered pages in Document1.


```vb
Set myDoc = Documents("Document1") 
myDoc.PageSetup.OddAndEvenPagesHeaderFooter = True 
With myDoc.Sections(1) 
 .Headers(wdHeaderFooterPrimary).Range _ 
 .InsertAfter "Odd Header" 
 .Headers(wdHeaderFooterEvenPages).Range _ 
 .InsertAfter "Even Header" 
End With
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

