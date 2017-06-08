---
title: PageSetup.DifferentFirstPageHeaderFooter Property (Word)
keywords: vbawd10.chm158400628
f1_keywords:
- vbawd10.chm158400628
ms.prod: word
api_name:
- Word.PageSetup.DifferentFirstPageHeaderFooter
ms.assetid: 50664181-4a3b-1b68-98e2-558fa9ee538e
ms.date: 06/08/2017
---


# PageSetup.DifferentFirstPageHeaderFooter Property (Word)

 **True** if a different header or footer is used on the first page. Can be **True** , **False** , or **wdUndefined** . Read/write **Long** .


## Syntax

 _expression_ . **DifferentFirstPageHeaderFooter**

 _expression_ An expression that returns a **[PageSetup](pagesetup-object-word.md)** object.


## Example

This example checks each section in the active document for headers and footers that are different on the first page and displays a message if any are found.


```vb
Dim secLoop As Section 
 
For Each secLoop In ActiveDocument.Sections 
 If secLoop.PageSetup _ 
 .DifferentFirstPageHeaderFooter = True Then 
 Msgbox "Section " &; secLoop.Index _ 
 &; " has different first page headers &; footers." 
 End If 
Next secLoop
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

