---
title: PageSetup.Gutter Property (Word)
keywords: vbawd10.chm158400616
f1_keywords:
- vbawd10.chm158400616
ms.prod: word
api_name:
- Word.PageSetup.Gutter
ms.assetid: ec16576d-1b77-543e-aa8a-b52457f56675
ms.date: 06/08/2017
---


# PageSetup.Gutter Property (Word)

Returns or sets the amount (in points) of extra margin space added to each page in a document or section for binding. Read/write  **Single** .


## Syntax

 _expression_ . **Gutter**

 _expression_ A variable that represents a **[PageSetup](pagesetup-object-word.md)** object.


## Remarks

If the  **[MirrorMargins](pagesetup-mirrormargins-property-word.md)** property is set to **True** , the **Gutter** property adds the extra space to the inside margins. Otherwise, the extra space is added to the left margin.


## Example

This example adds 1 inch (72 points) to the inside margins of the active document.


```vb
With ActiveDocument.PageSetup 
 .MirrorMargins = True 
 .Gutter = 72 
End With
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

