---
title: PageSetup.PaperSize Property (Word)
keywords: vbawd10.chm158400632
f1_keywords:
- vbawd10.chm158400632
ms.prod: word
api_name:
- Word.PageSetup.PaperSize
ms.assetid: 06431f1b-5484-67c6-8ae8-cace3aa9df62
ms.date: 06/08/2017
---


# PageSetup.PaperSize Property (Word)

Returns or sets the paper size. Read/write  **WdPaperSize** .


## Syntax

 _expression_ . **PaperSize**

 _expression_ Required. A variable that represents a **[PageSetup](pagesetup-object-word.md)** object.


## Remarks

Setting the  **PageHeight** or **PageWidth** property changes the **PaperSize** property to **wdPaperCustom** .


## Example

This example sets the paper size to legal for the first document.


```
Documents(1).PageSetup.PaperSize = wdPaperLegal
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

