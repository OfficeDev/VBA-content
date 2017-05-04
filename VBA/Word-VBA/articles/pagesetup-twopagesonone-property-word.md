---
title: PageSetup.TwoPagesOnOne Property (Word)
keywords: vbawd10.chm158400633
f1_keywords:
- vbawd10.chm158400633
ms.prod: WORD
api_name:
- Word.PageSetup.TwoPagesOnOne
ms.assetid: c9d8edac-1fea-5fdb-a4e2-193920fa89d1
---


# PageSetup.TwoPagesOnOne Property (Word)

 **True** if Microsoft Word prints the specified document two pages per sheet. Read/write **Boolean** .


## Syntax

 _expression_ . **TwoPagesOnOne**

 _expression_ An expression that returns a **[PageSetup](pagesetup-object-word.md)** object.


## Example

This example sets Microsoft Word to print the active document two pages per sheet.


```vb
ActiveDocument.PageSetup.TwoPagesOnOne = True
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

