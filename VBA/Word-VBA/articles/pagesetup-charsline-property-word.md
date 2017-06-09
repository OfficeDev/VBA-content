---
title: PageSetup.CharsLine Property (Word)
keywords: vbawd10.chm158400635
f1_keywords:
- vbawd10.chm158400635
ms.prod: word
api_name:
- Word.PageSetup.CharsLine
ms.assetid: 7539359a-aecd-0676-7e93-3e00cc2bf461
ms.date: 06/08/2017
---


# PageSetup.CharsLine Property (Word)

Returns or sets the number of characters per line in the document grid. Read/write  **Single** .


## Syntax

 _expression_ . **CharsLine**

 _expression_ A variable that represents a **[PageSetup](pagesetup-object-word.md)** object.


## Example

This example sets the number of characters per line to 42 for the active document.


```vb
ActiveDocument.PageSetup.CharsLine = 42
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

