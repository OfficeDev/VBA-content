---
title: PageSetup.FirstPageTray Property (Word)
keywords: vbawd10.chm158400620
f1_keywords:
- vbawd10.chm158400620
ms.prod: word
api_name:
- Word.PageSetup.FirstPageTray
ms.assetid: 60e26cae-2543-adc4-916f-0a0249179990
ms.date: 06/08/2017
---


# PageSetup.FirstPageTray Property (Word)

Returns or sets the paper tray to use for the first page of a document or section. Read/write  **WdPaperTray** .


## Syntax

 _expression_ . **FirstPageTray**

 _expression_ Required. A variable that represents a **[PageSetup](pagesetup-object-word.md)** object.


## Example

This example sets the tray to use for printing the first page of each section in the active document.


```vb
ActiveDocument.PageSetup.FirstPageTray = wdPrinterLowerBin
```

This example sets the tray to use for printing the first page of each section in the selection.




```
Selection.PageSetup.FirstPageTray = wdPrinterUpperBin
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

