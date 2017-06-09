---
title: PageSetup.OtherPagesTray Property (Word)
keywords: vbawd10.chm158400621
f1_keywords:
- vbawd10.chm158400621
ms.prod: word
api_name:
- Word.PageSetup.OtherPagesTray
ms.assetid: df6a8e6d-2b49-d633-cd2b-5d3099410a73
ms.date: 06/08/2017
---


# PageSetup.OtherPagesTray Property (Word)

Returns or sets the paper tray to be used for all but the first page of a document or section. Read/write  **WdPaperTray** .


## Syntax

 _expression_ . **OtherPagesTray**

 _expression_ Required. A variable that represents a **[PageSetup](pagesetup-object-word.md)** object.


## Example

This example sets the tray to be used for printing all but the first page of each section in the active document.


```vb
ActiveDocument.PageSetup.OtherPagesTray = wdPrinterUpperBin
```

This example sets the tray to be used for printing all but the first page of each section in the selection.




```
Selection.PageSetup.OtherPagesTray = wdPrinterLowerBin
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

