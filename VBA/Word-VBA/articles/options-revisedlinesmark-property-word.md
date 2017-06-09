---
title: Options.RevisedLinesMark Property (Word)
keywords: vbawd10.chm162988091
f1_keywords:
- vbawd10.chm162988091
ms.prod: word
api_name:
- Word.Options.RevisedLinesMark
ms.assetid: ecc358f2-4bf6-7546-5400-938a3dae6b77
ms.date: 06/08/2017
---


# Options.RevisedLinesMark Property (Word)

Returns or sets the placement of changed lines in a document with tracked changes. Read/write  **WdRevisedLinesMark** .


## Syntax

 _expression_ . **RevisedLinesMark**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets changed lines to appear in the left margin of every page.


```
Options.RevisedLinesMark = wdRevisedLinesMarkLeftBorder
```

This example returns the current status of the  **Mark** option under **Changed lines** on the **Track Changes** tab in the **Options** dialog box ( **Tools** menu).




```
temp = Options.RevisedLinesMark
```


## See also


#### Concepts


[Options Object](options-object-word.md)

