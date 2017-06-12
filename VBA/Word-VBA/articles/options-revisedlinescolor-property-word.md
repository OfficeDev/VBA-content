---
title: Options.RevisedLinesColor Property (Word)
keywords: vbawd10.chm162988094
f1_keywords:
- vbawd10.chm162988094
ms.prod: word
api_name:
- Word.Options.RevisedLinesColor
ms.assetid: bc8cd36f-49ac-119a-4f9f-f2e9b20f9bd6
ms.date: 06/08/2017
---


# Options.RevisedLinesColor Property (Word)

Returns or sets the color of changed lines in a document with tracked changes. Read/write  **WdColorIndex** .


## Syntax

 _expression_ . **RevisedLinesColor**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets the color of changed lines to pink.


```
Options.RevisedLinesColor = wdPink
```

This example returns the current status of the  **Color** option under **Changed lines** on the **Track Changes** tab in the **Options** dialog box ( **Tools** menu).




```
temp = Options.RevisedLinesColor
```


## See also


#### Concepts


[Options Object](options-object-word.md)

