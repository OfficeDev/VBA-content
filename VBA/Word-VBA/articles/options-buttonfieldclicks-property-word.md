---
title: Options.ButtonFieldClicks Property (Word)
keywords: vbawd10.chm162988059
f1_keywords:
- vbawd10.chm162988059
ms.prod: word
api_name:
- Word.Options.ButtonFieldClicks
ms.assetid: 64bb9624-b60d-3999-adf4-9795f18167cd
ms.date: 06/08/2017
---


# Options.ButtonFieldClicks Property (Word)

Returns or sets the number of clicks (either one or two) required to run a GOTOBUTTON or MACROBUTTON field. Read/write  **Long** .


## Syntax

 _expression_ . **ButtonFieldClicks**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Example

This example sets the number of clicks required to run a MACROBUTTON or GOTOBUTTON field to one.


```
Options.ButtonFieldClicks = 1
```


## See also


#### Concepts


[Options Object](options-object-word.md)

