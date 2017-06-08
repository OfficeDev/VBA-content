---
title: Options.RevisedPropertiesColor Property (Word)
keywords: vbawd10.chm162988109
f1_keywords:
- vbawd10.chm162988109
ms.prod: word
api_name:
- Word.Options.RevisedPropertiesColor
ms.assetid: 00b04099-0cb2-31e1-dc34-ad9203919f52
ms.date: 06/08/2017
---


# Options.RevisedPropertiesColor Property (Word)

Returns or sets the color used to mark formatting changes while change tracking is enabled. Read/write  **WdColorIndex** .


## Syntax

 _expression_ . **RevisedPropertiesColor**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Remarks

If deleted or inserted text has formatting changes, the  **RevisedPropertiesColor** property is overridden by the **DeletedTextColor** or **InsertedTextColor** property.


## Example

This example tracks changes in the active document, sets the color of text with changed formatting to teal, and applies bold formatting to the selection.


```vb
ActiveDocument.TrackRevisions = True 
Options.RevisedPropertiesColor = wdTeal 
Selection.Font.Bold = True
```

This example returns the option selected in the Color box under Track Changes options on the Track Changes tab in the Options dialog box (Tools menu).




```
temp = Options.RevisedPropertiesColor
```


## See also


#### Concepts


[Options Object](options-object-word.md)

