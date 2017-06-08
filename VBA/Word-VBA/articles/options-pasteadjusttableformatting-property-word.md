---
title: Options.PasteAdjustTableFormatting Property (Word)
keywords: vbawd10.chm162988463
f1_keywords:
- vbawd10.chm162988463
ms.prod: word
api_name:
- Word.Options.PasteAdjustTableFormatting
ms.assetid: 8c486ea0-d653-b82a-8507-c192d4d11ecb
ms.date: 06/08/2017
---


# Options.PasteAdjustTableFormatting Property (Word)

 **True** if Microsoft Word automatically adjusts the formatting of tables when cutting and pasting selections. Read/write **Boolean** .


## Syntax

 _expression_ . **PasteAdjustTableFormatting**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Example

This example sets Word to automatically adjust the formatting of tables when cutting and pasting if the option has been disabled.


```vb
Sub AdjustTableFormatting() 
 With Options 
 If .PasteAdjustTableFormatting = False Then 
 .PasteAdjustTableFormatting = True 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

