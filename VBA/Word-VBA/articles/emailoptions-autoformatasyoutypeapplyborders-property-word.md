---
title: EmailOptions.AutoFormatAsYouTypeApplyBorders Property (Word)
keywords: vbawd10.chm165347589
f1_keywords:
- vbawd10.chm165347589
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeApplyBorders
ms.assetid: 24e0f74e-b57e-d81c-2e36-c08adb59bf5e
ms.date: 06/08/2017
---


# EmailOptions.AutoFormatAsYouTypeApplyBorders Property (Word)

 **True** if a series of three or more hyphens (-), equal signs (=), or underscore characters (_) are automatically replaced by a specific border line when the ENTER key is pressed. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoFormatAsYouTypeApplyBorders**

 _expression_ A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Remarks

Hyphens (-) are replaced by a 0.75-point line, equal signs (=) are replaced by a 0.75-point double line, and underscore characters (_) are replaced by a 1.5-point line.


## Example

This example causes sequences of three or more hyphens (-), equal signs (=), or underscore characters (_) to be transformed into borders.


```vb
Options.AutoFormatAsYouTypeApplyBorders = True
```

This example returns the current setting for the  **Borders** option on the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```vb
MsgBox Options.AutoFormatAsYouTypeApplyBorders
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

