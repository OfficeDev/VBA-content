---
title: Options.CtrlClickHyperlinkToOpen Property (Word)
keywords: vbawd10.chm162988467
f1_keywords:
- vbawd10.chm162988467
ms.prod: word
api_name:
- Word.Options.CtrlClickHyperlinkToOpen
ms.assetid: 2180e99c-ab4c-3f75-2417-22cec6b2d130
ms.date: 06/08/2017
---


# Options.CtrlClickHyperlinkToOpen Property (Word)

 **True** if Microsoft Word requires holding down the CTRL key while clicking to open a hyperlink. Read/write **Boolean** .


## Syntax

 _expression_ . **CtrlClickHyperlinkToOpen**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example disables the option that requires holding down the CTRL key while clicking hyperlinks to open them.


```vb
Sub ToggleHyperlinkOption() 
 Options.CtrlClickHyperlinkToOpen = False 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

