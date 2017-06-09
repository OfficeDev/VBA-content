---
title: Hyperlink.ScreenTip Property (Word)
keywords: vbawd10.chm161285107
f1_keywords:
- vbawd10.chm161285107
ms.prod: word
api_name:
- Word.Hyperlink.ScreenTip
ms.assetid: 59df269f-3dfd-53fe-b4ac-7889eefef740
ms.date: 06/08/2017
---


# Hyperlink.ScreenTip Property (Word)

Returns or sets the text that appears as a ScreenTip when the mouse pointer is positioned over the specified hyperlink. Read/write  **String** .


## Syntax

 _expression_ . **ScreenTip**

 _expression_ An expression that returns a **[Hyperlink](hyperlink-object-word.md)** object.


## Example

This example sets the ScreenTip text for the first hyperlink in the active document.


```vb
ActiveDocument.Hyperlinks(1).ScreenTip = _ 
 "Home"
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-word.md)

