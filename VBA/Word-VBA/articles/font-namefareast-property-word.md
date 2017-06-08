---
title: Font.NameFarEast Property (Word)
keywords: vbawd10.chm156369052
f1_keywords:
- vbawd10.chm156369052
ms.prod: word
api_name:
- Word.Font.NameFarEast
ms.assetid: 8df21c3e-5f50-37ca-cde9-27f5b0551f8e
ms.date: 06/08/2017
---


# Font.NameFarEast Property (Word)

Returns or sets an East Asian font name. Read/write  **String** .


## Syntax

 _expression_ . **NameFarEast**

 _expression_ An expression that returns a **[Font](font-object-word.md)** object.


## Remarks

In the U.S. English version of Microsoft Word, the default value of this property is Times New Roman. This is the recommended way to return or set the font for Asian text in a document created in an Asian version of Word.


## Example

This example displays the East Asian font name that's applied to the selection.


```vb
MsgBox Selection.Font.NameFarEast
```


## See also


#### Concepts


[Font Object](font-object-word.md)

