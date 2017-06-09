---
title: EmailOptions.ReplyStyle Property (Word)
keywords: vbawd10.chm165347438
f1_keywords:
- vbawd10.chm165347438
ms.prod: word
api_name:
- Word.EmailOptions.ReplyStyle
ms.assetid: adb778ca-8943-4f30-48d8-98336ea81ea7
ms.date: 06/08/2017
---


# EmailOptions.ReplyStyle Property (Word)

Returns a  **[Style](style-object-word.md)** object that represents the style used when replying to e-mail messages.


## Syntax

 _expression_ . **ReplyStyle**

 _expression_ An expression that returns an **[EmailOptions](emailoptions-object-word.md)** object.


## Example

This example displays the name of the default style used when replying to e-mail messages.


```vb
MsgBox Application.EmailOptions.ReplyStyle.NameLocal
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

