---
title: EmailOptions.ComposeStyle Property (Word)
keywords: vbawd10.chm165347437
f1_keywords:
- vbawd10.chm165347437
ms.prod: word
api_name:
- Word.EmailOptions.ComposeStyle
ms.assetid: 0c1ada5e-7bf0-2ae1-3223-ed4f76252bb1
ms.date: 06/08/2017
---


# EmailOptions.ComposeStyle Property (Word)

Returns a  **[Style](style-object-word.md)** object that represents the style used to compose new e-mail messages. Read-only.


## Syntax

 _expression_ . **ComposeStyle**

 _expression_ A variable that represents a **[EmailOptions](emailoptions-object-word.md)** object.


## Example

This example displays the name of the default style used to compose new e-mail messages.


```vb
MsgBox Application.EmailOptions.ComposeStyle.NameLocal
```

This example changes the font color of the default style used to compose new e-mail messages.




```vb
Application.EmailOptions.ComposeStyle.Font.Color = _ 
 wdColorBrightGreen
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

