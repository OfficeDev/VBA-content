---
title: EmailOptions.UseThemeStyle Property (Word)
keywords: vbawd10.chm165347431
f1_keywords:
- vbawd10.chm165347431
ms.prod: word
api_name:
- Word.EmailOptions.UseThemeStyle
ms.assetid: e34f27c6-4222-aa9a-dfbc-40c7c5c55a67
ms.date: 06/08/2017
---


# EmailOptions.UseThemeStyle Property (Word)

 **True** if new e-mail messages use the character style defined by the default e-mail message theme. Read/write **Boolean** .


## Syntax

 _expression_ . **UseThemeStyle**

 _expression_ A variable that represents a **[EmailOptions](emailoptions-object-word.md)** object.


## Remarks

If no default e-mail message theme has been specified, the  **UseThemeStyle** property has no effect.


## Example

This example sets Microsoft Word to use the Artsy theme as the default theme for new e-mail messages and to use the character style defined in the Artsy theme.


```vb
Application.EmailOptions.ThemeName = "artsy" 
Application.EmailOptions.UseThemeStyle = True
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

