---
title: KeyBinding.KeyCategory Property (Word)
keywords: vbawd10.chm160956420
f1_keywords:
- vbawd10.chm160956420
ms.prod: word
api_name:
- Word.KeyBinding.KeyCategory
ms.assetid: 293371f6-7057-b579-b039-13e762f5ea62
ms.date: 06/08/2017
---


# KeyBinding.KeyCategory Property (Word)

Returns the type of item assigned to the specified key binding. Read-only  **WdKeyCategory** .


## Syntax

 _expression_ . **KeyCategory**

 _expression_ Required. A variable that represents a **[KeyBinding](keybinding-object-word.md)** object.


## Example

This example displays the keys assigned to font names. A message is displayed if no keys have been assigned to fonts.


```vb
Dim kbLoop As KeyBinding 
Dim intCount As Integer 
 
intCount = 0 
 
For Each kbLoop In KeyBindings 
 If kbLoop.KeyCategory = wdKeyCategoryFont Then 
 intCount = intCount + 1 
 MsgBox kbLoop.Command &; vbCr &; kbLoop.KeyString 
 End If 
Next kbLoop 
 
If intCount = 0 Then _ 
 MsgBox "Keys haven't been assigned to fonts"
```


## See also


#### Concepts


[KeyBinding Object](keybinding-object-word.md)

