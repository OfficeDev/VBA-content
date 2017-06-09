---
title: KeyBinding.Command Property (Word)
keywords: vbawd10.chm160956417
f1_keywords:
- vbawd10.chm160956417
ms.prod: word
api_name:
- Word.KeyBinding.Command
ms.assetid: 0693cc28-7498-03c6-0e24-53f78924db1e
ms.date: 06/08/2017
---


# KeyBinding.Command Property (Word)

Returns the command assigned to the specified key combination. Read-only  **String** .


## Syntax

 _expression_ . **Command**

 _expression_ A variable that represents a **[KeyBinding](keybinding-object-word.md)** object.


## Example

This example displays the keys assigned to font names. A message is displayed if no keys have been assigned to fonts.


```vb
Dim kbLoop As KeyBinding 
 
For Each kbLoop In KeyBindings 
 If kbLoop.KeyCategory = wdKeyCategoryFont Then 
 Count = Count + 1 
 MsgBox kbLoop.Command &; vbCr &; kbLoop.KeyString 
 End If 
Next kbLoop 
 
If Count = 0 Then MsgBox "Keys haven't been assigned to fonts"
```


## See also


#### Concepts


[KeyBinding Object](keybinding-object-word.md)

