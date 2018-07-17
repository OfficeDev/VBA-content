---
title: KeyBinding.KeyString Property (Word)
keywords: vbawd10.chm160956418
f1_keywords:
- vbawd10.chm160956418
ms.prod: word
api_name:
- Word.KeyBinding.KeyString
ms.assetid: 2ee7b80c-e923-7b0a-81f3-d807b38cba4e
ms.date: 06/08/2017
---


# KeyBinding.KeyString Property (Word)

Returns the key combination string for the specified keys (for example, CTRL+SHIFT+A). Read-only  **String** .


## Syntax

 _expression_ . **KeyString**

 _expression_ Required. A variable that represents a **[KeyBinding](keybinding-object-word.md)** object.


## Example

This example displays the key combination string for the first customized key combination in the Normal template.


```vb
CustomizationContext = NormalTemplate 
If KeyBindings.Count >= 1 Then 
 MsgBox KeyBindings(1).KeyString 
End If
```

This example displays a message if the  **KeyBindings** collection includes the ALT+CTRL+W key combination.




```vb
Dim aCode As Long 
Dim aKey As KeyBinding 
 
CustomizationContext = NormalTemplate 
aCode = BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyW) 
For Each aKey In KeyBindings 
 If aCode = aKey.KeyCode Then 
 MsgBox aKey.KeyString &; " is already in use" 
 End If 
Next aKey
```


## See also


#### Concepts


[KeyBinding Object](keybinding-object-word.md)

