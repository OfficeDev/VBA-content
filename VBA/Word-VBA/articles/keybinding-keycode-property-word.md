---
title: KeyBinding.KeyCode Property (Word)
keywords: vbawd10.chm160956422
f1_keywords:
- vbawd10.chm160956422
ms.prod: word
api_name:
- Word.KeyBinding.KeyCode
ms.assetid: 8ca07f1e-b60b-bc10-b1fe-cb0d7b890d33
ms.date: 06/08/2017
---


# KeyBinding.KeyCode Property (Word)

Returns a unique number for the first key in the specified key binding. Read-only  **Long** .


## Syntax

 _expression_ . **KeyCode**

 _expression_ An expression that returns a **[KeyBinding](keybinding-object-word.md)** object.


## Remarks

You create this number by using the  **[BuildKeyCode](application-buildkeycode-method-word.md)** method when you are adding key bindings by using the **[Add](keybindings-add-method-word.md)** method of the **[KeyBindings](keybindings-object-word.md)** object.


## Example

This example displays a message if the  **KeyBindings** collection includes the ALT+CTRL+W key combination.


```vb
Dim lngCode As Long 
Dim kbLoop As KeyBinding 
 
CustomizationContext = NormalTemplate 
lngCode = BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyW) 
For Each kbLoop In KeyBindings 
 If lngCode = kbLoop.KeyCode Then 
 MsgBox kbLoop.KeyString &; " is already in use" 
 End If 
Next kbLoop
```


## See also


#### Concepts


[KeyBinding Object](keybinding-object-word.md)

