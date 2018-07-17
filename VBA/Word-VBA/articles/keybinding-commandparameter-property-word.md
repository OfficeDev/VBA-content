---
title: KeyBinding.CommandParameter Property (Word)
keywords: vbawd10.chm160956424
f1_keywords:
- vbawd10.chm160956424
ms.prod: word
api_name:
- Word.KeyBinding.CommandParameter
ms.assetid: 041cea6f-6275-e008-43ce-9075ee6d8f8c
ms.date: 06/08/2017
---


# KeyBinding.CommandParameter Property (Word)

Returns the command parameter assigned to the specified shortcut key. Read-only  **String** .


## Syntax

 _expression_ . **CommandParameter**

 _expression_ A variable that represents a **[KeyBinding](keybinding-object-word.md)** object.


## Remarks

For information about commands that take parameters, see the  **[Add](keybindings-add-method-word.md)** method. Use the **Command** property to return the command name assigned to the specified shortcut key.


## Example

This example assigns a shortcut key to the FontSize command, with a command parameter of 8. Use the CommandParameter property to display the command parameter along with the command name and key string.


```vb
Dim kbNew As KeyBinding 
 
Set kbNew = KeyBindings.Add(KeyCategory:=wdKeyCategoryCommand, _ 
 Command:="FontSize", _ 
 KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyS), _ 
 CommandParameter:="8") 
MsgBox kbNew.Command &; Chr$(32) &; kbNew.CommandParameter _ 
 &; vbCr &; kbNew.KeyString
```


## See also


#### Concepts


[KeyBinding Object](keybinding-object-word.md)

