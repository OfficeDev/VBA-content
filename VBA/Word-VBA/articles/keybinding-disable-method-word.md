---
title: KeyBinding.Disable Method (Word)
keywords: vbawd10.chm160956518
f1_keywords:
- vbawd10.chm160956518
ms.prod: word
api_name:
- Word.KeyBinding.Disable
ms.assetid: 07463e08-1802-0f1b-7c3f-408f072386b5
ms.date: 06/08/2017
---


# KeyBinding.Disable Method (Word)

Removes the specified key combination if it is currently assigned to a command. After you use this method, the key combination has no effect.


## Syntax

 _expression_ . **Disable**

 _expression_ Required. A variable that represents a **[KeyBinding](keybinding-object-word.md)** object.


## Remarks

Using this method is the equivalent to clicking the  **Remove** button in the **Customize Keyboard** dialog box. Use the **Clear** method with a **KeyBinding** object to reset a built-in command to its default key assignment. You don't need to remove or rebind a **KeyBinding** object before adding it elsewhere.


## Example

This example removes the CTRL+SHIFT+B key assignment. This key combination is assigned to the Bold command by default.


```
CustomizationContext = NormalTemplate 
FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyB)).Disable
```

This example assigns the CTRL+SHIFT+O key combination to the  **Organizer** command. The example then uses the Disable method to remove the CTRL+SHIFT+O key combination and displays a message.




```vb
CustomizationContext = NormalTemplate 
KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyO, _ 
 wdKeyControl, wdKeyShift), _ 
 KeyCategory:=wdKeyCategoryCommand, Command:="Organizer" 
With FindKey(BuildKeyCode(wdKeyO, wdKeyControl, wdKeyShift)) 
 MsgBox .Command &; " is assigned to CTRL+Shift+O" 
 .Disable 
 If .Command = "" Then MsgBox _ 
 "Nothing is assigned to CTRL+Shift+O" 
End With
```

This example removes all key assignments for the global macro named "Macro1."




```vb
Dim kbLoop As KeyBinding 
 
CustomizationContext = NormalTemplate 
For Each kbLoop In KeysBoundTo _ 
 (KeyCategory:=wdKeyCategoryMacro, Command:="Macro1") 
 kbLoop.Disable 
Next kbLoop
```


## See also


#### Concepts


[KeyBinding Object](keybinding-object-word.md)

