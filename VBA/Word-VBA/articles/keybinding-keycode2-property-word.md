---
title: KeyBinding.KeyCode2 Property (Word)
keywords: vbawd10.chm160956423
f1_keywords:
- vbawd10.chm160956423
ms.prod: word
api_name:
- Word.KeyBinding.KeyCode2
ms.assetid: b041fb3f-1777-f56a-4808-f96e570f5440
ms.date: 06/08/2017
---


# KeyBinding.KeyCode2 Property (Word)

Returns a unique number for the second key in the specified key binding. Read-only  **Long** .


## Syntax

 _expression_ . **KeyCode2**

 _expression_ An expression that returns a **[KeyBinding](keybinding-object-word.md)** object.


## Example

This example displays the key codes of each key in the  **KeyBindings** collection (the collection of all the customized keys in the active document).


```vb
Dim aKey As KeyBinding 
 
CustomizationContext = ActiveDocument 
For Each aKey In KeyBindings 
 If aKey.KeyCode2 <> wdNoKey Then 
 MsgBox aKey.KeyString &; vbCr _ 
 &; "KeyCode1 = " &; aKey.KeyCode &; vbCr _ 
 &; "KeyCode2 = " &; aKey.KeyCode2 
 Else 
 MsgBox aKey.KeyString &; vbCr _ 
 &; "KeyCode1 = " &; aKey.KeyCode 
 End If 
Next aKey
```


## See also


#### Concepts


[KeyBinding Object](keybinding-object-word.md)

