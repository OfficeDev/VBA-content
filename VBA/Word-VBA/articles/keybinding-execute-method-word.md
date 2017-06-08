---
title: KeyBinding.Execute Method (Word)
keywords: vbawd10.chm160956519
f1_keywords:
- vbawd10.chm160956519
ms.prod: word
api_name:
- Word.KeyBinding.Execute
ms.assetid: ea8df8eb-50dc-307b-ea1a-ba5e6a5c683f
ms.date: 06/08/2017
---


# KeyBinding.Execute Method (Word)

Runs the command associated with the specified key combination.


## Syntax

 _expression_ . **Execute**

 _expression_ Required. A variable that represents a **[KeyBinding](keybinding-object-word.md)** object.


## Example

This example assigns the CTRL+SHIFT+C key combination to the  **FileClose** command and then executes the key combination (the document is closed).


```
CustomizationContext = ActiveDocument.AttachedTemplate 
Keybindings.Add KeyCode:=BuildKeyCode(wdKeyControl, _ 
 wdKeyShift, wdKeyC), KeyCategory:=wdKeyCategoryCommand, _ 
 Command:="FileClose" 
FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyC)).Execute
```


## See also


#### Concepts


[KeyBinding Object](keybinding-object-word.md)

