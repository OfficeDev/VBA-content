---
title: KeyBinding.Clear Method (Word)
keywords: vbawd10.chm160956517
f1_keywords:
- vbawd10.chm160956517
ms.prod: word
api_name:
- Word.KeyBinding.Clear
ms.assetid: 7f53f149-71e9-e2ff-c261-31cd1f0668de
ms.date: 06/08/2017
---


# KeyBinding.Clear Method (Word)

Removes the specified key binding from the  **KeyBindings** collection and resets a built-in command to its default key assignment.


## Syntax

 _expression_ . **Clear**

 _expression_ A variable that represents a **[KeyBinding](keybinding-object-word.md)** object.


## Example

This example removes the ALT+F1 key assignment from the Normal template.


```
CustomizationContext = NormalTemplateFindKey(BuildKeyCode(Arg1:=wdKeyAlt, Arg2:=wdKeyF1)).Clear
```


## See also


#### Concepts


[KeyBinding Object](keybinding-object-word.md)

