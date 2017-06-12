---
title: KeyBindings.ClearAll Method (Word)
keywords: vbawd10.chm160825446
f1_keywords:
- vbawd10.chm160825446
ms.prod: word
api_name:
- Word.KeyBindings.ClearAll
ms.assetid: d03f9e7e-12e6-940b-d0f4-7d83e098eb05
ms.date: 06/08/2017
---


# KeyBindings.ClearAll Method (Word)

Clears all the customized key assignments and restores the original Microsoft Word shortcut key assignments.


## Syntax

 _expression_ . **ClearAll**

 _expression_ A variable that represents a **[KeyBindings](keybindings-object-word.md)** collection.


## Example

This example clears the customized key assignments in the Normal template. The key assignments are reset to the default settings.


```
CustomizationContext = NormalTemplateKeyBindings.ClearAll
```


## See also


#### Concepts


[KeyBindings Collection Object](keybindings-object-word.md)

