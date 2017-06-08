---
title: Global.KeyBindings Property (Word)
keywords: vbawd10.chm163119173
f1_keywords:
- vbawd10.chm163119173
ms.prod: word
api_name:
- Word.Global.KeyBindings
ms.assetid: 76b3fb80-9169-06b6-8aa6-d70d960ea2f8
ms.date: 06/08/2017
---


# Global.KeyBindings Property (Word)

Returns a  **KeyBindings** collection that represents customized key assignments, which include a key code, a key category, and a command.


## Syntax

 _expression_ . **KeyBindings**

 _expression_ Required. A variable that represents a **[Global](global-object-word.md)** object.


## Example

This example assigns the CTRL+ALT+W key combination to the FileClose command. This keyboard customization is saved in the Normal template.


```
CustomizationContext = NormalTemplate 
KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, _ 
 wdKeyW), KeyCategory:=wdKeyCategoryCommand, _ 
 Command:="FileClose"
```

This example inserts the command name and key combination string for each item in the  **KeyBindings** collection.




```vb
Dim kbLoop As KeyBinding 
 
CustomizationContext = NormalTemplate 
For Each kbLoop In KeyBindings 
 Selection.InsertAfter kbLoop.Command &; vbTab _ 
 &; kbLoop.KeyString &; vbCr 
 Selection.Collapse Direction:=wdCollapseEnd 
Next kbLoop
```


## See also


#### Concepts


[Global Object](global-object-word.md)

