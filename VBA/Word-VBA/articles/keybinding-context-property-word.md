---
title: KeyBinding.Context Property (Word)
keywords: vbawd10.chm160956426
f1_keywords:
- vbawd10.chm160956426
ms.prod: word
api_name:
- Word.KeyBinding.Context
ms.assetid: 39612af3-b8b4-ab4d-3c83-35d1cf76f418
ms.date: 06/08/2017
---


# KeyBinding.Context Property (Word)

Returns an  **Object** that represents the storage location of the specified key binding. Read-only.


## Syntax

 _expression_ . **Context**

 _expression_ A variable that represents a **[KeyBinding](keybinding-object-word.md)** object.


## Remarks

This property can return a  **Document** , **Template** , or **Application** object. Built-in key assignments (for example, CTRL+I for **Italic** ) return the **Application** object as the context. Any key bindings you add will return a **Document** or **Template** object, depending on the customization context in effect when the **KeyBinding** object was added.


## Example

This example adds the F2 key to the Italic command and then uses the For Each...Next loop to display the keys assigned to the Italic command along with the context.


```vb
Dim kbLoop As KeyBinding 
 
CustomizationContext = NormalTemplate 
KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, _ 
 Command:="Italic", KeyCode:=wdKeyF2 
For Each kbLoop In _ 
 KeysBoundTo(KeyCategory:=wdKeyCategoryCommand, _ 
 Command:="Italic") 
 MsgBox kbLoop.KeyString &; vbCr &; kbLoop.Context.Name 
Next kbLoop
```


## See also


#### Concepts


[KeyBinding Object](keybinding-object-word.md)

