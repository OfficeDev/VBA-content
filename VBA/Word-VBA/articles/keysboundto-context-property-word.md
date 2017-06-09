---
title: KeysBoundTo.Context Property (Word)
keywords: vbawd10.chm160890890
f1_keywords:
- vbawd10.chm160890890
ms.prod: word
api_name:
- Word.KeysBoundTo.Context
ms.assetid: 9d5b2bf6-8cc5-eee8-bc3e-eb4b272b1775
ms.date: 06/08/2017
---


# KeysBoundTo.Context Property (Word)

Returns an  **Object** that represents the storage location of the specified key binding. Read-only.


## Syntax

 _expression_ . **Context**

 _expression_ A variable that represents a **[KeysBoundTo](keysboundto-object-word.md)** object.


## Remarks

This property can return a  **Document** , **Template** , or **Application** object. Built-in key assignments (for example, CTRL+I for **Italic** ) return the **Application** object as the context. Any key bindings you add will return a **Document** or **Template** object, depending on the customization context in effect when the **KeyBinding** object was added.


## Example

This example displays the name of the document or template where the macro named "Macro1" is stored.


```vb
Sub TestContext1() 
 Dim kbMacro1 As KeysBoundTo 
 
 Set kbMacro1 = KeysBoundTo(KeyCategory:=wdKeyCategoryMacro, _ 
 Command:="Macro1") 
 MsgBox kbMacro1.Context.Name 
End Sub
```


## See also


#### Concepts


[KeysBoundTo Collection Object](keysboundto-object-word.md)

