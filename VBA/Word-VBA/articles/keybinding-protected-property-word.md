---
title: KeyBinding.Protected Property (Word)
keywords: vbawd10.chm160956419
f1_keywords:
- vbawd10.chm160956419
ms.prod: word
api_name:
- Word.KeyBinding.Protected
ms.assetid: 7f56f218-178d-5c98-9c6b-05d228e48ff3
ms.date: 06/08/2017
---


# KeyBinding.Protected Property (Word)

 **True** if you cannot change the specified key binding in the **Customize Keyboard** dialog box. Read-only **Boolean** .


## Syntax

 _expression_ . **Protected**

 _expression_ An expression that returns a **[KeyBinding](keybinding-object-word.md)** object.


## Remarks

You can access the  **Customize Keyboard** dialog box from the **Tools** menu; click **Customize**, and then click the  **Keyboard** button.

Use the  **[Add](keybindings-add-method-word.md)** method of the **[KeyBindings](keybindings-object-word.md)** object to add a key binding regardless of the protected status.


## Example

This example displays the protection status for the CTRL+S key binding.


```
CustomizationContext = ActiveDocument.AttachedTemplate 
MsgBox FindKey(BuildKeyCode(wdKeyControl, wdKeyS)).Protected
```

This example displays a message if the A key binding is protected.




```vb
CustomizationContext = NormalTemplate 
If FindKey(BuildKeyCode(wdKeyA)).Protected = True Then 
 MsgBox "The A key is protected" 
End If
```


## See also


#### Concepts


[KeyBinding Object](keybinding-object-word.md)

