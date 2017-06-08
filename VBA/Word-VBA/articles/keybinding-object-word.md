---
title: KeyBinding Object (Word)
keywords: vbawd10.chm2456
f1_keywords:
- vbawd10.chm2456
ms.prod: word
api_name:
- Word.KeyBinding
ms.assetid: 0f691196-76ef-135d-a8c9-b2fb9f9ac695
ms.date: 06/08/2017
---


# KeyBinding Object (Word)

Represents a custom key assignment in the current context. The  **KeyBinding** object is a member of the **KeyBindings** collection.


## Remarks

Use  **KeyBindings** (Index), where Index is the index number, to return a single **KeyBinding** object. The following example displays the command associated with the first **KeyBinding** object in the **[KeyBindings](keybindings-object-word.md)** collection.


```vb
MsgBox KeyBindings(1).Command
```

You can also use the  **FindKey** property and the **Key** method to return a **KeyBinding** object.


 **Note**  Custom key assignments are made in the  **Customize Keyboard** dialog box.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


