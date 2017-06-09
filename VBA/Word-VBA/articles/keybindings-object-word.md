---
title: KeyBindings Object (Word)
ms.prod: word
ms.assetid: d2e38b04-b7e1-b35c-e511-5988d132b074
ms.date: 06/08/2017
---


# KeyBindings Object (Word)

A collection of  **[KeyBinding](keybinding-object-word.md)** objects that represent the custom key assignments in the current context. Custom key assignments are made in the **Customize Keyboard** dialog box.


## Remarks

Use the  **[KeyBindings](application-keybindings-property-word.md)** property to return the **KeyBindings** collection. The following example inserts after the selection the command name and key combination for each item in the **KeyBindings** collection.


```vb
CustomizationContext = NormalTemplate 
For Each aKey In KeyBindings 
 Selection.InsertAfter aKey.Command &; vbTab _ 
 &; aKey.KeyString &; vbCr 
 Selection.Collapse Direction:=wdCollapseEnd 
Next aKey
```

Use the  **Add** method to add a **KeyBinding** object to the **KeyBindings** collection. The following example adds the CTRL+ALT+H key combination to the Heading 1 style in the active document.




```
CustomizationContext = ActiveDocument 
KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _ 
 Command:="Heading 1", _ 
 KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyH)
```

Use  **KeyBindings** (Index), where Index is the index number, to return a single **KeyBinding** object. The following example displays the command associated with the first **KeyBinding** object in the **KeyBindings** collection.




```vb
MsgBox KeyBindings(1).Command
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


