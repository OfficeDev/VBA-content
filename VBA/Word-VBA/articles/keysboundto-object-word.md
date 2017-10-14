---
title: KeysBoundTo Object (Word)
keywords: vbawd10.chm2455
f1_keywords:
- vbawd10.chm2455
ms.prod: word
ms.assetid: 63ed40e5-8223-78d6-c90a-bf6be8a2fbf6
ms.date: 06/08/2017
---


# KeysBoundTo Object (Word)

A collection of  **[KeyBinding](keybinding-object-word.md)** objects assigned to a command, style, macro, or other item in the current context.


## Remarks

Use the  **[KeysBoundTo](application-keysboundto-property-word.md)** property to return the **KeysBoundTo** collection. The following example displays the key combinations assigned to the **FileNew** command in the Normal template.


```vb
CustomizationContext = NormalTemplate 
For Each myKey In KeysBoundTo(KeyCategory:=wdKeyCategoryCommand, _ 
 Command:="FileNew") 
 myStr = myStr &; myKey.KeyString &; vbCr 
Next myKey 
MsgBox myStr
```

The following example displays the name of the document or template where the keys for the macro named "Macro1" are stored.




```vb
Set kb = KeysBoundTo(KeyCategory:=wdKeyCategoryMacro, _ 
 Command:="Macro1") 
MsgBox kb.Context.Name
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


