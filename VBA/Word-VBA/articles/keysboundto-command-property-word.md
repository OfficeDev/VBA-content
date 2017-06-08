---
title: KeysBoundTo.Command Property (Word)
keywords: vbawd10.chm160890884
f1_keywords:
- vbawd10.chm160890884
ms.prod: word
api_name:
- Word.KeysBoundTo.Command
ms.assetid: a8c8a12b-5dce-5103-9309-b0cb36042b80
ms.date: 06/08/2017
---


# KeysBoundTo.Command Property (Word)

Returns a  **String** that represents the command assigned to the specified key combination. Read-only.


## Syntax

 _expression_ . **Command**

 _expression_ A variable that represents a **[KeysBoundTo](keysboundto-object-word.md)** object.


## Example

This example displays the keys assigned to font names. A message is displayed if no keys have been assigned to fonts.


```vb
Dim kbLoop As KeyBinding 
 
For Each kbLoop In KeyBindings 
 If kbLoop.KeyCategory = wdKeyCategoryFont Then 
 Count = Count + 1 
 MsgBox kbLoop.Command &; vbCr &; kbLoop.KeyString 
 End If 
Next kbLoop 
 
If Count = 0 Then MsgBox "Keys haven't been assigned to fonts"
```


## See also


#### Concepts


[KeysBoundTo Collection Object](keysboundto-object-word.md)

