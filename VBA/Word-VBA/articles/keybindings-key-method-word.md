---
title: KeyBindings.Key Method (Word)
keywords: vbawd10.chm160825454
f1_keywords:
- vbawd10.chm160825454
ms.prod: word
api_name:
- Word.KeyBindings.Key
ms.assetid: 0e20a18e-7812-8d99-3c4d-4d3e9e661d16
ms.date: 06/08/2017
---


# KeyBindings.Key Method (Word)

Returns a  **KeyBinding** object that represents the specified custom key combination.


## Syntax

 _expression_ . **Key**( **_KeyCode_** , **_KeyCode2_** )

 _expression_ A variable that represents a **[KeyBindings](keybindings-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|A key you specify by using one of the  **WdKey** constants.|
| _KeyCode2_|Optional| **Variant**|A second key you specify by using one of the  **WdKey** constants.|

### Return Value

KeyBinding


## Remarks

If the key combination doesn't exist, this method returns  **Nothing** .

You can use the  **BuildKeyCode** method to create the KeyCode or KeyCode2 argument.


## Example

This example assigns the ALT+F4 key combination to the Arial font and then displays the number of items in the  **KeyBindings** collection. The example then clears the key combinations (returns it to its default setting) and redisplays the number of items in the **KeyBindings** collection.


```
CustomizationContext = NormalTemplate 
KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF4), _ 
 KeyCategory:=wdKeyCategoryFont, Command:="Arial" 
MsgBox KeyBindings.Count &; " keys in KeyBindings collection" 
KeyBindings.Key(KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF4)).Clear 
MsgBox KeyBindings.Count &; " keys in KeyBindings collection"
```

This example assigns the CTRL+SHIFT+U key combination to the macro named "Macro1" in the active document. The example uses the  **Key** property to return a **KeyBinding** object so that Word can retrieve and display the command name.




```
CustomizationContext = ActiveDocument 
KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, _ 
 wdKeyShift, wdKeyU), KeyCategory:=wdKeyCategoryMacro, _ 
 Command:="Macro1" 
MsgBox KeyBindings.Key(BuildKeyCode(wdKeyControl, _ 
 wdKeyShift, wdKeyU)).Command
```

This example determines whether the CTRL+SHIFT+A key combination is part of the  **KeyBindings** collection.




```vb
Dim kbTemp As KeyBinding 
 
CustomizationContext = NormalTemplate 
Set kbTemp = KeyBindings.Key(BuildKeyCode(wdKeyControl, _ 
 wdKeyShift,wdKeyA)) 
If (kbTemp Is Nothing) Then MsgBox _ 
 "Key is not in the KeyBindings collection"
```


## See also


#### Concepts


[KeyBindings Collection Object](keybindings-object-word.md)

