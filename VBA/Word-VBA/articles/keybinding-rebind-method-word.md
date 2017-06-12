---
title: KeyBinding.Rebind Method (Word)
keywords: vbawd10.chm160956520
f1_keywords:
- vbawd10.chm160956520
ms.prod: word
api_name:
- Word.KeyBinding.Rebind
ms.assetid: edc938ff-5ee5-3134-5808-a861ef37a2da
ms.date: 06/08/2017
---


# KeyBinding.Rebind Method (Word)

Changes the command assigned to the specified key binding.


## Syntax

 _expression_ . **Rebind**( **_KeyCategory_** , **_Command_** , **_CommandParameter_** )

 _expression_ Required. A variable that represents a **[KeyBinding](keybinding-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCategory_|Required| **WdKeyCategory**|The key category of the specified key binding.|
| _Command_|Required| **String**|The name of the specified command.|
| _CommandParameter_|Optional| **Variant**|Additional text, if any, required for the command specified by Command. For information about values for this argument, see the  **[Add](keybindings-add-method-word.md)** method.|

## Example

This example reassigns the CTRL+SHIFT+S key binding to the  **FileSaveAs** command.


```vb
Dim kbTemp As KeyBinding 
 
CustomizationContext = NormalTemplate 
Set kbTemp = _ 
 FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyS)) 
kbTemp.Rebind KeyCategory:=wdKeyCategoryCommand, _ 
 Command:="FileSaveAs"
```

This example rebinds all keys assigned to the macro named "Macro1" to the macro named "ReportMacro."




```vb
Dim kbLoop As KeyBinding 
 
CustomizationContext = ActiveDocument.AttachedTemplate 
For Each kbLoop In _ 
 KeysBoundTo(KeyCategory:=wdKeyCategoryMacro, _ 
 Command:="Macro1") 
 kbLoop.Rebind KeyCategory:=wdKeyCategoryMacro, _ 
 Command:="ReportMacro" 
Next kbLoop
```


## See also


#### Concepts


[KeyBinding Object](keybinding-object-word.md)

