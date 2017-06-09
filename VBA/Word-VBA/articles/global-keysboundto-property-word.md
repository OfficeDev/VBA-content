---
title: Global.KeysBoundTo Property (Word)
keywords: vbawd10.chm163119174
f1_keywords:
- vbawd10.chm163119174
ms.prod: word
api_name:
- Word.Global.KeysBoundTo
ms.assetid: 6347357b-1c63-eea2-d236-2c5f182eb676
ms.date: 06/08/2017
---


# Global.KeysBoundTo Property (Word)

Returns a  **KeysBoundTo** object that represents all the key combinations assigned to the specified item.


## Syntax

 _expression_ . **KeysBoundTo**( **_KeyCategory_** , **_Command_** , **_CommandParameter_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCategory_|Required| **WdKeyCategory**|The category of the key combination.|
| _Command_|Required| **String**|The name of the command.|
| _CommandParameter_|Optional| **Variant**|Additional text, if any, required for the command specified by Command. For more information, see the "Remarks" section in the  **[Add](keybindings-add-method-word.md)** method for the **[KeyBindings](keybindings-object-word.md)** object.|

## Example

This example displays all the key combinations assigned to the FileOpen command in the template attached to the active document.


```vb
Dim kbLoop As KeyBinding 
Dim strOutput As String 
 
CustomizationContext = ActiveDocument.AttachedTemplate 
 
For Each kbLoop In _ 
 KeysBoundTo(KeyCategory:=wdKeyCategoryCommand, _ 
 Command:="FileOpen") 
 strOutput = strOutput &; kbLoop.KeyString &; vbCr 
Next kbLoop 
 
MsgBox strOutput
```

This example removes all key assignments from Macro1 in the Normal template.




```vb
Dim aKey As KeyBinding 
 
CustomizationContext = NormalTemplate 
For Each aKey In _ 
 KeysBoundTo(KeyCategory:=wdKeyCategoryMacro, _ 
 Command:="Macro1") 
 aKey.Disable 
Next aKey
```


## See also


#### Concepts


[Global Object](global-object-word.md)

