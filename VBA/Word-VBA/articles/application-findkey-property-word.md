---
title: Application.FindKey Property (Word)
keywords: vbawd10.chm158335047
f1_keywords:
- vbawd10.chm158335047
ms.prod: word
api_name:
- Word.Application.FindKey
ms.assetid: f648e9a5-626b-3923-46e4-a0c9c1dfc815
ms.date: 06/08/2017
---


# Application.FindKey Property (Word)

Returns a  **[KeyBinding](keybinding-object-word.md)** object that represents the specified key combination. Read-only.


## Syntax

 _expression_ . **FindKey**( **_KeyCode_** , **_ KeyCode2_** )

 _expression_ Optional. An expression that returns an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|A key you specify by using one of the  **WdKey** constants.|
| _KeyCode2_|Optional| **Variant**|A second key you specify by using one of the  **WdKey** constants.|

## Remarks

You can use the  **[BuildKeyCode](application-buildkeycode-method-word.md)** method to create the _KeyCode_ or _KeyCode2_ argument.


## Example

This example disables the ALT+SHIFT+F12 key combination in the template attached to the active document. To return a  **KeyBinding** object that includes more than two keys, use the **BuildKeyCode** method, as shown in the example.


```
CustomizationContext = ActiveDocument.AttachedTemplate 
FindKey(KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, _ 
 wdKeyF12)).Disable
```

This example displays the command assigned to the F1 key.




```
CustomizationContext = NormalTemplate 
MsgBox FindKey(KeyCode:=wdKeyF1).Command
```


## See also


#### Concepts


[Application Object](application-object-word.md)

