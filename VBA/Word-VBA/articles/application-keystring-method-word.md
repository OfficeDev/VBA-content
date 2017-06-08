---
title: Application.KeyString Method (Word)
keywords: vbawd10.chm158335293
f1_keywords:
- vbawd10.chm158335293
ms.prod: word
api_name:
- Word.Application.KeyString
ms.assetid: 20525053-3cf8-bdf8-cb67-cca39bf2b30c
ms.date: 06/08/2017
---


# Application.KeyString Method (Word)

Returns the key combination string for the specified keys (for example, CTRL+SHIFT+A).


## Syntax

 _expression_ . **KeyString**( **_KeyCode_** , **_KeyCode2_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|A key you specify by using one of the  **WdKey** constants.|
| _KeyCode2_|Optional| **Variant**|A second key you specify by using one of the  **WdKey** constants.|

### Return Value

String


## Remarks

You can use the  **BuildKeyCode** method to create the KeyCode or KeyCode2 argument.


## Example

This example displays the key combination string (CTRL+SHIFT+A) for the following  **WdKey** constants: **wdKeyControl** , **wdKeyShift** , and **wdKeyA** .


```
CustomizationContext = ActiveDocument.AttachedTemplate 
MsgBox KeyString(KeyCode:=BuildKeyCode(wdKeyControl, _ 
 wdKeyShift, wdKeyA))
```


## See also


#### Concepts


[Application Object](application-object-word.md)

