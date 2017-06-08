---
title: Global.KeyString Method (Word)
keywords: vbawd10.chm163119421
f1_keywords:
- vbawd10.chm163119421
ms.prod: word
api_name:
- Word.Global.KeyString
ms.assetid: 4ad72e74-d26d-093e-8404-b3dc10ebe1f0
ms.date: 06/08/2017
---


# Global.KeyString Method (Word)

Returns the key combination string for the specified keys (for example, CTRL+SHIFT+A).


## Syntax

 _expression_ . **KeyString**( **_KeyCode_** , **_KeyCode2_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object.


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


[Global Object](global-object-word.md)

