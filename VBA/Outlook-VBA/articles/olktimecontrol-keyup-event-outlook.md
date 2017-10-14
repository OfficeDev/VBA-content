---
title: OlkTimeControl.KeyUp Event (Outlook)
keywords: vbaol11.chm1000410
f1_keywords:
- vbaol11.chm1000410
ms.prod: outlook
api_name:
- Outlook.OlkTimeControl.KeyUp
ms.assetid: b2ff348b-6c94-09b3-e8ee-8eb25ac15ba0
ms.date: 06/08/2017
---


# OlkTimeControl.KeyUp Event (Outlook)

Occurs when the user releases a key.


## Syntax

 _expression_ . **KeyUp**( **_KeyCode_** , **_Shift_** )

 _expression_ A variable that represents an **OlkTimeControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](olshiftstate-enumeration-outlook.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys ( **SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyUp** event is accessible through the _Shift_ parameter.


## See also


#### Concepts


[OlkTimeControl Object](olktimecontrol-object-outlook.md)

