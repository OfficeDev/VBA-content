---
title: OlkTimeZoneControl.KeyUp Event (Outlook)
keywords: vbaol11.chm1000525
f1_keywords:
- vbaol11.chm1000525
ms.prod: outlook
api_name:
- Outlook.OlkTimeZoneControl.KeyUp
ms.assetid: 06869fbe-73dc-fd0f-0a6f-59505e0e80f8
ms.date: 06/08/2017
---


# OlkTimeZoneControl.KeyUp Event (Outlook)

Occurs when the user releases a key.


## Syntax

 _expression_ . **KeyUp**( **_KeyCode_** , **_Shift_** )

 _expression_ A variable that represents an **OlkTimeZoneControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **OlShiftState**|A bitwise-OR mask of constants in the  **[OlShiftState](olshiftstate-enumeration-outlook.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys ( **SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyUp** event is accessible through the _Shift_ parameter.


## See also


#### Concepts


[OlkTimeZoneControl Object](olktimezonecontrol-object-outlook.md)

