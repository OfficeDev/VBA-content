---
title: OlkDateControl.KeyDown Event (Outlook)
keywords: vbaol11.chm1000369
f1_keywords:
- vbaol11.chm1000369
ms.prod: outlook
api_name:
- Outlook.OlkDateControl.KeyDown
ms.assetid: 8b24fba9-5af4-9519-8391-1a57fab6e39e
ms.date: 06/08/2017
---


# OlkDateControl.KeyDown Event (Outlook)

Occurs when a user presses a key.


## Syntax

 _expression_ . **KeyDown**( **_KeyCode_** , **_Shift_** )

 _expression_ A variable that represents an **OlkDateControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](olshiftstate-enumeration-outlook.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys ( **SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyDown** event is accessible through the _Shift_ parameter.


## See also


#### Concepts


[OlkDateControl Object](olkdatecontrol-object-outlook.md)

