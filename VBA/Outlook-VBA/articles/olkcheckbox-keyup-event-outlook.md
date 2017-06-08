---
title: OlkCheckBox.KeyUp Event (Outlook)
keywords: vbaol11.chm1000158
f1_keywords:
- vbaol11.chm1000158
ms.prod: outlook
api_name:
- Outlook.OlkCheckBox.KeyUp
ms.assetid: 47ec2354-78c7-2929-504a-3e0155806aeb
ms.date: 06/08/2017
---


# OlkCheckBox.KeyUp Event (Outlook)

Occurs when the user releases a key.


## Syntax

 _expression_ . **KeyUp**( **_KeyCode_** , **_Shift_** )

 _expression_ A variable that represents an **OlkCheckBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](olshiftstate-enumeration-outlook.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys ( **SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyUp** event is accessible through the _Shift_ parameter.


## See also


#### Concepts


[OlkCheckBox Object](olkcheckbox-object-outlook.md)

