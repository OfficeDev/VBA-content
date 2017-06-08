---
title: OlkOptionButton.KeyUp Event (Outlook)
keywords: vbaol11.chm1000188
f1_keywords:
- vbaol11.chm1000188
ms.prod: outlook
api_name:
- Outlook.OlkOptionButton.KeyUp
ms.assetid: abca8eca-b1a2-a050-0418-daa10cc4cabc
ms.date: 06/08/2017
---


# OlkOptionButton.KeyUp Event (Outlook)

Occurs when the user releases a key.


## Syntax

 _expression_ . **KeyUp**( **_KeyCode_** , **_Shift_** )

 _expression_ A variable that represents an **OlkOptionButton** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](olshiftstate-enumeration-outlook.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys ( **SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyUp** event is accessible through the _Shift_ parameter.


## See also


#### Concepts


[OlkOptionButton Object](olkoptionbutton-object-outlook.md)

