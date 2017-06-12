---
title: OlkCommandButton.KeyUp Event (Outlook)
keywords: vbaol11.chm1000129
f1_keywords:
- vbaol11.chm1000129
ms.prod: outlook
api_name:
- Outlook.OlkCommandButton.KeyUp
ms.assetid: 63d8067e-1ec1-324e-e671-aa027bc22ace
ms.date: 06/08/2017
---


# OlkCommandButton.KeyUp Event (Outlook)

Occurs when the user releases a key.


## Syntax

 _expression_ . **KeyUp**( **_KeyCode_** , **_Shift_** )

 _expression_ A variable that represents an **OlkCommandButton** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](olshiftstate-enumeration-outlook.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys ( **SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyUp** event is accessible through the _Shift_ parameter.


## See also


#### Concepts


[OlkCommandButton Object](olkcommandbutton-object-outlook.md)

