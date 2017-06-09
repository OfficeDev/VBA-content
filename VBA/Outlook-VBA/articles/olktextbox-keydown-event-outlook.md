---
title: OlkTextBox.KeyDown Event (Outlook)
keywords: vbaol11.chm1000078
f1_keywords:
- vbaol11.chm1000078
ms.prod: outlook
api_name:
- Outlook.OlkTextBox.KeyDown
ms.assetid: a6e5a293-41a4-9237-851b-1352eeee0f41
ms.date: 06/08/2017
---


# OlkTextBox.KeyDown Event (Outlook)

Occurs when a user presses a key.


## Syntax

 _expression_ . **KeyDown**( **_KeyCode_** , **_Shift_** )

 _expression_ A variable that represents an **OlkTextBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|The numerical value of the key pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](olshiftstate-enumeration-outlook.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|

## Remarks

The state of the modifier keys ( **SHIFT**,  **CTRL**, or  **ALT**) that are pressed during the  **KeyDown** event is accessible through the _Shift_ parameter.


## See also


#### Concepts


[OlkTextBox Object](olktextbox-object-outlook.md)

