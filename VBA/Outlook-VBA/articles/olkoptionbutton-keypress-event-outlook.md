---
title: OlkOptionButton.KeyPress Event (Outlook)
keywords: vbaol11.chm1000187
f1_keywords:
- vbaol11.chm1000187
ms.prod: outlook
api_name:
- Outlook.OlkOptionButton.KeyPress
ms.assetid: ce52b457-46c0-99d5-955c-e468abec0211
ms.date: 06/08/2017
---


# OlkOptionButton.KeyPress Event (Outlook)

Occurs when the user presses an ANSI key.


## Syntax

 _expression_ . **KeyPress**( **_KeyAscii_** )

 _expression_ A variable that represents an **OlkOptionButton** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyAscii_|Required| **Long**|The numerical value of the key pressed.|

## Remarks

An ANSI key is one that produces a typeable character when the user presses it. The  **KeyPress** event occurs when the user presses an ANSI key on a running form while the form or a control on it has the focus. The event can occur either before or after the key is released.

A  **KeyPress** event does not occur under the following conditions:


- Pressing  **TAB**
    
- Pressing  **ENTER**
    
- Pressing an arrow key
    
- When a keystroke causes the focus to move from one control to another.
    



## See also


#### Concepts


[OlkOptionButton Object](olkoptionbutton-object-outlook.md)

