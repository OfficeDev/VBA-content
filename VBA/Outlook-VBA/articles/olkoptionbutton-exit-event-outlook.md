---
title: OlkOptionButton.Exit Event (Outlook)
keywords: vbaol11.chm1000185
f1_keywords:
- vbaol11.chm1000185
ms.prod: outlook
api_name:
- Outlook.OlkOptionButton.Exit
ms.assetid: 25967971-8d98-579e-a4f7-e6bfc3a16834
ms.date: 06/08/2017
---


# OlkOptionButton.Exit Event (Outlook)

Occurs just after the focus passes from this control to another control on the same form.


## Syntax

 _expression_ . **Exit**( **_Cancel_** )

 _expression_ A variable that represents an **OlkOptionButton** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the exit operation is not completed and the focus remains in this control.|

## See also


#### Concepts


[OlkOptionButton Object](olkoptionbutton-object-outlook.md)

