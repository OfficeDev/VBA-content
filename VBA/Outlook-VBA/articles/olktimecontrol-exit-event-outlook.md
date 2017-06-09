---
title: OlkTimeControl.Exit Event (Outlook)
keywords: vbaol11.chm1000407
f1_keywords:
- vbaol11.chm1000407
ms.prod: outlook
api_name:
- Outlook.OlkTimeControl.Exit
ms.assetid: 037013a6-170c-9859-1f0c-705064727c49
ms.date: 06/08/2017
---


# OlkTimeControl.Exit Event (Outlook)

Occurs just after the focus passes from this control to another control on the same form.


## Syntax

 _expression_ . **Exit**( **_Cancel_** )

 _expression_ A variable that represents an **OlkTimeControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the exit operation is not completed and the focus remains in this control.|

## See also


#### Concepts


[OlkTimeControl Object](olktimecontrol-object-outlook.md)

