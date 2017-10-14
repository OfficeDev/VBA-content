---
title: OlkComboBox.Exit Event (Outlook)
keywords: vbaol11.chm1000241
f1_keywords:
- vbaol11.chm1000241
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.Exit
ms.assetid: ce386495-2c81-b256-c1dd-ede086f7a0f3
ms.date: 06/08/2017
---


# OlkComboBox.Exit Event (Outlook)

Occurs just after the focus passes from this control to another control on the same form.


## Syntax

 _expression_ . **Exit**( **_Cancel_** )

 _expression_ A variable that represents an **OlkComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the exit operation is not completed and the focus remains in this control.|

## See also


#### Concepts


[OlkComboBox Object](olkcombobox-object-outlook.md)

