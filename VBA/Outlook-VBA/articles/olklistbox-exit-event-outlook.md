---
title: OlkListBox.Exit Event (Outlook)
keywords: vbaol11.chm1000286
f1_keywords:
- vbaol11.chm1000286
ms.prod: outlook
api_name:
- Outlook.OlkListBox.Exit
ms.assetid: 729d454a-4f52-c0c2-4125-7cbf8ea2d660
ms.date: 06/08/2017
---


# OlkListBox.Exit Event (Outlook)

Occurs just after the focus passes from this control to another control on the same form.


## Syntax

 _expression_ . **Exit**( **_Cancel_** )

 _expression_ A variable that represents an **OlkListBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the exit operation is not completed and the focus remains in this control.|

## See also


#### Concepts


[OlkListBox Object](olklistbox-object-outlook.md)

