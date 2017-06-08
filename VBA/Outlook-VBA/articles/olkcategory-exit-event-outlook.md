---
title: OlkCategory.Exit Event (Outlook)
keywords: vbaol11.chm1000455
f1_keywords:
- vbaol11.chm1000455
ms.prod: outlook
api_name:
- Outlook.OlkCategory.Exit
ms.assetid: bc1dac11-00f0-7fcb-9a8f-c8fde0d61e51
ms.date: 06/08/2017
---


# OlkCategory.Exit Event (Outlook)

Occurs just after the focus passes from this control to another control on the same form.


## Syntax

 _expression_ . **Exit**( **_Cancel_** )

 _expression_ A variable that represents an **OlkCategory** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the exit operation is not completed and the focus remains in this control.|

## See also


#### Concepts


[OlkCategory Object](olkcategory-object-outlook.md)

