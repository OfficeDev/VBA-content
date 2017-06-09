---
title: OlkCommandButton.Exit Event (Outlook)
keywords: vbaol11.chm1000126
f1_keywords:
- vbaol11.chm1000126
ms.prod: outlook
api_name:
- Outlook.OlkCommandButton.Exit
ms.assetid: be3f7740-8682-ecc5-3927-dd700f26b49c
ms.date: 06/08/2017
---


# OlkCommandButton.Exit Event (Outlook)

Occurs just after the focus passes from this control to another control on the same form.


## Syntax

 _expression_ . **Exit**( **_Cancel_** )

 _expression_ A variable that represents an **OlkCommandButton** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the exit operation is not completed and the focus remains in this control.|

## See also


#### Concepts


[OlkCommandButton Object](olkcommandbutton-object-outlook.md)

