---
title: OlkDateControl.Exit Event (Outlook)
keywords: vbaol11.chm1000368
f1_keywords:
- vbaol11.chm1000368
ms.prod: outlook
api_name:
- Outlook.OlkDateControl.Exit
ms.assetid: 6a8ec569-4e08-0400-95ad-934cbe2c20e4
ms.date: 06/08/2017
---


# OlkDateControl.Exit Event (Outlook)

Occurs just after the focus passes from this control to another control on the same form.


## Syntax

_expression_. **Exit** (**_Cancel_**)

_expression_ A variable that represents an **OlkDateControl** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|_Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the exit operation is not completed and the focus remains in this control.|

<br/>

## See also

#### Concepts

- [OlkDateControl Object](olkdatecontrol-object-outlook.md)

