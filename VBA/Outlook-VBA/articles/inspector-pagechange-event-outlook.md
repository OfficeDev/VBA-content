---
title: Inspector.PageChange Event (Outlook)
keywords: vbaol11.chm472
f1_keywords:
- vbaol11.chm472
ms.prod: outlook
api_name:
- Outlook.Inspector.PageChange
ms.assetid: f0ba9820-84bf-2367-364a-519e6ed88289
ms.date: 06/08/2017
---


# Inspector.PageChange Event (Outlook)

Occurs when the active form page changes, either programmatically or by user action, on an [Inspector](inspector-object-outlook.md) object.


## Syntax

 _expression_ . **PageChange**( **_ActivePageName_** )

 _expression_ A variable that represents an **Inspector** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ActivePageName_|Required| **String**|The name of the active page.|

## Remarks

An error occurs if the event handler for this event calls either the  **[Close](inspector-close-method-outlook.md)** or **[SetCurrentFormPage](inspector-setcurrentformpage-method-outlook.md)** methods.


## See also


#### Concepts


[Inspector Object](inspector-object-outlook.md)

