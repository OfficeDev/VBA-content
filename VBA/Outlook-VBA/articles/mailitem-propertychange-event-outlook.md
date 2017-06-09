---
title: MailItem.PropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.PropertyChange
ms.assetid: 768de21f-a474-4574-74f4-6d99e3ab542e
ms.date: 06/08/2017
---


# MailItem.PropertyChange Event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.


## Syntax

 _expression_ . **PropertyChange**( **_Name_** )

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

