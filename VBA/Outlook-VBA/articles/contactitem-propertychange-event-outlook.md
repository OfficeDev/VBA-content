---
title: ContactItem.PropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ContactItem.PropertyChange
ms.assetid: 4138deee-2915-f581-b003-16007e37f128
ms.date: 06/08/2017
---


# ContactItem.PropertyChange Event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.


## Syntax

 _expression_ . **PropertyChange**( **_Name_** )

 _expression_ A variable that represents a **ContactItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

