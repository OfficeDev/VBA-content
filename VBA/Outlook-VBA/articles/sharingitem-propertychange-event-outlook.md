---
title: SharingItem.PropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.PropertyChange
ms.assetid: 7c3cf73a-4b2c-3f74-4d3e-5a0e04870f07
ms.date: 06/08/2017
---


# SharingItem.PropertyChange Event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](sharingitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.


## Syntax

 _expression_ . **PropertyChange**( **_Name_** )

 _expression_ An expression that returns a **SharingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

