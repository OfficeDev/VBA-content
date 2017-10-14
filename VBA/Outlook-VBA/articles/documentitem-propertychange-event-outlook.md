---
title: DocumentItem.PropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DocumentItem.PropertyChange
ms.assetid: ec757f98-db44-585e-1a4a-5b3044428dec
ms.date: 06/08/2017
---


# DocumentItem.PropertyChange Event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.


## Syntax

 _expression_ . **PropertyChange**( **_Name_** )

 _expression_ A variable that represents a **DocumentItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


#### Concepts


[DocumentItem Object](documentitem-object-outlook.md)

