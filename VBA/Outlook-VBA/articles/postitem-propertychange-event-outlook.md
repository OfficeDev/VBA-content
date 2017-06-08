---
title: PostItem.PropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.PostItem.PropertyChange
ms.assetid: 71ca9e98-9ea5-e8da-a1af-0fd6c153df83
ms.date: 06/08/2017
---


# PostItem.PropertyChange Event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.


## Syntax

 _expression_ . **PropertyChange**( **_Name_** )

 _expression_ A variable that represents a **PostItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


#### Concepts


[PostItem Object](postitem-object-outlook.md)

