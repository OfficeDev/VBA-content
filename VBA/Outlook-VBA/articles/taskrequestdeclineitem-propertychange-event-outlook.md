---
title: TaskRequestDeclineItem.PropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.PropertyChange
ms.assetid: 89e39434-0b93-5b40-852a-33d0efdca9e6
ms.date: 06/08/2017
---


# TaskRequestDeclineItem.PropertyChange Event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.


## Syntax

 _expression_ . **PropertyChange**( **_Name_** )

 _expression_ A variable that represents a **TaskRequestDeclineItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


#### Concepts


[TaskRequestDeclineItem Object](taskrequestdeclineitem-object-outlook.md)

