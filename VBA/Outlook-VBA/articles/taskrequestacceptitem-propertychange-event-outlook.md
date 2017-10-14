---
title: TaskRequestAcceptItem.PropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.PropertyChange
ms.assetid: 4b26e4b6-607c-c9e6-088f-2e7605b0681f
ms.date: 06/08/2017
---


# TaskRequestAcceptItem.PropertyChange Event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.


## Syntax

 _expression_ . **PropertyChange**( **_Name_** )

 _expression_ A variable that represents a **TaskRequestAcceptItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


#### Concepts


[TaskRequestAcceptItem Object](taskrequestacceptitem-object-outlook.md)

