---
title: TaskRequestUpdateItem.PropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.PropertyChange
ms.assetid: 47121ba2-cd73-405a-9bd0-d8fc4a77a535
ms.date: 06/08/2017
---


# TaskRequestUpdateItem.PropertyChange Event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.


## Syntax

 _expression_ . **PropertyChange**( **_Name_** )

 _expression_ A variable that represents a **TaskRequestUpdateItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


#### Concepts


[TaskRequestUpdateItem Object](taskrequestupdateitem-object-outlook.md)

