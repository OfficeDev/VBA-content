---
title: TaskItem.PropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskItem.PropertyChange
ms.assetid: adc96ece-cea5-c939-7f9a-aa7d0f16960b
ms.date: 06/08/2017
---


# TaskItem.PropertyChange Event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](taskitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.


## Syntax

 _expression_ . **PropertyChange**( **_Name_** )

 _expression_ A variable that represents a **TaskItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

