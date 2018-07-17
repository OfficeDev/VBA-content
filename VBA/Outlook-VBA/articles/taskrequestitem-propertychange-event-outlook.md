---
title: TaskRequestItem.PropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.PropertyChange
ms.assetid: 96e99389-0393-1350-bdfd-45e097d5e185
ms.date: 06/08/2017
---


# TaskRequestItem.PropertyChange Event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.


## Syntax

 _expression_ . **PropertyChange**( **_Name_** )

 _expression_ A variable that represents a **TaskRequestItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


#### Concepts


[TaskRequestItem Object](taskrequestitem-object-outlook.md)

