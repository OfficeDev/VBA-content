---
title: JournalItem.PropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.JournalItem.PropertyChange
ms.assetid: a04a13ea-85ce-f93e-37af-fa7b77757204
ms.date: 06/08/2017
---


# JournalItem.PropertyChange Event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.


## Syntax

 _expression_ . **PropertyChange**( **_Name_** )

 _expression_ A variable that represents a **JournalItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


#### Concepts


[JournalItem Object](journalitem-object-outlook.md)

