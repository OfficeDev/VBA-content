---
title: RemoteItem.PropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.RemoteItem.PropertyChange
ms.assetid: 630d4423-cb56-eef0-e1b1-1afe227c140d
ms.date: 06/08/2017
---


# RemoteItem.PropertyChange Event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.


## Syntax

 _expression_ . **PropertyChange**( **_Name_** )

 _expression_ A variable that represents a **RemoteItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


#### Concepts


[RemoteItem Object](remoteitem-object-outlook.md)

