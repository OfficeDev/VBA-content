---
title: ReportItem.PropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ReportItem.PropertyChange
ms.assetid: 5fd89535-8fa4-202e-bb0a-1dc4d608acec
ms.date: 06/08/2017
---


# ReportItem.PropertyChange Event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.


## Syntax

 _expression_ . **PropertyChange**( **_Name_** )

 _expression_ A variable that represents a **ReportItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


#### Concepts


[ReportItem Object](reportitem-object-outlook.md)

