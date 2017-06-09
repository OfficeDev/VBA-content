---
title: SharingItem.CustomPropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.CustomPropertyChange
ms.assetid: faf015c1-aa18-67f4-e1af-b456b7c89523
ms.date: 06/08/2017
---


# SharingItem.CustomPropertyChange Event (Outlook)

Occurs when a custom property of an item (which is an instance of the parent object) is changed. 


## Syntax

 _expression_ . **CustomPropertyChange**( **_Name_** )

 _expression_ An expression that returns a **SharingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the custom property that was changed.|

## Remarks

The property name is passed to the procedure so that you can determine which custom property changed.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

