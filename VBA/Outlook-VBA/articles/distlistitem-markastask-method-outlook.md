---
title: DistListItem.MarkAsTask Method (Outlook)
keywords: vbaol11.chm3035
f1_keywords:
- vbaol11.chm3035
ms.prod: outlook
api_name:
- Outlook.DistListItem.MarkAsTask
ms.assetid: a8f5a666-95d6-9a97-14bb-ca0b6481e7a8
ms.date: 06/08/2017
---


# DistListItem.MarkAsTask Method (Outlook)

Marks a  **[DistListItem](distlistitem-object-outlook.md)** object as a task and assigns a task interval for the object.


## Syntax

 _expression_ . **MarkAsTask**( **_MarkInterval_** )

 _expression_ An expression that returns a **DistListItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MarkInterval_|Required| **[OlMarkInterval](olmarkinterval-enumeration-outlook.md)**|The task interval for the  **DistListItem** .|

## Remarks

Calling this method sets the value of several other properties, depending on the value provided in  _MarkInterval_. For more information about the properties set by specifying  _MarkInterval_, see [OlMarkInterval Enumeration](olmarkinterval-enumeration-outlook.md).


## See also


#### Concepts


[DistListItem Object](distlistitem-object-outlook.md)

