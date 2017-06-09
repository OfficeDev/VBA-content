---
title: PostItem.MarkAsTask Method (Outlook)
keywords: vbaol11.chm3043
f1_keywords:
- vbaol11.chm3043
ms.prod: outlook
api_name:
- Outlook.PostItem.MarkAsTask
ms.assetid: 78ead34b-3861-0204-1bc3-687a2c25ab73
ms.date: 06/08/2017
---


# PostItem.MarkAsTask Method (Outlook)

Marks a  **[PostItem](postitem-object-outlook.md)** object as a task and assigns a task interval for the object.


## Syntax

 _expression_ . **MarkAsTask**( **_MarkInterval_** )

 _expression_ An expression that returns a **PostItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MarkInterval_|Required| **[OlMarkInterval](olmarkinterval-enumeration-outlook.md)**|The task interval for the  **PostItem** .|

## Remarks

Calling this method sets the value of several other properties, depending on the value provided in  _MarkInterval_. For more information about the properties set by specifying  _MarkInterval_, see [OlMarkInterval Enumeration](olmarkinterval-enumeration-outlook.md).


## See also


#### Concepts


[PostItem Object](postitem-object-outlook.md)

