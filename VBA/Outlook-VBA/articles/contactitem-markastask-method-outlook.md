---
title: ContactItem.MarkAsTask Method (Outlook)
keywords: vbaol11.chm3031
f1_keywords:
- vbaol11.chm3031
ms.prod: outlook
api_name:
- Outlook.ContactItem.MarkAsTask
ms.assetid: def25d8d-6074-5e4d-18d9-82381b0b7876
ms.date: 06/08/2017
---


# ContactItem.MarkAsTask Method (Outlook)

Marks a  **[ContactItem](contactitem-object-outlook.md)** object as a task and assigns a task interval for the object.


## Syntax

 _expression_ . **MarkAsTask**( **_MarkInterval_** )

 _expression_ An expression that returns a **ContactItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MarkInterval_|Required| **[OlMarkInterval](olmarkinterval-enumeration-outlook.md)**|The task interval for the  **ContactItem** .|

## Remarks

Calling this method sets the value of several other properties, depending on the value provided in  _MarkInterval_. For more information about the properties set by specifying  _MarkInterval_, see [OlMarkInterval Enumeration](olmarkinterval-enumeration-outlook.md).


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

