---
title: MailItem.MarkAsTask Method (Outlook)
keywords: vbaol11.chm3039
f1_keywords:
- vbaol11.chm3039
ms.prod: outlook
api_name:
- Outlook.MailItem.MarkAsTask
ms.assetid: ee38093d-a180-07f7-eae8-c9dbb2e8f413
ms.date: 06/08/2017
---


# MailItem.MarkAsTask Method (Outlook)

Marks a  **[MailItem](mailitem-object-outlook.md)** object as a task and assigns a task interval for the object.


## Syntax

 _expression_ . **MarkAsTask**( **_MarkInterval_** )

 _expression_ An expression that returns a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MarkInterval_|Required| **[OlMarkInterval](olmarkinterval-enumeration-outlook.md)**|The task interval for the  **MailItem** .|

## Remarks

Calling this method sets the value of several other properties, depending on the value provided in  _MarkInterval_. For more information about the properties set by specifying  _MarkInterval_, see [OlMarkInterval Enumeration](olmarkinterval-enumeration-outlook.md).


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

