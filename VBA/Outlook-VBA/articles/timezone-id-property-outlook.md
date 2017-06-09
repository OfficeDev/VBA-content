---
title: TimeZone.ID Property (Outlook)
keywords: vbaol11.chm3304
f1_keywords:
- vbaol11.chm3304
ms.prod: outlook
api_name:
- Outlook.TimeZone.ID
ms.assetid: 13d4826f-5291-993c-2da1-f1dc65a1e086
ms.date: 06/08/2017
---


# TimeZone.ID Property (Outlook)

Returns a  **String** that uniquely identifies the time zone. Read-only.


## Syntax

 _expression_ . **ID**

 _expression_ A variable that represents a **TimeZone** object.


## Remarks

The  **ID** of a time zone is globally the same for that time zone. It is the name of the Windows registry key that contains the time zone information. Unlike the **[Name](timezone-name-property-outlook.md)** property, the value of **ID** is not localized.


## See also


#### Concepts


[TimeZone Object](timezone-object-outlook.md)

