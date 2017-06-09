---
title: TimeZone.StandardDate Property (Outlook)
keywords: vbaol11.chm3288
f1_keywords:
- vbaol11.chm3288
ms.prod: outlook
api_name:
- Outlook.TimeZone.StandardDate
ms.assetid: 61114f2b-e0cf-80e9-ef4c-2553fba68fe1
ms.date: 06/08/2017
---


# TimeZone.StandardDate Property (Outlook)

Returns a  **Date** value that represents the date and time in this time zone when time changes over to standard time. Read-only.


## Syntax

 _expression_ . **StandardDate**

 _expression_ A variable that represents a **TimeZone** object.


## Remarks

This value is stored as part of the  **TZI** value for the time zone in the Windows registry. The **TZI** value is mapped to the Windows **[TIME_ZONE_INFORMATION](http://msdn.microsoft.com/library/base.time_zone_information_str%28Office.15%29.aspx)** structure.


## See also


#### Concepts


[TimeZone Object](timezone-object-outlook.md)

