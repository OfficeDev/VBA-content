---
title: TimeZone.DaylightDate Property (Outlook)
keywords: vbaol11.chm3289
f1_keywords:
- vbaol11.chm3289
ms.prod: outlook
api_name:
- Outlook.TimeZone.DaylightDate
ms.assetid: a653b0ec-1462-165f-36e3-1be57513a2c7
ms.date: 06/08/2017
---


# TimeZone.DaylightDate Property (Outlook)

Returns a  **Date** value that represents the date and time in this time zone when time changes over to daylight time in the current year. Read-only.


## Syntax

 _expression_ . **DaylightDate**

 _expression_ A variable that represents a **TimeZone** object.


## Remarks

This value is stored as part of the  **TZI** value for the time zone in the Windows registry. The **TZI** value is mapped to the Windows **[TIME_ZONE_INFORMATION](http://msdn.microsoft.com/library/base.time_zone_information_str%28Office.15%29.aspx)** structure.


## See also


#### Concepts


[TimeZone Object](timezone-object-outlook.md)

