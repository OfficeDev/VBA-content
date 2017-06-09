---
title: TimeZone.DaylightBias Property (Outlook)
keywords: vbaol11.chm3287
f1_keywords:
- vbaol11.chm3287
ms.prod: outlook
api_name:
- Outlook.TimeZone.DaylightBias
ms.assetid: 59c83104-7ce5-95a9-71fa-df3b0a96e173
ms.date: 06/08/2017
---


# TimeZone.DaylightBias Property (Outlook)

Returns a  **Long** value that represents the time offset in minutes from the **[Bias](timezone-bias-property-outlook.md)** to account for daylight time in this time zone. Read-only.


## Syntax

 _expression_ . **DaylightBias**

 _expression_ A variable that represents a **TimeZone** object.


## Remarks

This value is stored as part of the value for  **TZI** for that time zone in the Windows registry. The **TZI** value is mapped to the Windows **[TIME_ZONE_INFORMATION](http://msdn.microsoft.com/library/base.time_zone_information_str%28Office.15%29.aspx)** structure.

In relation to the UTC time and the local time of the time zone, UTC time is the result of adding the  **Bias** and **DaylightBias** to the local time. For example, in a state adopting daylight time in the Pacific time zone, the **Bias** is 480 minutes and **DaylightBias** is -60 minutes. To determine the time in UTC for June 11, 2 A.M. PST, add a **Bias** of (480/60) hours and a **DaylightBias** of -(60/60) hours to the local time June 11, 2 A.M. The time in UTC is June 11, 9 A.M.


## See also


#### Concepts


[TimeZone Object](timezone-object-outlook.md)

