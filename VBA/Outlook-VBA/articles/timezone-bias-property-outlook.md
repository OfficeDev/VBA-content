---
title: TimeZone.Bias Property (Outlook)
keywords: vbaol11.chm3285
f1_keywords:
- vbaol11.chm3285
ms.prod: outlook
api_name:
- Outlook.TimeZone.Bias
ms.assetid: 18f55011-5d71-2e3b-4049-a37323f09478
ms.date: 06/08/2017
---


# TimeZone.Bias Property (Outlook)

Returns a  **Long** value that represents the difference in minutes of between the local time in this time zone and the Coordinated Universal Time (UTC). Read-only.


## Syntax

 _expression_ . **Bias**

 _expression_ A variable that represents a **TimeZone** object.


## Remarks

This value is stored as part of the value for  **TZI** for that time zone in the Windows registry. The **TZI** value is mapped to the Windows **[TIME_ZONE_INFORMATION](http://msdn.microsoft.com/library/base.time_zone_information_str%28Office.15%29.aspx)** structure.

 **Bias** does not take into account any time offset for daylight time or standard time in the time zone. To account for any daylight time offset, use **[DaylightBias](timezone-daylightbias-property-outlook.md)** . In general, when the local time zone is adopting daylight time, UTC time is the result of adding the **Bias** and **DaylightBias** to the local time. To account for any standard time offset, use **[StandardBias](timezone-standardbias-property-outlook.md)** . In general, when the local time zone is adopting standard time, UTC time is the result of adding the **Bias** and **StandardBias** to the local time.

For example, in a state adopting daylight time in the Pacific time zone, the  **Bias** is 480 minutes and **DaylightBias** is -60 minutes. To determine the time in UTC for June 11, 2 A.M. PST, add a **Bias** of (480/60) hours and a **DaylightBias** of -(60/60) hours to the local time June 11, 2 A.M. The time in UTC is June 11, 9 A.M.


## See also


#### Concepts


[TimeZone Object](timezone-object-outlook.md)

