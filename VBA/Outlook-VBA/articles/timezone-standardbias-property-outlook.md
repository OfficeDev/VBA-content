---
title: TimeZone.StandardBias Property (Outlook)
keywords: vbaol11.chm3286
f1_keywords:
- vbaol11.chm3286
ms.prod: outlook
api_name:
- Outlook.TimeZone.StandardBias
ms.assetid: 0400a70c-4a53-417d-8d6e-c0271b4c1dcb
ms.date: 06/08/2017
---


# TimeZone.StandardBias Property (Outlook)

Returns a  **Long** value that represents the time offset in minutes from the **[Bias](timezone-bias-property-outlook.md)** to account for standard time in this time zone. Read-only.


## Syntax

 _expression_ . **StandardBias**

 _expression_ A variable that represents a **TimeZone** object.


## Remarks

This value is stored as part of the value for  **TZI** for the time zone in the Windows registry. The **TZI** value is mapped to the Windows **[TIME_ZONE_INFORMATION](http://msdn.microsoft.com/library/base.time_zone_information_str%28Office.15%29.aspx)** structure.

In relation to the UTC time and the local time of the time zone, UTC time is the result of adding the  **Bias** and **StandardBias** to the local time. For example, in a state adopting standard time in the Pacific time zone, the **Bias** is 480 minutes and **StandardBias** is 0 minutes. To determine the time in UTC for June 11, 2 A.M. PST, add a **Bias** of (480/60) hours and a **StandardBias** of 0 hours to the local time June 11, 2 A.M. The time in UTC is June 11, 10 A.M.


## See also


#### Concepts


[TimeZone Object](timezone-object-outlook.md)

