---
title: OlCalendarMailFormat Enumeration (Outlook)
keywords: vbaol11.chm3117
f1_keywords:
- vbaol11.chm3117
ms.prod: outlook
api_name:
- Outlook.OlCalendarMailFormat
ms.assetid: b4b77080-1c8b-cfa4-3b3a-e59fec698bb1
ms.date: 06/08/2017
---


# OlCalendarMailFormat Enumeration (Outlook)

Determines the format of the calendar information in the body of the  **[MailItem](mailitem-object-outlook.md)** created by the **[ForwardAsICal](calendarsharing-forwardasical-method-outlook.md)** method.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olCalendarMailFormatDailySchedule**|0|The calendar information is formatted as a daily schedule of appointments, containing an hour-by-hour breakdown of the calendar, showing both free and busy time blocks along with working-hour information. This layout is intended to help show recipients which times you are available. |
| **olCalendarMailFormatEventList**|1|The calendar information is formatted as a list of events, containing a list of the calendar appointments without showing any time blocks. This layout is intended to help show recipients the events scheduled for a given time period.|

## Remarks

For more information, see [Sharing Calendars](http://msdn.microsoft.com/library/03e0b693-5446-ca62-f868-69a583087966%28Office.15%29.aspx) and[Export a Calendar using Payload Sharing](http://msdn.microsoft.com/library/acd7d29e-12d6-a5ea-c1a6-8b3165b27dc7%28Office.15%29.aspx).


