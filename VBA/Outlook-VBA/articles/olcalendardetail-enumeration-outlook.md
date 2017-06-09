---
title: OlCalendarDetail Enumeration (Outlook)
keywords: vbaol11.chm3118
f1_keywords:
- vbaol11.chm3118
ms.prod: outlook
api_name:
- Outlook.OlCalendarDetail
ms.assetid: 7ad41002-490e-824c-ff63-83a164218839
ms.date: 06/08/2017
---


# OlCalendarDetail Enumeration (Outlook)

Indicates the level of detail for calendar items that will be exported to an iCalendar (.ics) file. 



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olFreeBusyAndSubject**|1|Free/busy information and the appointment subjects are exported to the iCalendar file. |
| **olFreeBusyOnly**|0|Only free/busy information is exported to the iCalendar file.|
| **olFullDetails**|2|Full details of each appointment item are exported to the iCalendar file. |

## Remarks

This enumeration is used by the [CalendarDetail ](calendarsharing-calendardetail-property-outlook.md) property of the[CalendarSharing Object (Outlook)](calendarsharing-object-outlook.md) to determine the level of detail for calendar items stored in the iCalendar file created by the[ForwardAsICal](calendarsharing-forwardasical-method-outlook.md) and[SaveAsICal](calendarsharing-saveasical-method-outlook.md) methods.


