---
title: MeetingItem.IsLatestVersion Property (Outlook)
keywords: vbaol11.chm3535
f1_keywords:
- vbaol11.chm3535
ms.prod: outlook
api_name:
- Outlook.MeetingItem.IsLatestVersion
ms.assetid: aee3a832-b1b5-538d-dd45-e64769662dfc
ms.date: 06/08/2017
---


# MeetingItem.IsLatestVersion Property (Outlook)

Returns a  **Boolean** value that indicates whether the **[MeetingItem](meetingitem-object-outlook.md)** represents the latest version of the item on the organizer's calendar. Read-only.


## Syntax

 _expression_ . **IsLatestVersion**

 _expression_ A variable that represents a **MeetingItem** object.


## Remarks

This property applies to all  **MeetingItem** objects including meeting requests, meeting updates, and meeting cancellation.

 **True** indicates that the latest version of the meeting item is on the organizer's calendar; **False** indicates that the meeting item on the calendar is not the latest version, or that there is a conflict between the meeting request and another appointment item on the calendar.


## See also


#### Concepts


[MeetingItem Object](meetingitem-object-outlook.md)

