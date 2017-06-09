---
title: Application.TimeZones Property (Outlook)
keywords: vbaol11.chm3270
f1_keywords:
- vbaol11.chm3270
ms.prod: outlook
api_name:
- Outlook.Application.TimeZones
ms.assetid: 920e55d1-9914-fa74-101a-921083328d23
ms.date: 06/08/2017
---


# Application.TimeZones Property (Outlook)

Returns a  **[TimeZones](timezones-object-outlook.md)** collection that represents the set of time zones supported by Outlook. Read-only.


## Syntax

 _expression_ . **TimeZones**

 _expression_ A variable that represents an **Application** object.


## Remarks

The set of time zones supported by Outlook is based on the data stored in the Windows registry key HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones.


## See also


#### Concepts


[Application Object](application-object-outlook.md)

