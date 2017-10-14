---
title: TimeZones.ConvertTime Method (Outlook)
keywords: vbaol11.chm3297
f1_keywords:
- vbaol11.chm3297
ms.prod: outlook
api_name:
- Outlook.TimeZones.ConvertTime
ms.assetid: 6a935961-2030-ed9c-5c1b-4e6641ee3913
ms.date: 06/08/2017
---


# TimeZones.ConvertTime Method (Outlook)

Converts a date/time value from one time zone to another time zone.


## Syntax

 _expression_ . **ConvertTime**( **_SourceDateTime_** , **_SourceTimeZone_** , **_DestinationTimeZone_** )

 _expression_ A variable that represents a **TimeZones** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SourceDateTime_|Required| **Date**|A date/time value expressed in the original time zone.|
| _SourceTimeZone_|Required| **[TimeZone](timezone-object-outlook.md)**|The original time zone of the date/time value that is to be converted.|
| _DestinationTimeZone_|Required| **TimeZone**|The target time zone to which the date/time value is to be converted.|

### Return Value

A  **Date** value that represents the date and time expressed in the _DestinationTimeZone_ .


## See also


#### Concepts


[TimeZones Object](timezones-object-outlook.md)

