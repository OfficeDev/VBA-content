---
title: CalendarSharing.IncludeWholeCalendar Property (Outlook)
keywords: vbaol11.chm2420
f1_keywords:
- vbaol11.chm2420
ms.prod: outlook
api_name:
- Outlook.CalendarSharing.IncludeWholeCalendar
ms.assetid: 6cb75f0e-afb9-48fc-5b96-9f64a3b2ed6f
ms.date: 06/08/2017
---


# CalendarSharing.IncludeWholeCalendar Property (Outlook)

Returns or sets a  **Boolean** value that indicates whether all calendar items in the folder should be included in the iCalendar (.ics) file created by the **[ForwardAsICal](calendarsharing-forwardasical-method-outlook.md)** or **[SaveAsICal](calendarsharing-saveasical-method-outlook.md)** methods of the **[CalendarSharing](calendarsharing-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **IncludeWholeCalendar**

 _expression_ An expression that returns a **CalendarSharing** object.


### Return Value

 **True** if all calendar items in the folder should be included; otherwise, **False** .


## Remarks

If this property is set to  **true** , the **[StartDate](calendarsharing-startdate-property-outlook.md)** and **[EndDate](calendarsharing-enddate-property-outlook.md)** properties of the **CalendarSharing** object are ignored and all calendar items in the folder are included.

If this property is set to  **false** , the **StartDate** and **EndDate** properties determine the range of calendar items to be included.


## See also


#### Concepts


[CalendarSharing Object](calendarsharing-object-outlook.md)

