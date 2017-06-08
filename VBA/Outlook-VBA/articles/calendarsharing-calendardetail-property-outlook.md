---
title: CalendarSharing.CalendarDetail Property (Outlook)
keywords: vbaol11.chm2413
f1_keywords:
- vbaol11.chm2413
ms.prod: outlook
api_name:
- Outlook.CalendarSharing.CalendarDetail
ms.assetid: f3f0ba8d-23db-505f-58c4-6e3a33a468e7
ms.date: 06/08/2017
---


# CalendarSharing.CalendarDetail Property (Outlook)

Returns or sets an  **[OlCalendarDetail](olcalendardetail-enumeration-outlook.md)** value indicating the level of detail for calendar items included in the iCalendar (.ics) file created by the **[ForwardAsICal](calendarsharing-forwardasical-method-outlook.md)** or **[SaveAsICal](calendarsharing-saveasical-method-outlook.md)** methods of the **[CalendarSharing](calendarsharing-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **CalendarDetail**

 _expression_ An expression that returns a **CalendarSharing** object.


### Return Value

A  **OlCalendarDetail** value that indicates the level of detail for calendar items.


## Remarks

The value of this property determines the allowable values for the following properties of the  **CalendarSharing** object:


-  **[IncludeAttachments](calendarsharing-includeattachments-property-outlook.md)** must be set to **False** if **CalendarDetail** is set to **olFreeBusyOnly** or **olFreeBusyAndSubject** .
    
-  **[IncludePrivateDetails](calendarsharing-includeprivatedetails-property-outlook.md)** must be set to **False** if **CalendarDetail** is set to **olFreeBusyOnly** .
    
-  **[RestrictToWorkingHours](calendarsharing-restricttoworkinghours-property-outlook.md)** must be set to **False** if **CalendarDetail** is set to **olFreeBusyAndSubject** or **olFullDetails** .
    

## See also


#### Concepts


[CalendarSharing Object](calendarsharing-object-outlook.md)

