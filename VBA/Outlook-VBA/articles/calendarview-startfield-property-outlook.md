---
title: CalendarView.StartField Property (Outlook)
keywords: vbaol11.chm2625
f1_keywords:
- vbaol11.chm2625
ms.prod: outlook
api_name:
- Outlook.CalendarView.StartField
ms.assetid: 085c6605-0bff-98a5-fb48-ce32b76037db
ms.date: 06/08/2017
---


# CalendarView.StartField Property (Outlook)

Returns or sets a  **String** value that represents the name of the property that starts the time duration for Outlook items displayed in the **[CalendarView](calendarview-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **StartField**

 _expression_ A variable that represents a **CalendarView** object.


## Remarks

The values of the  **StartField** and **[EndField](calendarview-endfield-property-outlook.md)** properties indicate which Outlook item properties are used by the **CalendarView** object to represent the duration of an Outlook item. Both custom and built-in properties can be specified, but only date/time properties are allowed.


## See also


#### Concepts


[CalendarView Object](calendarview-object-outlook.md)

