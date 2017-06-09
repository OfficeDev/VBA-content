---
title: Folder.GetCalendarExporter Method (Outlook)
keywords: vbaol11.chm2020
f1_keywords:
- vbaol11.chm2020
ms.prod: outlook
api_name:
- Outlook.Folder.GetCalendarExporter
ms.assetid: 7c67e208-65dd-8904-4b6f-8ec2df4e530d
ms.date: 06/08/2017
---


# Folder.GetCalendarExporter Method (Outlook)

Creates a  **[CalendarSharing](calendarsharing-object-outlook.md)** object for the specified **[Folder](folder-object-outlook.md)** .


## Syntax

 _expression_ . **GetCalendarExporter**

 _expression_ An expression that returns a **Folder** object.


### Return Value

A  **CalendarSharing** object for the specified folder.


## Remarks

The  **GetCalendarExporter** method automatically sets the defaults for the **CalendarSharing** class to the standard default options used by the **Folder** object. The **GetCalendarExporter** method can only be used on calendar folders. An error occurs if you use the method on **Folder** objects that represent other folder types.


 **Note**  The  **CalendarSharing** object only supports exporting the iCalendar (.ics) file format.


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

