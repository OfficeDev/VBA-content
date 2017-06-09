---
title: Calendar Object (Project)
ms.prod: project-server
api_name:
- Project.Calendar
ms.assetid: 2d3b0f05-4762-0058-15d4-47e1d2b9d9a9
ms.date: 06/08/2017
---


# Calendar Object (Project)



Represents the calendar for a resource or project. The  **Calendar** object is a member of the **[Calendars](calendars-object-project.md)** collection.
 **Using the Calendar Object**
Use  **BaseCalendars(** _Index_ **)**, where _Index_ is the calendar index number or calendar name, to return a single **Calendar** object.
 **Using the Calendars Collection**
Use the  **[BaseCalendars](http://msdn.microsoft.com/library/fb7f55f6-6618-fb82-dae1-320953bcf79d%28Office.15%29.aspx)** property to return a **Calendars** collection. The following example resets the properties of each base calendar in the active project to their default values.
Use the  **[BaseCalendarCreate](http://msdn.microsoft.com/library/c9c92dff-255a-041b-c18d-49d6d75884e3%28Office.15%29.aspx)** method to add a **Calendar** object to the **Calendars** collection. The following example creates a new base calendar.

## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/8bc3e8cc-34f4-17be-d142-51290ee4bea3%28Office.15%29.aspx)|
|[Period](http://msdn.microsoft.com/library/b717bcbe-654b-5791-2002-d65e2a96617f%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/fc638f47-36b5-aa36-55c2-882bd570b9cb%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/f3963ec1-923b-ea62-855b-107519dd7e13%28Office.15%29.aspx)|
|[BaseCalendar](http://msdn.microsoft.com/library/3ea2b0e2-8d73-b564-fdd1-a098a8428562%28Office.15%29.aspx)|
|[Enterprise](http://msdn.microsoft.com/library/1e160265-1c49-e95d-f04e-e87ce0222f85%28Office.15%29.aspx)|
|[Exceptions](http://msdn.microsoft.com/library/2631d4c8-1e71-ca75-8291-8e2544e53c00%28Office.15%29.aspx)|
|[Guid](http://msdn.microsoft.com/library/08230f82-fd1b-ef99-18e3-f6be75c3d2a8%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/ad177421-1e7b-5c85-e437-f3d2b83a66c5%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/e437e29c-ed61-c83a-53b7-8a0d1cb7cb4e%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/b86fa5e8-f468-862e-f8a9-7ab2cb6b43b3%28Office.15%29.aspx)|
|[ResourceGuid](http://msdn.microsoft.com/library/c66c3e90-06e0-5b48-3e44-48e366377258%28Office.15%29.aspx)|
|[WeekDays](http://msdn.microsoft.com/library/4495a739-156b-8cda-d3d0-acbc56b767ff%28Office.15%29.aspx)|
|[WorkWeeks](http://msdn.microsoft.com/library/c4a3887b-0518-2b22-0288-500ad567a301%28Office.15%29.aspx)|
|[Years](http://msdn.microsoft.com/library/63f17754-d258-3fd2-5f20-33b8998e7e4d%28Office.15%29.aspx)|

