---
title: WeekDay Object (Project)
ms.prod: project-server
api_name:
- Project.WeekDay
ms.assetid: fc460e89-784b-6764-c22d-e1dcd8a9f297
ms.date: 06/08/2017
---


# WeekDay Object (Project)


 

Represents a weekday in a calendar. The  **Weekday** object is a member of the **[Weekdays](weekdays-object-project.md)** collection.
 
 **Using the Weekday Object**
 
Use  **Weekdays** (*Index* ), where*Index* is the weekday index number, three-letter abbreviation of the day name, or **PjWeekday** constant, to return a single **Weekday** object. The following example sets Friday (the sixth day of a week starting on Sunday) as a half-day by setting the start and finish times for the first shift and clearing the values of the second and third shifts.
 
A much better way to return the same object is to use the predefined constant for Friday instead of the nonintuitive number 6. Thus, the first line of the preceding example would be as follows:
 
 **Using the Weekdays Collection**
 
Use the  **[Weekdays](calendar-weekdays-property-project.md)** property to return a **Weekdays** collection.
 

## Methods



|**Name**|
|:-----|
|[Default](weekday-default-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](weekday-application-property-project.md)|
|[Calendar](weekday-calendar-property-project.md)|
|[Count](weekday-count-property-project.md)|
|[Index](weekday-index-property-project.md)|
|[Name](weekday-name-property-project.md)|
|[Parent](weekday-parent-property-project.md)|
|[Shift1](weekday-shift1-property-project.md)|
|[Shift2](weekday-shift2-property-project.md)|
|[Shift3](weekday-shift3-property-project.md)|
|[Shift4](weekday-shift4-property-project.md)|
|[Shift5](weekday-shift5-property-project.md)|
|[Working](weekday-working-property-project.md)|

