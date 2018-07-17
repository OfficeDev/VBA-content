---
title: WeekDays Object (Project)
ms.prod: project-server
ms.assetid: 757437a0-e2ff-0027-f044-87d1cb357f62
ms.date: 06/08/2017
---


# WeekDays Object (Project)

Contains a collection of  **[Weekday](weekday-object-project.md)** objects.
 


## Example

 **Using the Weekday Object**
 

 
Use  **Weekdays** (*Index* ), where*Index* is the weekday index number, three-letter abbreviation of the day name, or **PjWeekday** constant, to return a single **Weekday** object. The following example sets Friday (the sixth day of a week starting on Sunday) as a half-day by setting the start and finish times for the first shift and clearing the values of the second and third shifts.
 

 



```
With ActiveProject.Calendar.WeekDays(6) 

 .Shift1.Start = #8:00:00 AM# 

 .Shift1.Finish = #12:00:00 PM# 

 .Shift2.Clear 

 .Shift3.Clear 

End With
```

A much better way to return the same object is to use the predefined constant for Friday instead of the nonintuitive number 6. Thus, the first line of the preceding example would be as follows: 
 

 



```
With ActiveProject.Calendar.WeekDays(pjFriday)
```

 **Using the Weekdays Collection**
 

 
Use the  **[Weekdays](calendar-weekdays-property-project.md)** property to return a **Weekdays** collection.
 

 



```
ActiveProject.Calendar.WeekDays
```


## Properties



|**Name**|
|:-----|
|[Application](weekdays-application-property-project.md)|
|[Count](weekdays-count-property-project.md)|
|[Item](weekdays-item-property-project.md)|
|[Parent](weekdays-parent-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
