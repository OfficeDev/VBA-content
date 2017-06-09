---
title: Day Object (Project)
ms.prod: project-server
api_name:
- Project.Day
ms.assetid: 411fe04f-b68d-08c2-8b6c-f2c1e9927a34
ms.date: 06/08/2017
---


# Day Object (Project)

Represents a day in a month. The  **Day** object is a member of the **[Days](days-object-project.md)** collection.
 


## Example

 **Using the Day Object**
 

 
Use  **Days** (*Index* ), where*Index* is the day index number or **[PjWeekday](pjweekday-enumeration-project.md)** constant, to return a single **Day** object. The following example counts the number of working days in the month of September 2008 for each selected resource.
 

 



```
Dim R As Resource, D As Integer, WorkingDays As Integer 
 
For Each R In ActiveSelection.Resources() 
    WorkingDays = 0 
    With R.Calendar.Years(2008).Months(pjSeptember) 
        For D = 1 To .Days.Count 
            If .Days(D).Working = True Then 
                WorkingDays = WorkingDays + 1 
            End If 
        Next D 
    End With 
    MsgBox "There are " &amp; WorkingDays &amp; " working days in " _ 
        &amp; R.Name &amp; "'s calendar." 
Next R
```

 **Using the Days Collection**
 

 
Use the  **[Days](month-days-property-project.md)** property to return a **Days** collection. The following example counts the number of working days in the month of September 2008.
 

 



```
ActiveProject.Calendar.Years(2008).Months(pjSeptember).Days.Count
```


## Methods



|**Name**|
|:-----|
|[Default](day-default-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](day-application-property-project.md)|
|[Calendar](day-calendar-property-project.md)|
|[Count](day-count-property-project.md)|
|[Index](day-index-property-project.md)|
|[Name](day-name-property-project.md)|
|[Parent](day-parent-property-project.md)|
|[Shift1](day-shift1-property-project.md)|
|[Shift2](day-shift2-property-project.md)|
|[Shift3](day-shift3-property-project.md)|
|[Shift4](day-shift4-property-project.md)|
|[Shift5](day-shift5-property-project.md)|
|[Working](day-working-property-project.md)|

