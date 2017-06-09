---
title: Month Object (Project)
ms.prod: project-server
api_name:
- Project.Month
ms.assetid: 5ee32f12-72aa-fa16-ead2-97949005cd7c
ms.date: 06/08/2017
---


# Month Object (Project)

Represents a month in a year. The  **Month** object is a member of the **[Months](months-object-project.md)** collection.
 


## Example

 **Using the Month Object**
 

 
Use  **Months** (*Index* ), where*Index* is the month index number, month name, or **PjMonth** constant, to return a single **Month** object. The following example counts the number of working days in each month of 2012 for each selected resource.
 

 



```
Dim R As Resource 
Dim D As Integer, M As Integer, WorkingDays As Integer 
 
For Each R In ActiveSelection.Resources() 
    WorkingDays = 0 

    With R.Calendar.Years(2012) 
        For M = 1 To .Months.Count 
            WorkingDays = 0 
            For D = 1 To .Months(M).Days.Count 
                If .Months(M).Days(D).Working = True Then 
                    WorkingDays = WorkingDays + 1 
                End If 
            Next D 

            MsgBox "There are " &amp; WorkingDays &amp; " working days in " &amp; _
                .Months(M).Name &amp; " for " &amp; R.Name &amp; "." 
        Next M 
    End With 
Next R
```

 **Using the Months Collection**
 

 
Use the  **[Months](year-months-property-project.md)** property to return a **Months** collection. The following example counts the number of months in 2012.
 

 



```
ActiveProject.Calendar.Years(2012).Months.Count
```


## Methods



|**Name**|
|:-----|
|[Default](month-default-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](month-application-property-project.md)|
|[Calendar](month-calendar-property-project.md)|
|[Count](month-count-property-project.md)|
|[Days](month-days-property-project.md)|
|[Index](month-index-property-project.md)|
|[Name](month-name-property-project.md)|
|[Parent](month-parent-property-project.md)|
|[Shift1](month-shift1-property-project.md)|
|[Shift2](month-shift2-property-project.md)|
|[Shift3](month-shift3-property-project.md)|
|[Shift4](month-shift4-property-project.md)|
|[Shift5](month-shift5-property-project.md)|
|[Working](month-working-property-project.md)|

