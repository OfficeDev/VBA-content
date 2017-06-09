---
title: Years Object (Project)
ms.prod: project-server
ms.assetid: 3aa139cf-2fc2-7039-5659-8e2d833b5a4f
ms.date: 06/08/2017
---


# Years Object (Project)

Contains a collection of  **[Year](year-object-project.md)** objects.
 


## Remarks

The  **Years** collection in Project begins in 1984 and ends in 2149. In previous versions of Project, scheduling can run from 1984 to 2049.
 

 

## Examples

 **Using the Year Object**
 

 
Use  **Years** ( _Index_), where  _Index_ is the year index number, to return a single **Year** object. The following example counts the number of working days in the month of September 2012 for each selected resource.
 

 



```
Dim r As Resource
Dim d As Integer
Dim workingDays As Integer
Dim theMonth As PjMonth

theMonth = pjSeptember

For Each r In ActiveSelection.Resources()
    workingDays = 0
    With r.Calendar.Years(2012).Months(theMonth)
        For d = 1 To .Days.Count
            If .Days(d).Working = True Then
                workingDays = workingDays + 1
            End If
        Next d
    End With
    MsgBox "There are " &amp; workingDays &amp; " working days in " _
        &amp; r.Name &amp; "'s calendar for month " &amp; theMonth
Next r
```

 **Using the Years Collection**
 

 
Use the  **[Years](calendar-years-property-project.md)** property to return a **Years** collection. The following example lists all the years in the calendar of the active project.
 

 



```
Sub CountYears()
    Dim c As Long
    Dim temp As String
        
    For c = 1 To ActiveProject.Calendar.Years.Count
        temp = temp &amp; ListSeparator &amp; " " &amp; _
            ActiveProject.Calendar.Years(c + 1983).Name
    Next c
            
    MsgBox Right$(temp, Len(temp) - Len(ListSeparator &amp; " "))
End Sub
```

Figure 1 shows the results of the  **CountYears** macro.
 

 

**Figure 1. Getting the list of years available**

 
![Years available for project planning](images/pj15_VBA_Years.gif)
 

 

## Properties



|**Name**|
|:-----|
|[Application](years-application-property-project.md)|
|[Count](years-count-property-project.md)|
|[Item](years-item-property-project.md)|
|[Parent](years-parent-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
