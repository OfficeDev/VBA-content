---
title: Months Object (Project)
ms.prod: project-server
ms.assetid: 5db0ed37-cc23-7bc8-ebe5-fdaf6275b5db
ms.date: 06/08/2017
---


# Months Object (Project)

Contains a collection of  **[Month](month-object-project.md)** objects.
 


## Remarks

Use  **Months** (*Index* ), where*Index* is the month index number, month name, or **PjMonth** constant, to return a single **Month** object.
 

 

## Example

 **Using the Months Collection Object**
 

 
The following example counts the number of working days in each month of 2012 for each selected resource. 
 

 



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


## Properties



|**Name**|
|:-----|
|[Application](months-application-property-project.md)|
|[Count](months-count-property-project.md)|
|[Item](months-item-property-project.md)|
|[Parent](months-parent-property-project.md)|

## See also


#### Other resources


 
[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
