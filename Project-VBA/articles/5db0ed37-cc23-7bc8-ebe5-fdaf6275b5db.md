
# Months Object (Project)

 **Last modified:** July 28, 2015

Contains a collection of  ** [Month](5ee32f12-72aa-fa16-ead2-97949005cd7c.md)** objects.

## Remarks

Use  **Months** ( _Index_ ), where _Index_ is the month index number, month name, or **PjMonth** constant, to return a single **Month** object.


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

Use the  ** [Months](615a4f5c-bda7-f684-1c29-d8003badf3a8.md)** property to return a **Months** collection. The following example counts the number of months in 2012.




```
ActiveProject.Calendar.Years(2012).Months.Count
```


## See also


#### Concepts


 [Project Object Model](900b167b-88ec-ea88-15b7-27bb90c22ac6.md)
