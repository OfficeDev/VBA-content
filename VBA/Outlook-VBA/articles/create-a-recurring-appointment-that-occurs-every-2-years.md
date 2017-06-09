---
title: Create a Recurring Appointment that Occurs Every 2 Years
ms.prod: outlook
ms.assetid: ce15c1ad-2029-413f-4f03-8206ba7b112d
ms.date: 06/08/2017
---


# Create a Recurring Appointment that Occurs Every 2 Years

This topic shows a Visual Basic for Applications (VBA) code example that creates an appointment that occurs in the following pattern:


- Starts at 2 P.M. and ends at 5 P.M.
    
- Occurs on the last Monday of June.
    
- Occurs every other year for three instances.
    
- Becomes effective June 1, 2009.
    



 The code example results in a recurring appointment from 2 P.M. to 5 P.M., on the last Monday of June in 2009 (June 29, 2009), 2011 (June 27, 2011), and 2013 (June 24, 2013). The appointment is saved in the default calendar and is then displayed.



```vb
Sub RecurringYearNth() 
 Dim oAppt As AppointmentItem 
 Dim oPattern As RecurrencePattern 
 Set oAppt = Application.CreateItem(olAppointmentItem) 
 Set oPattern = oAppt.GetRecurrencePattern 
 With oPattern 
 ' Appointment occurs every n-th year (with n indicated by the Interval property). 
 .RecurrenceType = olRecursYearNth 
 ' Appointment occurs on Monday. 
 .DayOfWeekMask = olMonday 
 ' Appointment occurs in June. 
 .MonthOfYear = 6 
 ' Appointment occurs on the 5th or last Monday (per the DayOfWeekMask property). 
 .Instance = 5 
 ' Appointment occurs three times. 
 .Occurrences = 3 
 ' Appointment lasts for 180 minutes each time. 
 .Duration = 180 
 ' Appointment becomes effective on June 1, 2009. 
 .PatternStartDate = #6/1/2009# 
 ' Appointment starts at 2 P.M. 
 .StartTime = #2:00:00 PM# 
 ' Appointment ends at 5 P.M. 
 .EndTime = #5:00:00 PM# 
 ' Appointment recurs every 2 years (per a RecurrenceType of olRecursYearNth). 
 .Interval = 2 
 End With 
 oAppt.Subject = "Recurring every 2 years YearNth Appointment" 
 oAppt.Save 
 oAppt.Display 
End Sub
```


