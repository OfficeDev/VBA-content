---
title: RecurrencePattern.Interval Property (Outlook)
keywords: vbaol11.chm279
f1_keywords:
- vbaol11.chm279
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern.Interval
ms.assetid: e3220174-38dc-d1e3-8d26-b3f208b554a4
ms.date: 06/08/2017
---


# RecurrencePattern.Interval Property (Outlook)

Returns or sets a  **Long** specifying the number of units of a given recurrence type between occurrences. Read/write.


## Syntax

 _expression_ . **Interval**

 _expression_ A variable that represents a **RecurrencePattern** object.


## Remarks

The  **Interval** property must be set before setting **[PatternEndDate](recurrencepattern-patternenddate-property-outlook.md)** .

For example, setting the  **Interval** property to 2 and the **[RecurrenceType](recurrencepattern-recurrencetype-property-outlook.md)** property to **olRecursWeekly** would cause the pattern to occur every second week.

When  **RecurrenceType** is set to **olRecursYearNth** or **olRecursYear** , the **Interval** property indicates the number of years between occurrences. For example, **Interval** equals 1 indicates the recurrence is every year, **Interval** equals 2 indicates the recurrence is every 2 years, and so on.


## Example

This Visual Basic for Applications example uses  **[GetRecurrencePattern](appointmentitem-getrecurrencepattern-method-outlook.md)** to obtain the **[RecurrencePattern](recurrencepattern-object-outlook.md)** object for the newly-created **[AppointmentItem](appointmentitem-object-outlook.md)** . The properties, **[RecurrenceType](recurrencepattern-recurrencetype-property-outlook.md)** , **[DayOfWeekMask](recurrencepattern-dayofweekmask-property-outlook.md)** , **[PatternStartDate](recurrencepattern-patternstartdate-property-outlook.md)** , **[Interval](recurrencepattern-interval-property-outlook.md)** , **[PatternEndDate](recurrencepattern-patternenddate-property-outlook.md)** , and **[Subject](appointmentitem-subject-property-outlook.md)** are set, the appointment is saved and then displayed with the pattern: "Occurs every 3 week(s) on Monday effective 1/21/2003 until 12/21/2004 from 2:00 PM to 5:00 PM."


```vb
Sub CreateAppointment() 
 
 Dim myApptItem As AppointmentItem 
 
 Dim myRecurrPatt As RecurrencePattern 
 
 
 
 
 
 Set myApptItem = Application.CreateItem(olAppointmentItem) 
 
 Set myRecurrPatt = myApptItem.GetRecurrencePattern 
 
 myRecurrPatt.RecurrenceType = olRecursWeekly 
 
 myRecurrPatt.DayOfWeekMask = olMonday 
 
 myRecurrPatt.PatternStartDate = #1/21/2003 2:00:00 PM# 
 
 myRecurrPatt.Interval = 3 
 
 myRecurrPatt.PatternEndDate = #12/21/2004 5:00:00 PM# 
 
 myApptItem.Subject = "Important Appointment" 
 
 myApptItem.Save 
 
 myApptItem.Display 
 
 Set myOlApp = Nothing 
 
 Set myApptItem = Nothing 
 
 Set myRecurrPatt = Nothing 
 
End Sub
```


## See also


#### Concepts


[RecurrencePattern Object](recurrencepattern-object-outlook.md)
#### Other resources


[How to: Create an Appointment as a Meeting on the Calendar](http://msdn.microsoft.com/library/130b6ae1-d1a4-3805-7e9c-75543b93fff5%28Office.15%29.aspx)


