---
title: RecurrencePattern.RecurrenceType Property (Outlook)
keywords: vbaol11.chm285
f1_keywords:
- vbaol11.chm285
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern.RecurrenceType
ms.assetid: bc9b35b5-ef00-e5cf-09cc-ee8743efddcf
ms.date: 06/08/2017
---


# RecurrencePattern.RecurrenceType Property (Outlook)

Returns or sets an  **[OlRecurrenceType](olrecurrencetype-enumeration-outlook.md)** constant specifying the frequency of occurrences for the recurrence pattern. Read/write.


## Syntax

 _expression_ . **RecurrenceType**

 _expression_ A variable that represents a **RecurrencePattern** object.


## Remarks

You must set the  **RecurrenceType** property before you set other properties for a **[RecurrencePattern](recurrencepattern-object-outlook.md)** object. The **RecurrencePattern** properties that you can set subsequently depends on the value of **RecurrenceType** , as shown in the following table:



| **OlRecurrenceType**| **Valid RecurrencePattern Properties**|
| **olRecursDaily**| **[Duration](recurrencepattern-duration-property-outlook.md)** , **[EndTime](recurrencepattern-endtime-property-outlook.md)** , **[Interval](recurrencepattern-interval-property-outlook.md)** , **[NoEndDate](recurrencepattern-noenddate-property-outlook.md)** , **[Occurrences](recurrencepattern-occurrences-property-outlook.md)** , **[PatternStartDate](recurrencepattern-patternstartdate-property-outlook.md)** , **[PatternEndDate](recurrencepattern-patternenddate-property-outlook.md)** , **[StartTime](recurrencepattern-starttime-property-outlook.md)**|
| **olRecursWeekly**| **[DayOfWeekMask](recurrencepattern-dayofweekmask-property-outlook.md)** , **Duration** , **EndTime** , **Interval** , **NoEndDate** , **Occurrences** , **PatternStartDate** , **PatternEndDate** , **StartTime**|
| **olRecursMonthly**| **[DayOfMonth](recurrencepattern-dayofmonth-property-outlook.md)** , **Duration** , **EndTime** , **Interval** , **NoEndDate** , **Occurrences** , **PatternStartDate** , **PatternEndDate** , **StartTime**|
| **olRecursMonthNth**| **DayOfWeekMask** , **Duration** , **EndTime** , **Interval** , **[Instance](recurrencepattern-instance-property-outlook.md)** , **NoEndDate** , **Occurrences** , **PatternStartDate** , **PatternEndDate** , **StartTime**|
| **olRecursYearly**| **DayOfMonth** , **Duration** , **EndTime** , **Interval** , **MonthOfYear** , **NoEndDate** , **Occurrences** , **PatternStartDate** , **PatternEndDate** , **StartTime**|
| **olRecursYearNth**| **DayOfWeekMask** , **Duration** , **EndTime** , **Interval** , **Instance** , **NoEndDate** , **Occurrences** , **PatternStartDate** , **PatternEndDate** , **StartTime**|

## Example

This Visual Basic for Applications example uses  **[GetRecurrencePattern](appointmentitem-getrecurrencepattern-method-outlook.md)** to obtain the **[RecurrencePattern](recurrencepattern-object-outlook.md)** object for the newly-created **[AppointmentItem](appointmentitem-object-outlook.md)** . The properties, **RecurrenceType** , **DayOfWeekMask** , **[MonthOfYear](recurrencepattern-monthofyear-property-outlook.md)** , **Instance** , **Occurences** , **StartTime** , **EndTime** , and **[Subject](appointmentitem-subject-property-outlook.md)** are set, the appointment is saved and then displayed with the pattern: "Occurs the first Monday of June effective 6/1/2007 until 6/6/2016 from 2:00 PM to 5:00 PM."


```vb
Sub RecurringYearNth() 
 
 Dim oAppt As AppointmentItem 
 
 Dim oPattern As RecurrencePattern 
 
 Set oAppt = Application.CreateItem(olAppointmentItem) 
 
 Set oPattern = oAppt.GetRecurrencePattern 
 
 With oPattern 
 
 .RecurrenceType = olRecursYearNth 
 
 .DayOfWeekMask = olMonday 
 
 .MonthOfYear = 6 
 
 .Instance = 1 
 
 .Occurrences = 10 
 
 .Duration = 180 
 
 .PatternStartDate = #6/1/2007# 
 
 .StartTime = #2:00:00 PM# 
 
 .EndTime = #5:00:00 PM# 
 
 End With 
 
 oAppt.Subject = "Recurring YearNth Appointment" 
 
 oAppt.Save 
 
 oAppt.Display 
 
End Sub
```


## See also


#### Concepts


[RecurrencePattern Object](recurrencepattern-object-outlook.md)

