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



| <strong>OlRecurrenceType</strong>| <strong>Valid RecurrencePattern Properties</strong>|
| 
<strong>olRecursDaily</strong>| <strong><a href="recurrencepattern-duration-property-outlook.md" data-raw-source="[Duration](recurrencepattern-duration-property-outlook.md)">Duration</a></strong> , <strong><a href="recurrencepattern-endtime-property-outlook.md" data-raw-source="[EndTime](recurrencepattern-endtime-property-outlook.md)">EndTime</a></strong> , <strong><a href="recurrencepattern-interval-property-outlook.md" data-raw-source="[Interval](recurrencepattern-interval-property-outlook.md)">Interval</a></strong> , <strong><a href="recurrencepattern-noenddate-property-outlook.md" data-raw-source="[NoEndDate](recurrencepattern-noenddate-property-outlook.md)">NoEndDate</a></strong> , <strong><a href="recurrencepattern-occurrences-property-outlook.md" data-raw-source="[Occurrences](recurrencepattern-occurrences-property-outlook.md)">Occurrences</a></strong> , <strong><a href="recurrencepattern-patternstartdate-property-outlook.md" data-raw-source="[PatternStartDate](recurrencepattern-patternstartdate-property-outlook.md)">PatternStartDate</a></strong> , <strong><a href="recurrencepattern-patternenddate-property-outlook.md" data-raw-source="[PatternEndDate](recurrencepattern-patternenddate-property-outlook.md)">PatternEndDate</a></strong> , <strong><a href="recurrencepattern-starttime-property-outlook.md" data-raw-source="[StartTime](recurrencepattern-starttime-property-outlook.md)">StartTime</a></strong>|
| 
<strong>olRecursWeekly</strong>| <strong><a href="recurrencepattern-dayofweekmask-property-outlook.md" data-raw-source="[DayOfWeekMask](recurrencepattern-dayofweekmask-property-outlook.md)">DayOfWeekMask</a></strong> , <strong>Duration</strong> , <strong>EndTime</strong> , <strong>Interval</strong> , <strong>NoEndDate</strong> , <strong>Occurrences</strong> , <strong>PatternStartDate</strong> , <strong>PatternEndDate</strong> , <strong>StartTime</strong>|
| 
<strong>olRecursMonthly</strong>| <strong><a href="recurrencepattern-dayofmonth-property-outlook.md" data-raw-source="[DayOfMonth](recurrencepattern-dayofmonth-property-outlook.md)">DayOfMonth</a></strong> , <strong>Duration</strong> , <strong>EndTime</strong> , <strong>Interval</strong> , <strong>NoEndDate</strong> , <strong>Occurrences</strong> , <strong>PatternStartDate</strong> , <strong>PatternEndDate</strong> , <strong>StartTime</strong>|
| 
<strong>olRecursMonthNth</strong>| <strong>DayOfWeekMask</strong> , <strong>Duration</strong> , <strong>EndTime</strong> , <strong>Interval</strong> , <strong><a href="recurrencepattern-instance-property-outlook.md" data-raw-source="[Instance](recurrencepattern-instance-property-outlook.md)">Instance</a></strong> , <strong>NoEndDate</strong> , <strong>Occurrences</strong> , <strong>PatternStartDate</strong> , <strong>PatternEndDate</strong> , <strong>StartTime</strong>|
| 
<strong>olRecursYearly</strong>| <strong>DayOfMonth</strong> , <strong>Duration</strong> , <strong>EndTime</strong> , <strong>Interval</strong> , <strong>MonthOfYear</strong> , <strong>NoEndDate</strong> , <strong>Occurrences</strong> , <strong>PatternStartDate</strong> , <strong>PatternEndDate</strong> , <strong>StartTime</strong>|
| 
<strong>olRecursYearNth</strong>| <strong>DayOfWeekMask</strong> , <strong>Duration</strong> , <strong>EndTime</strong> , <strong>Interval</strong> , <strong>Instance</strong> , <strong>NoEndDate</strong> , <strong>Occurrences</strong> , <strong>PatternStartDate</strong> , <strong>PatternEndDate</strong> , <strong>StartTime</strong>|

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

