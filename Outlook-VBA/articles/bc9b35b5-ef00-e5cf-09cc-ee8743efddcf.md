
# RecurrencePattern.RecurrenceType Property (Outlook)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets an  ** [OlRecurrenceType](63bc267e-6b9d-2cb5-3a96-4beb41afff72.md)** constant specifying the frequency of occurrences for the recurrence pattern. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **RecurrenceType**

 _expression_A variable that represents a  **RecurrencePattern** object.


## Remarks
<a name="sectionSection1"> </a>

You must set the  **RecurrenceType** property before you set other properties for a ** [RecurrencePattern](36c098f7-59fb-879a-5173-ed0260d13fa4.md)** object. The **RecurrencePattern** properties that you can set subsequently depends on the value of **RecurrenceType**, as shown in the following table:



| **OlRecurrenceType**| **Valid RecurrencePattern Properties**|
| **olRecursDaily**| ** [Duration](91cceed3-fd56-bae3-ee00-16f4b02eb2e3.md)**,  ** [EndTime](7babda13-9e57-4c80-1ab3-56025753ed9d.md)**,  ** [Interval](e3220174-38dc-d1e3-8d26-b3f208b554a4.md)**,  ** [NoEndDate](47c5841a-c0d2-2b06-ec73-7093779ceafa.md)**,  ** [Occurrences](a99a8a1c-dcd3-e96d-6091-0a005ca3b05f.md)**,  ** [PatternStartDate](20c82dbd-a622-91b6-618c-7cbe8bff2ca7.md)**,  ** [PatternEndDate](0f78ea71-3d92-2d38-be10-e05ab7bcf44a.md)**,  ** [StartTime](557e0f8d-c95d-e1f9-91a2-0734248d8628.md)**|
| **olRecursWeekly**| ** [DayOfWeekMask](79268798-90ab-4161-5a6e-97669daa475a.md)**,  **Duration**,  **EndTime**,  **Interval**,  **NoEndDate**,  **Occurrences**,  **PatternStartDate**,  **PatternEndDate**,  **StartTime**|
| **olRecursMonthly**| ** [DayOfMonth](d89a9a55-060c-d25d-4bf6-21e345da36d1.md)**,  **Duration**,  **EndTime**,  **Interval**,  **NoEndDate**,  **Occurrences**,  **PatternStartDate**,  **PatternEndDate**,  **StartTime**|
| **olRecursMonthNth**| **DayOfWeekMask**,  **Duration**,  **EndTime**,  **Interval**,  ** [Instance](3458aeff-97b7-02f8-e352-203ecc92dedd.md)**,  **NoEndDate**,  **Occurrences**,  **PatternStartDate**,  **PatternEndDate**,  **StartTime**|
| **olRecursYearly**| **DayOfMonth**,  **Duration**,  **EndTime**,  **Interval**,  **MonthOfYear**,  **NoEndDate**,  **Occurrences**,  **PatternStartDate**,  **PatternEndDate**,  **StartTime**|
| **olRecursYearNth**| **DayOfWeekMask**,  **Duration**,  **EndTime**,  **Interval**,  **Instance**,  **NoEndDate**,  **Occurrences**,  **PatternStartDate**,  **PatternEndDate**,  **StartTime**|

## Example
<a name="sectionSection2"> </a>

This Visual Basic for Applications example uses  ** [GetRecurrencePattern](a9f67c5b-a77f-4e34-e654-d12560a6dba0.md)** to obtain the ** [RecurrencePattern](36c098f7-59fb-879a-5173-ed0260d13fa4.md)** object for the newly-created ** [AppointmentItem](204a409d-654e-27aa-643a-8344c631b82d.md)**. The properties,  **RecurrenceType** , **DayOfWeekMask**,  ** [MonthOfYear](14112950-1e2a-a99a-7c48-3e76358de645.md)**,  **Instance**,  **Occurences**,  **StartTime**,  **EndTime**, and  ** [Subject](57f0f242-6d04-175f-4ea2-25145787f5bd.md)** are set, the appointment is saved and then displayed with the pattern: "Occurs the first Monday of June effective 6/1/2007 until 6/6/2016 from 2:00 PM to 5:00 PM."


```
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
<a name="sectionSection2"> </a>


#### Concepts


 [RecurrencePattern Object](36c098f7-59fb-879a-5173-ed0260d13fa4.md)
#### Other resources


 [RecurrencePattern Object Members](d282fdb2-2b6d-983d-fe5f-698113d35f89.md)
