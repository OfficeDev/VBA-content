---
title: RecurrencePattern Object (Outlook)
keywords: vbaol11.chm268
f1_keywords:
- vbaol11.chm268
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern
ms.assetid: 36c098f7-59fb-879a-5173-ed0260d13fa4
ms.date: 06/08/2017
---


# RecurrencePattern Object (Outlook)

Represents the pattern of incidence of recurring appointments and tasks for the associated  **[AppointmentItem](appointmentitem-object-outlook.md)** and **[TaskItem](taskitem-object-outlook.md)** object.


## Remarks

Use the  **GetRecurrencePattern** method to return the **RecurrencePattern** object associated with an **AppointmentItem** or **TaskItem** object.

Calling  **GetRecurrencePattern** or **ClearRecurrencePattern** has the side effect of setting the **IsRecurring** property of the item accordingly. This property can be used as required for efficient filtering of the **[Items](items-object-outlook.md)** object.

The type of recurrence pattern is indicated by the  **[RecurrenceType](http://msdn.microsoft.com/library/bc9b35b5-ef00-e5cf-09cc-ee8743efddcf%28Office.15%29.aspx)** property. The **RecurrenceType** property is the first property you should set.

The following properties are valid for all recurrence patterns:  **[EndTime](http://msdn.microsoft.com/library/7babda13-9e57-4c80-1ab3-56025753ed9d%28Office.15%29.aspx)**, **[Occurrences](http://msdn.microsoft.com/library/a99a8a1c-dcd3-e96d-6091-0a005ca3b05f%28Office.15%29.aspx)**, **StartDate**, **[StartTime](http://msdn.microsoft.com/library/557e0f8d-c95d-e1f9-91a2-0734248d8628%28Office.15%29.aspx)**, or **Type**.

The following table shows the properties that are valid for the different recurrence types. An error occurs if the item is saved and the property is null or contains an invalid value. Monthly and yearly patterns are only valid for a single day. Weekly patterns are only valid as the  **Or** of the **[DayOfWeekMask](http://msdn.microsoft.com/library/79268798-90ab-4161-5a6e-97669daa475a%28Office.15%29.aspx)**.



|**RecurrenceType**|**Properties**|**Examples**|
|:-----|:-----|:-----|
|**olRecursDaily**|**[Duration](http://msdn.microsoft.com/library/91cceed3-fd56-bae3-ee00-16f4b02eb2e3%28Office.15%29.aspx)**, **EndTime**, **[Interval](http://msdn.microsoft.com/library/e3220174-38dc-d1e3-8d26-b3f208b554a4%28Office.15%29.aspx)**, **[NoEndDate](http://msdn.microsoft.com/library/47c5841a-c0d2-2b06-ec73-7093779ceafa%28Office.15%29.aspx)**, **Occurrences**, **[PatternStartDate](http://msdn.microsoft.com/library/20c82dbd-a622-91b6-618c-7cbe8bff2ca7%28Office.15%29.aspx)**, **[PatternEndDate](http://msdn.microsoft.com/library/0f78ea71-3d92-2d38-be10-e05ab7bcf44a%28Office.15%29.aspx)**, **StartTime**|A value N for  **Interval** is every N days.|
|**olRecursWeekly**|**DayOfWeekMask**, **Duration**, **EndTime**, **Interval**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|A value N for  **Interval** is every N weeks. An example of **DayofWeekMask** is every Tuesday, Wednesday, and Thursday.|
|**olRecursMonthly**|**[DayOfMonth](http://msdn.microsoft.com/library/d89a9a55-060c-d25d-4bf6-21e345da36d1%28Office.15%29.aspx)**, **Duration**, **EndTime**, **Interval**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|A value N for  **Interval** is every N months. A value N for **DayofMonth** is every Nth day of the month.|
|**olRecursMonthNth**|**DayOfWeekMask**, **Duration**, **EndTime**, **Interval**, **[Instance](http://msdn.microsoft.com/library/3458aeff-97b7-02f8-e352-203ecc92dedd%28Office.15%29.aspx)**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|A value N for  **Interval** is every N months. An example of value N for **Instance** is every Nth Tuesday. An example of **DayofWeekMask** is every Tuesday and Wednesday.|
|**olRecursYearly**|**DayOfMonth**, **Duration**, **EndTime**, **Interval**, **[MonthOfYear](http://msdn.microsoft.com/library/14112950-1e2a-a99a-7c48-3e76358de645%28Office.15%29.aspx)**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|A value N for  **DayofMonth** is the Nth day of the month. An example of **MonthOfYear** is February.|
|**olRecursYearNth**|**DayOfWeekMask**, **Duration**, **EndTime**, **Interval**, **Instance**, **NoEndDate**, **Occurrences**, **PatternStartDate**, **PatternEndDate**, **StartTime**|An example of value N for  **Instance** is the Nth Tuesday. An example of **DayofWeekMask** is Tuesday, Wednesday, and Thursday. An example of **MonthOfYear** is February.|
When you work with recurring appointment items, you should release any prior references, obtain new references to the recurring appointment item before you access or modify the item, and release these references as soon as you are finished and have saved the changes. This practice applies to the recurring  **AppointmentItem** object, and any **[Exception](http://msdn.microsoft.com/library/010552b0-9ba6-c81b-1e3a-fd6a681e5163%28Office.15%29.aspx)** or **[RecurrencePattern](recurrencepattern-object-outlook.md)** object. To release a reference in Visual Basic for Applications (VBA) or Visual Basic, set that existing object to **Nothing**. In C#, explicitly release the memory for that object. For a code example, see the topic for the **AppointmentItem** object.

Note that even after you release your reference and attempt to obtain a new reference, if there is still an active reference, held by another add-in or Outlook, to one of the above objects, your new reference will still point to an out-of-date copy of the object. Therefore, it is important that you release your references as soon as you are finished with the recurring appointment.


## Methods



|**Name**|
|:-----|
|[GetOccurrence](http://msdn.microsoft.com/library/2a0cd7d2-d16d-7b07-eb5d-43df0bbf022f%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/dd7068ee-385a-5bfc-fe15-f6a76e5441c9%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/64e7d0b9-9a86-7e81-1747-306c28bd5611%28Office.15%29.aspx)|
|[DayOfMonth](http://msdn.microsoft.com/library/d89a9a55-060c-d25d-4bf6-21e345da36d1%28Office.15%29.aspx)|
|[DayOfWeekMask](http://msdn.microsoft.com/library/79268798-90ab-4161-5a6e-97669daa475a%28Office.15%29.aspx)|
|[Duration](http://msdn.microsoft.com/library/91cceed3-fd56-bae3-ee00-16f4b02eb2e3%28Office.15%29.aspx)|
|[EndTime](http://msdn.microsoft.com/library/7babda13-9e57-4c80-1ab3-56025753ed9d%28Office.15%29.aspx)|
|[Exceptions](http://msdn.microsoft.com/library/e068565b-5418-897a-9f06-92e87a532144%28Office.15%29.aspx)|
|[Instance](http://msdn.microsoft.com/library/3458aeff-97b7-02f8-e352-203ecc92dedd%28Office.15%29.aspx)|
|[Interval](http://msdn.microsoft.com/library/e3220174-38dc-d1e3-8d26-b3f208b554a4%28Office.15%29.aspx)|
|[MonthOfYear](http://msdn.microsoft.com/library/14112950-1e2a-a99a-7c48-3e76358de645%28Office.15%29.aspx)|
|[NoEndDate](http://msdn.microsoft.com/library/47c5841a-c0d2-2b06-ec73-7093779ceafa%28Office.15%29.aspx)|
|[Occurrences](http://msdn.microsoft.com/library/a99a8a1c-dcd3-e96d-6091-0a005ca3b05f%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/81ecfc56-b35d-e99d-9312-6b83a8dc58bf%28Office.15%29.aspx)|
|[PatternEndDate](http://msdn.microsoft.com/library/0f78ea71-3d92-2d38-be10-e05ab7bcf44a%28Office.15%29.aspx)|
|[PatternStartDate](http://msdn.microsoft.com/library/20c82dbd-a622-91b6-618c-7cbe8bff2ca7%28Office.15%29.aspx)|
|[RecurrenceType](http://msdn.microsoft.com/library/bc9b35b5-ef00-e5cf-09cc-ee8743efddcf%28Office.15%29.aspx)|
|[Regenerate](http://msdn.microsoft.com/library/c1db398b-5f13-85e0-981d-795c8c7ac8ea%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/f30fce75-350c-6893-276a-47b19f211249%28Office.15%29.aspx)|
|[StartTime](http://msdn.microsoft.com/library/557e0f8d-c95d-e1f9-91a2-0734248d8628%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[RecurrencePattern Object Members](http://msdn.microsoft.com/library/d282fdb2-2b6d-983d-fe5f-698113d35f89%28Office.15%29.aspx)
