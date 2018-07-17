---
title: Exceptions Object (Outlook)
keywords: vbaol11.chm289
f1_keywords:
- vbaol11.chm289
ms.prod: outlook
api_name:
- Outlook.Exceptions
ms.assetid: fa3b6c2e-33b0-0f04-4e60-af2c582f2caa
ms.date: 06/08/2017
---


# Exceptions Object (Outlook)

Contains a group of  **[Exception](exception-object-outlook.md)** objects.


## Remarks

If you have a recurring  **[AppointmentItem](appointmentitem-object-outlook.md)**, the **[RecurrencePattern](recurrencepattern-object-outlook.md)** object defines the recurrence of these appointments. The **Exceptions** object contains the group of **Exception** objects that define the exceptions to that series of appointments.

 **Exception** objects are added to the **Exceptions** object whenever a property in the corresponding **AppointmentItem** object is altered.


## Example

The following example sets a reference to the  **Exceptions** object.


```
Set myExceptions = myRecurrencePattern.Exceptions
```


## Methods



|**Name**|
|:-----|
|[Item](exceptions-item-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](exceptions-application-property-outlook.md)|
|[Class](exceptions-class-property-outlook.md)|
|[Count](exceptions-count-property-outlook.md)|
|[Parent](exceptions-parent-property-outlook.md)|
|[Session](exceptions-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
