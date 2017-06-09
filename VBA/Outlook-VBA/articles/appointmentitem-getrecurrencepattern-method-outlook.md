---
title: AppointmentItem.GetRecurrencePattern Method (Outlook)
keywords: vbaol11.chm905
f1_keywords:
- vbaol11.chm905
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.GetRecurrencePattern
ms.assetid: a9f67c5b-a77f-4e34-e654-d12560a6dba0
ms.date: 06/08/2017
---


# AppointmentItem.GetRecurrencePattern Method (Outlook)

Returns a  **[RecurrencePattern](recurrencepattern-object-outlook.md)** object that represents the recurrence attributes of an appointment.


## Syntax

 _expression_ . **GetRecurrencePattern**

 _expression_ A variable that represents an **AppointmentItem** object.


### Return Value

A  **RecurrencePattern** object that represents the recurrence attributes of an appointment.


## Remarks

If there is no existing recurrence pattern, a new empty  **RecurrencePattern** object is returned.

When you work with recurring appointment items, you should release any prior references, obtain new references to the recurring appointment item before you access or modify the item, and release these references as soon as you are finished and have saved the changes. This practice applies to the recurring  **[AppointmentItem](appointmentitem-object-outlook.md)** object, and any **[Exception](exception-object-outlook.md)** or **RecurrencePattern** object. To release a reference in Visual Basic for Applications (VBA) or Visual Basic, set that existing object to **Nothing** . In C#, explicitly release the memory for that object. For a code example, see the topic for the **AppointmentItem** object.

Note that even after you release your reference and attempt to obtain a new reference, if there is still an active reference, held by another add-in or Outlook, to one of the above objects, your new reference will still point to an out-of-date copy of the object. Therefore, it is important that you release your references as soon as you are finished with the recurring appointment.


## Example

This Visual Basic for Applications (VBA) example uses  **[CreateItem](application-createitem-method-outlook.md)** to create an **AppointmentItem** object. The **RecurrencePattern** is obtained for this item using the **GetRecurrencePattern** method. By setting the **RecurrencePattern** properties, **[RecurrenceType](recurrencepattern-recurrencetype-property-outlook.md)** , **[PatternStartDate](recurrencepattern-patternstartdate-property-outlook.md)** , and **[PatternEndDate](recurrencepattern-patternenddate-property-outlook.md)** , the appointments are now a recurring series that occur on a daily basis for the period of one year.

An  **Exception** object is created when one instance of this recurring appointment is obtained using the **[GetOccurrence](recurrencepattern-getoccurrence-method-outlook.md)** method and properties for this instance are altered. This exception to the series of appointments is obtained using the **GetRecurrencePattern** method to access the **[Exceptions](exceptions-object-outlook.md)** collection associated with this series. Message boxes display the original **[Subject](appointmentitem-subject-property-outlook.md)** and **[OriginalDate](exception-originaldate-property-outlook.md)** for this exception to the series of appointments and the current date, time, and subject for this exception.




```vb
Public Sub cmdExample() 
 
 Dim myApptItem As Outlook.AppointmentItem 
 
 Dim myRecurrPatt As Outlook.RecurrencePattern 
 
 Dim myNamespace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myItems As Outlook.Items 
 
 Dim myDate As Date 
 
 Dim myOddApptItem As Outlook.AppointmentItem 
 
 Dim saveSubject As String 
 
 Dim newDate As Date 
 
 Dim myException As Outlook.Exception 
 
 
 
 Set myApptItem = Application.CreateItem(olAppointmentItem) 
 
 myApptItem.Start = #2/2/2003 3:00:00 PM# 
 
 myApptItem.End = #2/2/2003 4:00:00 PM# 
 
 myApptItem.Subject = "Meet with Boss" 
 
 
 
 'Get the recurrence pattern for this appointment 
 
 'and set it so that this is a daily appointment 
 
 'that begins on 2/2/03 and ends on 2/2/04 
 
 'and save it. 
 
 Set myRecurrPatt = myApptItem.GetRecurrencePattern 
 
 myRecurrPatt.RecurrenceType = olRecursDaily 
 
 myRecurrPatt.PatternStartDate = #2/2/2003# 
 
 myRecurrPatt.PatternEndDate = #2/2/2004# 
 
 myApptItem.Save 
 
 
 
 'Access the items in the Calendar folder to locate 
 
 'the master AppointmentItem for the new series. 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNamespace.GetDefaultFolder(olFolderCalendar) 
 
 Set myItems = myFolder.Items 
 
 Set myApptItem = myItems("Meet with Boss") 
 
 
 
 'Get the recurrence pattern for this appointment 
 
 'and obtain the occurrence for 3/12/03. 
 
 myDate = #3/12/2003 3:00:00 PM# 
 
 Set myRecurrPatt = myApptItem.GetRecurrencePattern 
 
 Set myOddApptItem = myRecurrPatt.GetOccurrence(myDate) 
 
 
 
 'Save the existing subject. Change the subject and 
 
 'starting time for this particular appointment 
 
 'and save it. 
 
 saveSubject = myOddApptItem.Subject 
 
 myOddApptItem.Subject = "Meet NEW Boss" 
 
 newDate = #3/12/2003 3:30:00 PM# 
 
 myOddApptItem.Start = newDate 
 
 myOddApptItem.Save 
 
 
 
 'Release references to the appointment series 
 
 Set myApptItem = Nothing 
 
 Set myRecurrPatt = Nothing 
 
 
 
 'Get the recurrence pattern for the master 
 
 'AppointmentItem. Access the collection of 
 
 'exceptions to the regular appointments. 
 
 Set myItems = myFolder.Items 
 
 Set myApptItem = myItems("Meet with Boss") 
 
 
 
 Set myRecurrPatt = myApptItem.GetRecurrencePattern 
 
 Set myException = myRecurrPatt.Exceptions.Item(1) 
 
 
 
 'Display the original date, time, and subject 
 
 'for this exception. 
 
 MsgBox myException.OriginalDate &; ": " &; saveSubject 
 
 
 
 'Display the current date, time, and subject 
 
 'for this exception. 
 
 MsgBox myException.AppointmentItem.Start &; ": " &; _ 
 
 myException.AppointmentItem.Subject 
 
End Sub
```


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

