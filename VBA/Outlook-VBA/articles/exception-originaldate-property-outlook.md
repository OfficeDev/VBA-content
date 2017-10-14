---
title: Exception.OriginalDate Property (Outlook)
keywords: vbaol11.chm303
f1_keywords:
- vbaol11.chm303
ms.prod: outlook
api_name:
- Outlook.Exception.OriginalDate
ms.assetid: 0777de75-b32d-fe23-03d8-bb3deb18a69e
ms.date: 06/08/2017
---


# Exception.OriginalDate Property (Outlook)

Returns a  **Date** indicating the original date and time of an **[AppointmentItem](appointmentitem-object-outlook.md)** before it was altered. This property will return the original date even if the **AppointmentItem** has been deleted. However, it will not return the original time if deletion has occurred. Read-only.


## Syntax

 _expression_ . **OriginalDate**

 _expression_ A variable that represents an **Exception** object.


## Example

This Visual Basic for Applications (VBA) example uses  **[CreateItem](application-createitem-method-outlook.md)** to create an **[AppointmentItem](appointmentitem-object-outlook.md)** object. The **[RecurrencePattern](recurrencepattern-object-outlook.md)** is obtained for this item using the **[GetRecurrencePattern](appointmentitem-getrecurrencepattern-method-outlook.md)** method. By setting the these properties: **[RecurrenceType](recurrencepattern-recurrencetype-property-outlook.md)** , **[PatternStartDate](recurrencepattern-patternstartdate-property-outlook.md)** , and **[PatternEndDate](recurrencepattern-patternenddate-property-outlook.md)** , the appointments are now a recurring series that occur on a daily basis for the period of one year. An **[Exception](exception-object-outlook.md)** object is created when one instance of this recurring appointment is obtained using the **[GetOccurrence](recurrencepattern-getoccurrence-method-outlook.md)** method and properties for this instance are altered. This exception to the series of appointments is obtained using the **GetRecurrencePattern** method to access the **[Exceptions](exceptions-object-outlook.md)** collection associated with this series. Message boxes display the original **[Subject](appointmentitem-subject-property-outlook.md)** and **[OriginalDate](exception-originaldate-property-outlook.md)** for this exception to the series of appointments and the current date, time, and subject for this exception.


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
 
 
 
 'Get the recurrence pattern for the master 
 
 'AppointmentItem. Access the collection of 
 
 'exceptions to the regular appointments. 
 
 Set myRecurrPatt = myApptItem.GetRecurrencePattern 
 
 Set myException = myRecurrPatt.Exceptions.item(1) 
 
 
 
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


[Exception Object](exception-object-outlook.md)

