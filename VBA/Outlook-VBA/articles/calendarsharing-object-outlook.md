---
title: CalendarSharing Object (Outlook)
keywords: vbaol11.chm3176
f1_keywords:
- vbaol11.chm3176
ms.prod: outlook
api_name:
- Outlook.CalendarSharing
ms.assetid: 37a8a15e-51c2-b1a0-7db6-cf2a1f4e8405
ms.date: 06/08/2017
---


# CalendarSharing Object (Outlook)

Represents a set of utilities for sharing calendar information.


## Remarks

You can use the  **[GetCalendarExporter](folder-getcalendarexporter-method-outlook.md)** method of a **[Folder](folder-object-outlook.md)** object that represents a calendar folder to create a **CalendarSharing** object. The **GetCalendarExporter** method can only be used on calendar folders. An error occurs if you use the method on **Folder** objects that represent other folder types.

You can use the  **[SaveAsICal](calendarsharing-saveasical-method-outlook.md)** method to save calendar information in an iCalendar (.ics) file for sharing a calendar as a URL, or use the **[ForwardAsICal](calendarsharing-forwardasical-method-outlook.md)** method to create a **[MailItem](mailitem-object-outlook.md)** for sharing a calendar as a payload.


 **Note**  The  **CalendarSharing** object only supports exporting the iCalendar format.


## Example

The following Visual Basic for Applications (VBA) example creates a  **CalendarSharing** object for the Calendar folder, then exports the contents of the entire folder (including attachments and private items) to an iCalendar calendar (.ics) file.


```
Public Sub ExportEntireCalendar() 
 
 
 
 Dim oNamespace As NameSpace 
 
 Dim oFolder As Folder 
 
 Dim oCalendarSharing As CalendarSharing 
 
 
 
 On Error GoTo ErrRoutine 
 
 
 
 ' Get a reference to the Calendar default folder 
 
 Set oNamespace = Application.GetNamespace("MAPI") 
 
 Set oFolder = oNamespace.GetDefaultFolder(olFolderCalendar) 
 
 
 
 ' Get a CalendarSharing object for the Calendar default folder. 
 
 Set oCalendarSharing = oFolder.GetCalendarExporter 
 
 
 
 ' Set the CalendarSharing object to export the contents of 
 
 ' the entire Calendar folder, including attachments and 
 
 ' private items, in full detail. 
 
 With oCalendarSharing 
 
 .CalendarDetail = olFullDetails 
 
 .IncludeAttachments = True 
 
 .IncludePrivateDetails = True 
 
 .IncludeWholeCalendar = True 
 
 End With 
 
 
 
 ' Export calendar to an iCalendar calendar (.ics) file. 
 
 oCalendarSharing.SaveAsICal "C:\SampleCalendar.ics" 
 
 
 
EndRoutine: 
 
 On Error GoTo 0 
 
 Set oCalendarSharing = Nothing 
 
 Set oFolder = Nothing 
 
 Set oNamespace = Nothing 
 
Exit Sub 
 
 
 
ErrRoutine: 
 
 Select Case Err.Number 
 
 Case 287 ' &amp;H0000011F 
 
 ' The user denied access to the Address Book. 
 
 ' This error occurs if the code is run by an 
 
 ' untrusted application, and the user chose not to 
 
 ' allow access. 
 
 MsgBox "Access to Outlook was denied by the user.", _ 
 
 vbOKOnly, _ 
 
 Err.Number &amp; " - " &amp; Err.Source 
 
 Case -2147467259 ' &amp;H80004005 
 
 ' Export failed. 
 
 ' This error typically occurs if the CalendarSharing 
 
 ' method cannot export the calendar information because 
 
 ' of conflicting property settings. 
 
 MsgBox Err.Description, _ 
 
 vbOKOnly, _ 
 
 Err.Number &amp; " - " &amp; Err.Source 
 
 Case -2147221233 ' &amp;H8004010F 
 
 ' Operation failed. 
 
 ' This error typically occurs if the GetCalendarExporter method 
 
 ' is called on a folder that doesn't contain calendar items. 
 
 MsgBox Err.Description, _ 
 
 vbOKOnly, _ 
 
 Err.Number &amp; " - " &amp; Err.Source 
 
 Case Else 
 
 ' Any other error that may occur. 
 
 MsgBox Err.Description, _ 
 
 vbOKOnly, _ 
 
 Err.Number &amp; " - " &amp; Err.Source 
 
 End Select 
 
 
 
 GoTo EndRoutine 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[ForwardAsICal](calendarsharing-forwardasical-method-outlook.md)|
|[SaveAsICal](calendarsharing-saveasical-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](calendarsharing-application-property-outlook.md)|
|[CalendarDetail](calendarsharing-calendardetail-property-outlook.md)|
|[Class](calendarsharing-class-property-outlook.md)|
|[EndDate](calendarsharing-enddate-property-outlook.md)|
|[Folder](calendarsharing-folder-property-outlook.md)|
|[IncludeAttachments](calendarsharing-includeattachments-property-outlook.md)|
|[IncludePrivateDetails](calendarsharing-includeprivatedetails-property-outlook.md)|
|[IncludeWholeCalendar](calendarsharing-includewholecalendar-property-outlook.md)|
|[Parent](calendarsharing-parent-property-outlook.md)|
|[RestrictToWorkingHours](calendarsharing-restricttoworkinghours-property-outlook.md)|
|[Session](calendarsharing-session-property-outlook.md)|
|[StartDate](calendarsharing-startdate-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
