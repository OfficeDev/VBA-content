---
title: AppointmentItem.Respond Method (Outlook)
keywords: vbaol11.chm906
f1_keywords:
- vbaol11.chm906
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Respond
ms.assetid: 060d1fcb-0011-bea0-5c6b-fa3538ff9a2d
ms.date: 06/08/2017
---


# AppointmentItem.Respond Method (Outlook)

Responds to a meeting request.


## Syntax

 _expression_ . **Respond**( **_Response_** , **_fNoUI_** , **_fAdditionalTextDialog_** )

 _expression_ A variable that represents an **[AppointmentItem](appointmentitem-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **[OlMeetingResponse](olmeetingresponse-enumeration-outlook.md)**|The response to the request.|
| _fNoUI_|Optional| **Variant**| **True** to not display a dialog box; the response is sent automatically. **False** to display the dialog box for responding.|
| _fAdditionalTextDialog_|Optional| **Variant**| **False** to not prompt the user for input; the response is displayed in the inspector for editing. **True** to prompt the user to either send or send with comments. This argument is valid only if **fNoUI** is **False** .|

### Return Value

A  **[MeetingItem](meetingitem-object-outlook.md)** object that represents the response to the meeting request.


## Remarks

When you call the  **Respond** method with the **olMeetingAccepted** or **olMeetingTentative** parameter, Outlook will create a new appointment item that duplicates the original appointment item. The new item will have a different Entry ID. Outlook will then remove the original item. You should no longer use the Entry ID of the original item, but instead call the **[EntryID](appointmentitem-entryid-property-outlook.md)** property to obtain the Entry ID for the new item for any subsequent needs. This is to ensure that this appointment item will be properly synchronized on your calendar if more than one client computer accesses your calendar but may be offline using the cache mode occasionally.

The following table describes the behavior of the  **Respond** method depending on the parent object, and the _fNoUI_ and _fAdditionalTextDialog_ parameters.



|**_fNoUI, fAdditionalTextDialog_**|**_Result_**|
|:-----|:-----|
| **True, True**|Response item is returned with no user interface. To send the response, you must call the  **[Send](appointmentitem-send-method-outlook.md)** method.|
| **True, False**|Same result as with  **True, True** .|
| **False, True**|Prompts user to  **Send** or **Edit** before sending the response.|
| **False, False**|New response item appears in the user interface, but no prompt is displayed. |

## Example

This Visual Basic for Applications (VBA) example finds a  **MeetingItem** in the default **Inbox** folder and adds the associated appointment to the **Calendar** folder. It then responds to the sender by accepting the meeting.


```vb
Sub AcceptMeeting() 
 Dim myNameSpace As Outlook.NameSpace 
 Dim myFolder As Outlook.Folder 
 Dim myMtgReq As Outlook.MeetingItem 
 Dim myAppt As Outlook.AppointmentItem 
 Dim myMtg As Outlook.MeetingItem 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox) 
 Set myMtgReq = myFolder.Items.Find("[MessageClass] = 'IPM.Schedule.Meeting.Request'") 
 If TypeName(myMtgReq) <> "Nothing" Then 
 Set myAppt = myMtgReq.GetAssociatedAppointment(True) 
 Set myMtg = myAppt.Respond(olResponseAccepted, True) 
 myMtg.Send 
 End If 
End Sub
```


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

