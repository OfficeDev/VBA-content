---
title: NameSpace.CompareEntryIDs Method (Outlook)
keywords: vbaol11.chm794
f1_keywords:
- vbaol11.chm794
ms.prod: outlook
api_name:
- Outlook.NameSpace.CompareEntryIDs
ms.assetid: 4e935803-9c73-03d2-17c9-dcaf169fdbbe
ms.date: 06/08/2017
---


# NameSpace.CompareEntryIDs Method (Outlook)

Returns a  **Boolean** value that indicates if two entry ID values refer to the same Outlook item.


## Syntax

 _expression_ . **CompareEntryIDs**( **_FirstEntryID_** , **_SecondEntryID_** )

 _expression_ An expression that returns a **NameSpace** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FirstEntryID_|Required| **String**|The first entry ID to be compared.|
| _SecondEntryID_|Required| **String**|The second entry ID to be compared.|

### Return Value

 **True** if the entry ID values refer to the same Outlook item; otherwise, **False** .


## Remarks

Entry identifiers cannot be compared directly because one object can be represented by two different binary values. Use this method to determine whether two entry identifiers represent the same object.


## Example

The following Visual Basic for Applications (VBA) example compares the entry identifier associated with the organizer of a specified  **[AppointmentItem](appointmentitem-object-outlook.md)** object with the entry identifier of a specified **[Recipient](recipient-object-outlook.md)** object, using the **CompareEntryIDs** method, and returns **True** if the organizer and the specified recipient represent the same user.


```vb
Function IsRecipientTheOrganizer( _ 
 
 ByVal Appt As Outlook.AppointmentItem, _ 
 
 ByVal Recipient As Outlook.Recipient) As Boolean 
 
 
 
 Dim objAddrEntry As Outlook.AddressEntry 
 
 Dim objPropAc As Outlook.PropertyAccessor 
 
 Dim strOrganizerEntryId As String 
 
 Dim bytResult() As Byte 
 
 Dim objRecipientUser As Outlook.ExchangeUser 
 
 Dim objOrganizerUser As Outlook.ExchangeUser 
 
 Dim blnReturn As Boolean 
 
 
 
 'Property tag for Organizer EntryID 
 
 Const PR_SENT_REPRESENTING_ENTRYID As String = _ 
 
 "http://schemas.microsoft.com/mapi/proptag/0x00410102" 
 
 
 
 ' Retrieve an AddressEntry object reference for the 
 
 ' specified recipient. 
 
 Set objAddrEntry = Recipient.AddressEntry 
 
 
 
 ' If the address entry represents an Exchange user 
 
 ' or Exchange remote user, retrieve an 
 
 ' ExchangeUser object reference for the sender and 
 
 ' compare the EntryID value of that object with 
 
 ' the EntryID of the specified recipient. 
 
 If objAddrEntry.AddressEntryUserType = _ 
 
 OlAddressEntryUserType.olExchangeUserAddressEntry _ 
 
 Or objAddrEntry.AddressEntryUserType = _ 
 
 OlAddressEntryUserType.olExchangeRemoteUserAddressEntry Then 
 
 
 
 ' Attempt to retrieve an ExchangeUser 
 
 ' object reference for the specified 
 
 ' recipient. 
 
 Set objRecipientUser = objAddrEntry.GetExchangeUser() 
 
 
 
 If objRecipientUser Is Nothing Then 
 
 ' An Exchange user could not be retrieved 
 
 ' for the specified recipient. 
 
 blnReturn = False 
 
 Else 
 
 ' Retrieve the EntryID property value of the organizer. 
 
 ' The Organizer property of the AppointmentItem object only 
 
 ' contains a string representation of the name of the 
 
 ' organizer, so the PR_SENT_REPRESENTING_ENTRYID property value 
 
 ' is instead retrieved, using the PropertyAccessor object 
 
 ' associated with the appointment item. 
 
 Set objPropAc = Appt.PropertyAccessor 
 
 bytResult = objPropAc.GetProperty( _ 
 
 PR_SENT_REPRESENTING_ENTRYID) 
 
 
 
 If Not IsEmpty(bytResult) Then 
 
 ' Convert the binary value retrieved from the 
 
 ' PR_SENT_REPRESENTING_ENTRYID property into 
 
 ' a string value for comparison. 
 
 strOrganizerEntryId = _ 
 
 objPropAc.BinaryToString(bytResult) 
 
 
 
 ' Attempt to retrieve an ExchangeUser 
 
 ' object reference for the organizer. 
 
 Set objOrganizerUser = Appt.Application.Session. _ 
 
 GetAddressEntryFromID(strOrganizerEntryId).GetExchangeUser() 
 
 
 
 If objOrganizerUser Is Nothing Then 
 
 ' An Exchange user could not be retrieved 
 
 ' for the organizer. 
 
 blnReturn = False 
 
 Else 
 
 ' Compare the EntryIDs of the organizer 
 
 ' and the specified recipient. 
 
 blnReturn = Appt.Application.Session. _ 
 
 CompareEntryIDs( _ 
 
 objRecipientUser.ID, _ 
 
 objOrganizerUser.ID) 
 
 End If 
 
 End If 
 
 End If 
 
 End If 
 
 
 
EndRoutine: 
 
 ' Clean up 
 
 Set objOrganizerUser = Nothing 
 
 Set objRecipientUser = Nothing 
 
 Set objAddrEntry = Nothing 
 
 Set objPropAc = Nothing 
 
 
 
 ' Return the results. 
 
 IsRecipientTheOrganizer = blnReturn 
 
 
 
 Exit Function 
 
 
 
ErrRoutine: 
 
 Debug.Print Err.Number &; " - " &; Err.Description, _ 
 
 vbOKOnly Or vbCritical, _ 
 
 "IsRecipientTheOrganizer" 
 
 
 
 GoTo EndRoutine 
 
End Function
```


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

