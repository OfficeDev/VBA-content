---
title: Explorer.Selection Property (Outlook)
keywords: vbaol11.chm2770
f1_keywords:
- vbaol11.chm2770
ms.prod: outlook
api_name:
- Outlook.Explorer.Selection
ms.assetid: 11002043-9dab-a5ad-b36e-52ddb04c1859
ms.date: 06/08/2017
---


# Explorer.Selection Property (Outlook)

Returns a  **[Selection](selection-object-outlook.md)** object that contains the item or items that are selected in the explorer window. Read-only.


## Syntax

 _expression_ . **Selection**

 _expression_ A variable that represents an **[Explorer](explorer-object-outlook.md)** object.


## Remarks

The location of a selection in the explorer can be in the view list, the appointment list or task list in the To-Do Bar, or the daily tasks list in a calendar view. For more information, see the  **[Location](selection-location-property-outlook.md)** property.

The  **Selection** property does not include any conversation header objects. Call the **[Selection.GetSelection](selection-getselection-method-outlook.md)** method, providing **olConversationHeaders** as the argument, to obtain conversation header objects that are selected in the explorer.

If the current folder displays a folder home page, this property returns an empty collection. Also, if a group header such as  **Today**, or a conversation group header is selected, the  **[Count](selection-count-property-outlook.md)** property on the returned **Selection** object is zero.


## Example



The following Microsoft Visual Basic for Applications (VBA) example displays the sender of each selected item in the active explorer. It uses the  **Count** property and the **[Item](selection-item-method-outlook.md)** method of the **[Selection](selection-object-outlook.md)** object that is returned by the **[Explorer.Selection](explorer-selection-property-outlook.md)** property to display the senders of all messages that are selected in the active explorer.






```vb
Sub GetSelectedItems() 
 
 Dim myOlExp As Outlook.Explorer 
 
 Dim myOlSel As Outlook.Selection 
 
 Dim mySender As Outlook.AddressEntry 
 
 Dim oMail As Outlook.MailItem 
 
 Dim oAppt As Outlook.AppointmentItem 
 
 Dim oPA As Outlook.PropertyAccessor 
 
 Dim strSenderID As String 
 
 Const PR_SENT_REPRESENTING_ENTRYID As String = _ 
 
 "http://schemas.microsoft.com/mapi/proptag/0x00410102" 
 
 Dim MsgTxt As String 
 
 Dim x As Long 
 
 
 
 MsgTxt = "Senders of selected items:" 
 
 Set myOlExp = Application.ActiveExplorer 
 
 Set myOlSel = myOlExp.Selection 
 
 For x = 1 To myOlSel.Count 
 
 If myOlSel.Item(x).Class = OlObjectClass.olMail Then 
 
 ' For mail item, use the SenderName property. 
 
 Set oMail = myOlSel.Item(x) 
 
 MsgTxt = MsgTxt &; oMail.SenderName &; ";" 
 
 ElseIf myOlSel.Item(x).Class = OlObjectClass.olAppointment Then 
 
 ' For appointment item, use the Organizser property. 
 
 Set oAppt = myOlSel.Item(x) 
 
 MsgTxt = MsgTxt &; oAppt.Organizer &; ";" 
 
 Else 
 
 ' For other items, use the property accessor to get the sender ID, 
 
 ' then get the address entry to display the sender name. 
 
 Set oPA = myOlSel.Item(x).PropertyAccessor 
 
 strSenderID = oPA.GetProperty(PR_SENT_REPRESENTING_ENTRYID) 
 
 Set mySender = Application.Session.GetAddressEntryFromID(strSenderID) 
 
 MsgTxt = MsgTxt &; mySender.Name &; ";" 
 
 End If 
 
 Next x 
 
 Debug.Print MsgTxt 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

