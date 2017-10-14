---
title: Display Address Entry Details for the Sender of a Message
ms.prod: outlook
ms.assetid: 6d8224a6-b565-699a-7e05-f0f9331bf089
ms.date: 06/08/2017
---


# Display Address Entry Details for the Sender of a Message

The recipient of each mail message deliverable by a transport provider has an address entry in the provider's hierarchy of address books for the session. This topic describes how to programmatically display the address entry information of the sender of a mail item that is currently displayed in an inspector.


1. For the currently displayed mail item, use the  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object to determine the Entry ID of the sender.
    
2. Use the  **[NameSpace.GetAddressEntryFromID](namespace-getaddressentryfromid-method-outlook.md)** method of the current session to return an **[AddressEntry](addressentry-object-outlook.md)** object.
    
3. Use the  **[AddressEntry.AddressEntryUserType](addressentry-addressentryusertype-property-outlook.md)** property to determine the type of the **AddressEntry**, and then display the details accordingly: 
    
      - If the address entry is a contact item in the Outlook Contacts folder, or if the SMTP address of the sender matches an e-mail address of one contact item in the default Contacts folder, then display the address entry information in a Contacts inspector. To match e-mail addresses in the Contacts folder, use the Table object to do a quick filter on the  **[ContactItem.Email1Address](contactitem-email1address-property-outlook.md)**,  **[ContactItem.Email2Address](contactitem-email2address-property-outlook.md)**, and  **[ContactItem.Email3Address](contactitem-email3address-property-outlook.md)** properties of items in that folder.
    
  - In all other cases, display the address entry information in the  **E-mail Properties** dialog box.
    

## Remarks

To run this code sample:


1. Open a mail message to have it displayed in the active inspector.
    
2. Place the code in the built-in  **ThisOutlookSession** module.
    
3. Run the  `TestAddressEntryDetails` procedure to display address entry details on the mail message in the active inspector:
    





```vb
Sub TestAddressEntryDetails() 
 Dim oMail As MailItem 
 
 Set oMail = Application.ActiveInspector.CurrentItem 
 DisplayAddressEntryDetails oMail 
End Sub 
 
 
Sub DisplayAddressEntryDetails(oM As MailItem) 
 Dim oPA As Outlook.PropertyAccessor 
 Dim oContact As Outlook.ContactItem 
 Dim oSender As Outlook.AddressEntry 
 Dim SenderID As String 
 
 'Create an instance of PropertyAccessor 
 Set oPA = oM.PropertyAccessor 
 
 'Obtain PidTagSenderEntryId and convert to string 
 SenderID = oPA.BinaryToString _ 
 (oPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C190102")) 
 
 'Obtain AddressEntry Object of the sender 
 Set oSender = Application.Session.GetAddressEntryFromID(SenderID) 
 
 'Examine AddressEntryUserType 
 If oSender.AddressEntryUserType = olOutlookContactAddressEntry Then 
 'Obtain ContactItem for AddressEntry 
 Set oContact = oSender.GetContact 
 oContact.Display 
 'Display details for Exchange or SMTP sender 
 Else 
 oSender.Details 
 End If 
End Sub
```


