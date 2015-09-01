
# Map a Display Name to an E-mail Address

 **Last modified:** July 28, 2015

 _**Applies to:** Outlook 2013_

This topic shows a code sample in Visual Basic for Applications (VBA) that takes a display name and tries to map it to an e-mail address known to the messaging system in the current session. 

For each Outlook session, the transport provider defines a set of address book containers that the messaging system can deliver messages to. Each address book container corresponds to an address list in Outlook. If a display name is defined in the set of address book containers, the display name can be resolved in the current session and there exists an entry in an address list that maps to this display name. Note that entries in an address list can be of various types, including an Exchange user and Exchange distribution list.
In this code sample, the function  `ResolveDisplayNameToSMTP` uses the display name "Dan Wilson" as an example. It first tries to verify that the display name is defined in an address list by creating a ** [Recipient](8cee4d79-ec55-52a4-710b-6456944ca86d.md)** object based on this display name and then calling ** [Recipient.Resolve](2c4f9243-2e31-642e-78a7-fe74cd73b385.md)**. If the name is resolved, then  `ResolveDisplayNameToSMTP` uses the ** [AddressEntry](d4a0a85e-8bab-bc56-57bc-d70c3c570c8e.md)** object that is mapped to the **Recipient** object to further obtain the type and, if possible, the e-mail address:

- If the type of the  **AddressEntry** object is an Exchange user, `ResolveDisplayNameToSMTP` calls ** [AddressEntry.GetExchangeUser](eaaafd52-42c9-7f6b-1acb-0b987496d604.md)** to obtain the corresponding ** [ExchangeUser](6ec117d1-7fdb-aa36-b567-1242f8238df0.md)** object. ** [ExchangeUser.PrimarySmtpAddress](2dda21da-44a2-fbfe-babc-58646c76689d.md)** provides the e-mail address that maps to the display name.
    
- If the  **AddressEntry** object is an Exchange distribution list, `ResolveDisplayNameToSMTP` calls ** [AddressEntry.GetExchangeDistributionList](060ac302-b916-d85d-5ba8-c682894129e2.md)** to obtain an ** [ExchangeDistributionList](2830dfba-6c0a-a81f-6b98-92ac2aafb59d.md)** object. ** [ExchangeDistributionList.PrimarySmtpAddress](f64bbc29-14c4-be68-402a-16d9ac34a727.md)** provides the e-mail address that maps to the display name.
    




```
Sub ResolveDisplayNameToSMTP() 
 Dim oRecip As Outlook.Recipient 
 Dim oEU As Outlook.ExchangeUser 
 Dim oEDL As Outlook.ExchangeDistributionList 
 
 Set oRecip = Application.Session.CreateRecipient("Dan Wilson") 
 oRecip.Resolve 
 If oRecip.Resolved Then 
 Select Case oRecip.AddressEntry.AddressEntryUserType 
 Case OlAddressEntryUserType.olExchangeUserAddressEntry 
 Set oEU = oRecip.AddressEntry.GetExchangeUser 
 If Not (oEU Is Nothing) Then 
 Debug.Print oEU.PrimarySmtpAddress 
 End If 
 Case OlAddressEntryUserType.olExchangeDistributionListAddressEntry 
 Set oEDL = oRecip.AddressEntry.GetExchangeDistributionList 
 If Not (oEDL Is Nothing) Then 
 Debug.Print oEDL.PrimarySmtpAddress 
 End If 
 End Select 
 End If 
End Sub
```

