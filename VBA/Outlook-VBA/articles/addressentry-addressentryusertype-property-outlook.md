---
title: AddressEntry.AddressEntryUserType Property (Outlook)
keywords: vbaol11.chm2057
f1_keywords:
- vbaol11.chm2057
ms.prod: outlook
api_name:
- Outlook.AddressEntry.AddressEntryUserType
ms.assetid: 082ff106-c7c8-a505-fc82-170540d851fe
ms.date: 06/08/2017
---


# AddressEntry.AddressEntryUserType Property (Outlook)

Returns a constant from the  **[OlAddressEntryUserType](oladdressentryusertype-enumeration-outlook.md)** enumeration representing the user type of the **[AddressEntry](addressentry-object-outlook.md)** . Read-only.


## Syntax

 _expression_ . **AddressEntryUserType**

 _expression_ A variable that represents an **AddressEntry** object.


## Remarks

 **AddressEntryUserType** provides a level of granularity for user types that is finer than that of **[AddressEntry.DisplayType](addressentry-displaytype-property-outlook.md)** . The **DisplayType** property does not distinguish users with different types of **AddressEntry** , such as an **AddressEntry** that has a Simple Mail Transfer Protocol (SMTP) e-mail address, a Lightweight Directory Access Protocol (LDAP) address, an Exchange user address, or an **AddressEntry** in the Outlook Contacts Address Book. All these entires have **olUser** as their **AddressEntry.DisplayType** .


## Example

The following code sample shows how to obtain the business phone number, office location, and job title for all Exchange user entries in the Exchange Global Address List. It first uses  **[AddressList.AddressListType](addresslist-addresslisttype-property-outlook.md)** to find the Global Address List. Since the Global Address List can contain multiple types of entries such as Exchange user, Exchange distribution list, and Exchange public folder, for each **AddressEntry** on that **[AddressList](addresslist-object-outlook.md)** , the code sample uses **AddressEntryUserType** to verify if the **AddressEntry** represents an Exchange user. After it finds an Exchange user, it obtains and prints the various pieces of data for the user.


```vb
Sub DemoAE() 
 
 Dim colAL As Outlook.AddressLists 
 
 Dim oAL As Outlook.AddressList 
 
 Dim colAE As Outlook.AddressEntries 
 
 Dim oAE As Outlook.AddressEntry 
 
 Dim oExUser As Outlook.ExchangeUser 
 
 Set colAL = Application.Session.AddressLists 
 
 For Each oAL In colAL 
 
 'Address list is an Exchange Global Address List 
 
 If oAL.AddressListType = olExchangeGlobalAddressList Then 
 
 Set colAE = oAL.AddressEntries 
 
 For Each oAE In colAE 
 
 If oAE.AddressEntryUserType = _ 
 
 olExchangeUserAddressEntry Then 
 
 Set oExUser = oAE.GetExchangeUser 
 
 Debug.Print(oExUser.JobTitle) 
 
 Debug.Print(oExUser.OfficeLocation) 
 
 Debug.Print(oExUser.BusinessTelephoneNumber) 
 
 End If 
 
 Next 
 
 End If 
 
 Next 
 
End Sub
```


## See also


#### Concepts


[AddressEntry Object](addressentry-object-outlook.md)

