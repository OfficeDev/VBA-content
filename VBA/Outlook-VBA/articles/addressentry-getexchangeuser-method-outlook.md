---
title: AddressEntry.GetExchangeUser Method (Outlook)
keywords: vbaol11.chm2056
f1_keywords:
- vbaol11.chm2056
ms.prod: outlook
api_name:
- Outlook.AddressEntry.GetExchangeUser
ms.assetid: eaaafd52-42c9-7f6b-1acb-0b987496d604
ms.date: 06/08/2017
---


# AddressEntry.GetExchangeUser Method (Outlook)

Returns an  **[ExchangeUser](exchangeuser-object-outlook.md)** object that represents the **[AddressEntry](addressentry-object-outlook.md)** if the **AddressEntry** belongs to an Exchange **[AddressList](addresslist-object-outlook.md)** object such as the Global Address List (GAL) and corresponds to an Exchange user.


## Syntax

 _expression_ . **GetExchangeUser**

 _expression_ A variable that represents an **AddressEntry** object.


### Return Value

An  **ExchangeUser** object that represents the **AddressEntry** . Returns **Null** ( **Nothing** in Visual Basic) if the **AddressEntry** object does not correspond to an Exchange user.


## Remarks

 You have to be connected to the Exchange server to use this method.

If a string passed using this method has a character set that is similar to an existing address entry, the return value may include an entry that is matched based on the first letter of the string passed.

For example, you pass the string "Jack" for an Exchange user who has an address entry "Jai" in his Outlook address book, but not "Jack". Even though the "Jack" entry is not available in the Outlook address book, the email address returned is "Jai" rather than "Null".


## Example

The following code sample shows how to obtain the business phone number, office location, and job title for all Exchange user entries in the Exchange Global Address List. It first uses  **[AddressList.AddressListType](addresslist-addresslisttype-property-outlook.md)** to find the Global Address List. For each **AddressEntry** on that **[AddressList](addresslist-object-outlook.md)** , it uses **AddressEntryUserType** to verify if the **AddressEntry** represents an Exchange user. After it finds an Exchange user, it uses **GetExchangeUser** to obtain and print the various pieces of data for the user.


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
 
 olExchangeUserAddressEntry _ 
 
 Or oAE.AddressEntryUserType = _ 
 
 olExchangeRemoteUserAddressEntry Then 
 
 Set oExUser = oAE.GetExchangeUser 
 
 Debug.Print (oExUser.JobTitle) 
 
 Debug.Print (oExUser.OfficeLocation) 
 
 Debug.Print (oExUser.BusinessTelephoneNumber) 
 
 End If 
 
 Next 
 
 End If 
 
 Next 
 
End Sub
```


## See also


#### Concepts


[AddressEntry Object](addressentry-object-outlook.md)

