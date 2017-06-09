---
title: Account.GetAddressEntryFromID Method (Outlook)
keywords: vbaol11.chm3427
f1_keywords:
- vbaol11.chm3427
ms.prod: outlook
api_name:
- Outlook.Account.GetAddressEntryFromID
ms.assetid: 5aa9c67e-579f-5519-ed38-c80009cf506b
ms.date: 06/08/2017
---


# Account.GetAddressEntryFromID Method (Outlook)

Returns an  **[AddressEntry](addressentry-object-outlook.md)** object that represents the address entry specified by the given entry ID.


## Syntax

 _expression_ . **GetAddressEntryFromID**( **_ID_** )

 _expression_ A variable that represents an **[Account](account-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ID_|Required| **String**|Used to identify an address entry that is maintained for the session.|

### Return Value

An  **AddressEntry** that has the **[ID](addressentry-id-property-outlook.md)** property that matches the specified _ID_ .


## Remarks

This method is similar to the  **[GetAddressEntryFromID](namespace-getaddressentryfromid-method-outlook.md)** method of the **[NameSpace](namespace-object-outlook.md)** object, but has some additional contextual information about which account to use for the look-up. If there are multiple Microsoft Exchange accounts in the current profile, use the **GetAddressEntryFromID** method for the corresponding account.

The  **ID** property for an **AddressEntry** is a permanent, unique string identifier that the transport provider assigns when an **AddressEntry** is created. Outlook maintains a hierarchy of address books for a session, and the address entry that is returned must match the given ID and be in one of the address books.

 **GetAddressEntryFromID** returns an error if no item with the given ID can be found, if no connection is available, or if the user is set to work offline.


## See also


#### Concepts


[Account Object](account-object-outlook.md)

