---
title: ExchangeUser.GetMemberOfList Method (Outlook)
keywords: vbaol11.chm2084
f1_keywords:
- vbaol11.chm2084
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.GetMemberOfList
ms.assetid: 1f4e8910-8998-85ab-05dc-d06f6fd323c3
ms.date: 06/08/2017
---


# ExchangeUser.GetMemberOfList Method (Outlook)

Returns an  **[AddressEntries](addressentries-object-outlook.md)** collection object that contains the **[AddressEntry](addressentry-object-outlook.md)** objects representing all the Exchange distribution lists to which the user belongs.


## Syntax

 _expression_ . **GetMemberOfList**

 _expression_ A variable that represents an **ExchangeUser** object.


### Return Value

An  **AddressEntries** collection object that represents the Exchange distribution lists to which the **[ExchangeUser](exchangeuser-object-outlook.md)** belongs. Returns an **AddressEntries** collection object with a count of zero (0) if the **ExchangeUser** is not a member of any Exchange distribution list.


## Remarks

 **GetMemberOfList** is an expensive operation in terms of performance if there is a slow connection to Exchange Server.


## Example

The following code sample shows how to obtain the names of the Exchange distribution lists to which the manager of the current user belongs. It uses the  **ExchangeUser** object to obtain specific Exchange user information such as the user's Exchange account alias, details about the user's manager, and the distribution lists that the user's manager has joined.


```vb
Sub ShowManagerDistLists() 
 
 Dim oAE As Outlook.AddressEntry 
 
 Dim oExUser As Outlook.ExchangeUser 
 
 Dim oDistListEntries As Outlook.AddressEntries 
 
 
 
 'Obtain the AddressEntry for CurrentUser 
 
 Set oExUser = _ 
 
 Application.Session.CurrentUser.AddressEntry.GetExchangeUser 
 
 
 
 'Obtain distribution lists that the user's manager has joined 
 
 Set oDistListEntries = oExUser.GetExchangeUserManager.GetMemberOfList 
 
 For Each oAE In oDistListEntries 
 
 If oAE.AddressEntryUserType = _ 
 
 olExchangeDistributionListAddressEntry Then 
 
 Debug.Print (oAE.name) 
 
 End If 
 
 Next 
 
End Sub 
```


## See also


#### Concepts


[ExchangeUser Object](exchangeuser-object-outlook.md)

