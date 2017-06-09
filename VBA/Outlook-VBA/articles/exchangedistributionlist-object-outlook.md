---
title: ExchangeDistributionList Object (Outlook)
keywords: vbaol11.chm3159
f1_keywords:
- vbaol11.chm3159
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList
ms.assetid: 2830dfba-6c0a-a81f-6b98-92ac2aafb59d
ms.date: 06/08/2017
---


# ExchangeDistributionList Object (Outlook)

The  **ExchangeDistributionList** object provides detailed information about an **[AddressEntry](addressentry-object-outlook.md)** that represents an Exchange distribution list.


## Remarks

 **ExchangeDistributionList** is a derived class of **AddressEntry**, and is returned instead of an **AddressEntry** when the caller performs a **QueryInterface** on the **AddressEntry**.

The  **AddressEntry.Members** property supports enumerating members of a distribution list. **ExchangeDistributionList** adds the first-class properties for **[Alias](exchangedistributionlist-alias-property-outlook.md)**, **[Comments](exchangedistributionlist-comments-property-outlook.md)**, and **[PrimarySmtpAddress](exchangedistributionlist-primarysmtpaddress-property-outlook.md)**. You can also access other properties specific to the Exchange distribution list that are not exposed in the object model through the **[PropertyAccessor](propertyaccessor-object-outlook.md)** object.

Some properties such as  **Comments** are read-write properties. Setting these properties requires the code to be running under an appropriate Exchange administrator account; without sufficient permissions, calling the **[ExchangeUser.Update](exchangeuser-update-method-outlook.md)** method will result in a "permission denied" error.


## Example

The following code sample shows how to obtain the names of the Exchange distribution lists that the current user's manager belongs to. It uses the  **[ExchangeUser.GetExchangeUserManager](exchangeuser-getexchangeusermanager-method-outlook.md)** method to obtain information about the user's manager, and uses **[ExchangeUser.GetMemberOfList](exchangeuser-getmemberoflist-method-outlook.md)** to obtain the distribution lists (represented by **ExchangeDistributionList** objects) that the manager has joined.


```
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


## Methods



|**Name**|
|:-----|
|[Delete](exchangedistributionlist-delete-method-outlook.md)|
|[Details](exchangedistributionlist-details-method-outlook.md)|
|[GetContact](exchangedistributionlist-getcontact-method-outlook.md)|
|[GetExchangeDistributionList](exchangedistributionlist-getexchangedistributionlist-method-outlook.md)|
|[GetExchangeDistributionListMembers](exchangedistributionlist-getexchangedistributionlistmembers-method-outlook.md)|
|[GetExchangeUser](exchangedistributionlist-getexchangeuser-method-outlook.md)|
|[GetFreeBusy](exchangedistributionlist-getfreebusy-method-outlook.md)|
|[GetMemberOfList](exchangedistributionlist-getmemberoflist-method-outlook.md)|
|[GetOwners](exchangedistributionlist-getowners-method-outlook.md)|
|[Update](exchangedistributionlist-update-method-outlook.md)|
|[GetUnifiedGroup](exchangedistributionlist-getunifiedgroup-method-outlook.md)|
|[GetUnifiedGroupFromStore](exchangedistributionlist-getunifiedgroupfromstore-method-outlook.md)|
|[IsUnifiedGroup](exchangedistributionlist-isunifiedgroup-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Address](exchangedistributionlist-address-property-outlook.md)|
|[AddressEntryUserType](exchangedistributionlist-addressentryusertype-property-outlook.md)|
|[Alias](exchangedistributionlist-alias-property-outlook.md)|
|[Application](exchangedistributionlist-application-property-outlook.md)|
|[Class](exchangedistributionlist-class-property-outlook.md)|
|[Comments](exchangedistributionlist-comments-property-outlook.md)|
|[DisplayType](exchangedistributionlist-displaytype-property-outlook.md)|
|[ID](exchangedistributionlist-id-property-outlook.md)|
|[Name](exchangedistributionlist-name-property-outlook.md)|
|[Parent](exchangedistributionlist-parent-property-outlook.md)|
|[PrimarySmtpAddress](exchangedistributionlist-primarysmtpaddress-property-outlook.md)|
|[PropertyAccessor](exchangedistributionlist-propertyaccessor-property-outlook.md)|
|[Session](exchangedistributionlist-session-property-outlook.md)|
|[Type](exchangedistributionlist-type-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
