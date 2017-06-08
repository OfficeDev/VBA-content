---
title: Access Exchange User or Distribution List Information from the Address Book
ms.prod: outlook
ms.assetid: 077a8666-09c5-e641-0b9b-7d83133d931f
ms.date: 06/08/2017
---


# Access Exchange User or Distribution List Information from the Address Book

This topic describes the objects that support accessing information about an Exchange user or distribution list from the Address Book. 

The Address Book contains address lists of users, distribution lists, and other types of address entries, as enumerated by  **[OlAddressEntryUserType](oladdressentryusertype-enumeration-outlook.md)**. Specifically, the Exchange user address entry and the Exchange distribution list address entry have many of their properties exposed as explicit built-in properties in the Outlook object model through the  **[ExchangeUser](exchangeuser-object-outlook.md)** and **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)** objects. Both of these objects inherit from the **[AddressEntry](addressentry-object-outlook.md)** object. They also support specific methods that facilitate accessing information about these entry types.

## Exchange User

The  **ExchangeUser** object supports properties like **[OfficeLocation](exchangeuser-officelocation-property-outlook.md)**,  **[JobTitle](exchangeuser-jobtitle-property-outlook.md)**,  **[FirstName](exchangeuser-firstname-property-outlook.md)**, and  **[LastName](exchangeuser-lastname-property-outlook.md)** that the parent **AddressEntry** object does not support. You can access these properties directly through the **ExchangeUser** object. You can access other properties of the Exchange user that are not exposed in the object model using **[ExchangeUser.PropertyAccessor](exchangeuser-propertyaccessor-property-outlook.md)**.

The  **ExchangeUser** object also supports methods like **[GetDirectReports](exchangeuser-getdirectreports-method-outlook.md)**,  **[GetExchangeUserManager](exchangeuser-getexchangeusermanager-method-outlook.md)**, and  **[GetMemberOfList](exchangeuser-getmemberoflist-method-outlook.md)** to facilitate accessing information specific to this Exchange user, such as full **AddressEntry** information for the associated direct reports, manager, and distribution lists.


## Security

Certain properties like  **OfficeLocation** and **JobTitle** are read-write and can only be updated (using **[ExchangeUser.Update](exchangeuser-update-method-outlook.md)**) by code that is running under an appropriate Exchange administrator account.


## Exchange Distribution List

 The **ExchangeDistributionList** obect supports properties like **Alias**,  **[Comments](exchangedistributionlist-comments-property-outlook.md)**, and  **[PrimarySmtpAddress](exchangedistributionlist-primarysmtpaddress-property-outlook.md)** that the parent **AddressEntry** object does not support. Other properties of the Exchange distribution list that are not exposed in the object model are accessible through **[ExchangeDistributionList.PropertyAccessor](exchangedistributionlist-propertyaccessor-property-outlook.md)**.

The  **ExchangeDistributionList** object also supports methods like **[GetExchangeDistributionListMembers](exchangedistributionlist-getexchangedistributionlistmembers-method-outlook.md)**,  **[GetMemberOfList](exchangedistributionlist-getmemberoflist-method-outlook.md)** and **[GetOwners](exchangedistributionlist-getowners-method-outlook.md)** to facilitate accessing information specific to a distribution list, such as full **AddressEntry** information for the associated members in this distribution list, other distribution lists that this list is a member of, and owners of this list.


## Security

Certain properties like  **Comments** are read-write and can only be updated (using **[ExchangeDistributionList.Update](exchangedistributionlist-update-method-outlook.md)**) by code that is running under an appropriate Exchange administrator account.


