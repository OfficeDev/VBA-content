---
title: ExchangeUser.GetContact Method (Outlook)
keywords: vbaol11.chm2078
f1_keywords:
- vbaol11.chm2078
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.GetContact
ms.assetid: 443fb23a-cd26-e385-bd9d-e978aec56458
ms.date: 06/08/2017
---


# ExchangeUser.GetContact Method (Outlook)

Returns  **Null** ( **Nothing** in Visual Basic) because the **[ExchangeUser](exchangeuser-object-outlook.md)** object does not correspond to a contact in a Contacts Address Book.


## Syntax

 _expression_ . **GetContact**

 _expression_ A variable that represents an **ExchangeUser** object.


### Return Value

 **Null** ( **Nothing** in Visual Basic) because the **ExchangeUser** object does not correspond to a contact in a Contacts Address Book.


## Remarks

The  **ExchangeUser** object is derived from the **[AddressEntry](addressentry-object-outlook.md)** object. It inherits the **GetContact** method from the **AddressEntry** object, and in the case of **ExchangeUser** , this method always returns **Null**.


## See also


#### Concepts


[ExchangeUser Object](exchangeuser-object-outlook.md)

