---
title: AddressList Object (Outlook)
keywords: vbaol11.chm2022
f1_keywords:
- vbaol11.chm2022
ms.prod: outlook
api_name:
- Outlook.AddressList
ms.assetid: 84611afe-48b1-185b-df4b-0f004e7436ff
ms.date: 06/08/2017
---


# AddressList Object (Outlook)

Represents an address book that contains a set of  **[AddressEntry](addressentry-object-outlook.md)** objects.


## Remarks

The  **AddressList** object is an address book that contains a set of **[AddressEntry](addressentry-object-outlook.md)** objects.

The  **AddressList** object supplies a list of address entries to which a messaging system can deliver messages. An **AddressList** object represents one address book container available under the transport provider's address book hierarchy for the current session. The entire hierarchy is available through the parent **[AddressLists](http://msdn.microsoft.com/library/b8c5ce75-3030-0179-45bb-f44fe6628074%28Office.15%29.aspx)** object.


## Example

The following example retrieves an  **AddressList** object that represents the Personal Address List.


```
Set myAddressList = Application.Session.AddressLists("Personal Address Book")
```


## Methods



|**Name**|
|:-----|
|[GetContactsFolder](http://msdn.microsoft.com/library/9ea91624-bd7d-af64-7220-a2d9b659787a%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddressEntries](http://msdn.microsoft.com/library/53248439-4781-c084-0905-8fb99f2fb4a9%28Office.15%29.aspx)|
|[AddressListType](http://msdn.microsoft.com/library/3a62cdec-3d8d-3bcf-b2c3-f9dd496fd6e0%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/721c34fd-c9df-612e-52e1-b65a51a8f6f5%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/b2649892-a30f-165f-8352-17f14b5e3b3d%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/c0c6953f-5d99-a18a-a64f-b9446f38e774%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/0d0a3072-c39e-debb-04ef-313c8612b325%28Office.15%29.aspx)|
|[IsInitialAddressList](http://msdn.microsoft.com/library/cc3f1f6a-7377-6db1-2f7c-3baf9a7361db%28Office.15%29.aspx)|
|[IsReadOnly](http://msdn.microsoft.com/library/45d40efc-08c0-e2d7-572a-a5e60efb7d2f%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/313072e7-937f-d0d6-6372-9dbbaa488ce1%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/cb7f5779-bd69-74a8-1986-6c2dafce8d20%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/8cc763f0-e73f-97f9-5a30-e6f50b17ca2c%28Office.15%29.aspx)|
|[ResolutionOrder](http://msdn.microsoft.com/library/e92bd83f-349b-d6e7-a5fb-7a6d893406a0%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/ac7d208a-49c8-fe1a-ea33-f7c6d8a700d7%28Office.15%29.aspx)|

## See also


#### Other resources


[AddressList Object Members](http://msdn.microsoft.com/library/49ce35c2-400b-16b0-5f74-7f7d6260e45b%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
