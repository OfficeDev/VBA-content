---
title: AddressEntry Object (Outlook)
keywords: vbaol11.chm2037
f1_keywords:
- vbaol11.chm2037
ms.prod: outlook
api_name:
- Outlook.AddressEntry
ms.assetid: d4a0a85e-8bab-bc56-57bc-d70c3c570c8e
ms.date: 06/08/2017
---


# AddressEntry Object (Outlook)

Represents a person, group, or public folder to which the messaging system can deliver messages.


## Remarks

The  **AddressEntry** object is an address in an **[AddressEntries](addressentries-object-outlook.md)** object. Each **AddressEntry** object in the **AddressEntries** object holds information that represents a person, group, or public folder to which the messaging system can deliver messages.

Use  **AddressEntries** ( _index_ ), where _index_ is the index number of an address entry or a value used to match the default property of an address entry, to return a single **AddressEntry** object.


## Example

The following example sets a reference to an  **AddressEntry** object.


```
Set myAddressEntry = myRecipient.AddressEntry 
 

```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/5aea93e6-cf3f-897a-41dd-5c5bfd59d4bb%28Office.15%29.aspx)|
|[Details](http://msdn.microsoft.com/library/85457da6-c97a-387d-6c7e-40eb005b25aa%28Office.15%29.aspx)|
|[GetContact](http://msdn.microsoft.com/library/2364f180-475d-aff1-01e8-30a54e870404%28Office.15%29.aspx)|
|[GetExchangeDistributionList](http://msdn.microsoft.com/library/060ac302-b916-d85d-5ba8-c682894129e2%28Office.15%29.aspx)|
|[GetExchangeUser](http://msdn.microsoft.com/library/eaaafd52-42c9-7f6b-1acb-0b987496d604%28Office.15%29.aspx)|
|[GetFreeBusy](http://msdn.microsoft.com/library/8f3c7cbe-a4b5-ef5c-d7d3-1b38273f6f59%28Office.15%29.aspx)|
|[Update](http://msdn.microsoft.com/library/099d83cf-01ff-21f8-aabb-ccfd497bab24%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Address](http://msdn.microsoft.com/library/39241126-a652-47e0-17c9-4566efd7ca4f%28Office.15%29.aspx)|
|[AddressEntryUserType](http://msdn.microsoft.com/library/082ff106-c7c8-a505-fc82-170540d851fe%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/76593413-e1f0-0311-abe2-7efa7570edbb%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/59868f39-d557-aae2-49a9-0c6892122618%28Office.15%29.aspx)|
|[DisplayType](http://msdn.microsoft.com/library/d61f5e35-d4d7-17c7-08e3-c0c1e3ce3f1f%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/06c806f1-5ca8-c46e-399d-c307e9428866%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/30a754ab-6265-56e0-fbbf-55bec7fa1b11%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/191bc4b8-0e55-8676-569f-7fde61033298%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/6fc091ac-ee82-a246-952c-6a7e75051e9a%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/e2fdc0ed-a470-eca7-0709-ea7938df3516%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/23c9da02-e687-cc1a-b505-0644289362e9%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[AddressEntry Object Members](http://msdn.microsoft.com/library/74c88069-aec4-952b-556f-03873fbb488b%28Office.15%29.aspx)
