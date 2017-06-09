---
title: AddressEntries Object (Outlook)
keywords: vbaol11.chm24
f1_keywords:
- vbaol11.chm24
ms.prod: outlook
api_name:
- Outlook.AddressEntries
ms.assetid: db91b717-07c6-d1f2-c545-b766ee1f0c6b
ms.date: 06/08/2017
---


# AddressEntries Object (Outlook)

Contains a collection of addresses for an  **[AddressList](addresslist-object-outlook.md)** object.


## Remarks

The object may contain zero or more  **[AddressEntry](addressentry-object-outlook.md)** objects and provides access to the entries in a transport provider's address book container.


## Example

The following example sets a reference to an  **AddressEntries** object.






```
Set myNameSpace = Application.GetNameSpace("MAPI") 
 
Set myAddressList = myNameSpace.AddressLists("Personal Address Book") 
 
Set myAddressEntries = myAddressList.AddressEntries
```

You can also index directly into the  **AddressEntries** object, returning an **AddressEntry** object.




```
Set myAddressEntry = myAddressList.AddressEntries(index)
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/b4c37547-8fbd-b1e4-40f3-5cba3cffd6e9%28Office.15%29.aspx)|
|[GetFirst](http://msdn.microsoft.com/library/f8f03b6e-d79e-09b5-2f75-6886e699a4b3%28Office.15%29.aspx)|
|[GetLast](http://msdn.microsoft.com/library/22b54c0f-5167-ac76-0cff-7ee4a142e1b3%28Office.15%29.aspx)|
|[GetNext](http://msdn.microsoft.com/library/7579909c-90a2-660f-6cf5-039a441ccc93%28Office.15%29.aspx)|
|[GetPrevious](http://msdn.microsoft.com/library/3d5aa211-212e-9a97-58aa-47d4447c9f47%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/42156250-3e72-c82c-7038-12cfa02f5f0a%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/9b381837-9fe9-1041-8297-e8c8dbcdc2e4%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/2ba2a2e6-e584-935b-e24a-77b2d14ebd58%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/ee94c79e-ecff-cd35-cf5c-2733ef77d25e%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/24b5bdb3-174d-1366-b2f5-c8243c71fa9d%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/42155333-c917-a950-6162-0ddc8f3616d5%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/bdd2afb2-a4f7-e31b-9413-94ba7e6ca213%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[AddressEntries Object Members](http://msdn.microsoft.com/library/1a38c073-06f9-06ad-4483-21ad59143f14%28Office.15%29.aspx)
