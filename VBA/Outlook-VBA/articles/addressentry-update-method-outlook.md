---
title: AddressEntry.Update Method (Outlook)
keywords: vbaol11.chm2053
f1_keywords:
- vbaol11.chm2053
ms.prod: outlook
api_name:
- Outlook.AddressEntry.Update
ms.assetid: 099d83cf-01ff-21f8-aabb-ccfd497bab24
ms.date: 06/08/2017
---


# AddressEntry.Update Method (Outlook)

Posts a change to the  **[AddressEntry](addressentry-object-outlook.md)** object in the messaging system.


## Syntax

 _expression_ . **Update**( **_MakePermanent_** , **_Refresh_** )

 _expression_ An expression that returns a **AddressEntry** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MakePermanent_|Optional| **Variant**| A value of **True** indicates that the property cache is flushed and all changes are committed in the underlying address book. A value of **False** indicates that the property cache is flushed but not committed to persistent storage. The default value is **True** .|
| _Refresh_|Optional| **Variant**|A value of  **True** indicates that the property cache is reloaded from the values in the underlying address book. A value of **False** indicates that the property cache is not reloaded. The default value is **False** .|

## Remarks

New entries or changes to existing entries are not persisted in the collection until the  **Update** method has been called with its _MakePermanent_ parameter set to **True** .

To flush the cache and then reload the values from the address book, call  **Update** with the _MakePermanent_ parameter set to **False** and the _Refresh_ parameter set to **True** .


## See also


#### Concepts


[AddressEntry Object](addressentry-object-outlook.md)

