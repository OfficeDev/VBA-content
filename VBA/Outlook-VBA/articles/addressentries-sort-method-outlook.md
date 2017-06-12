---
title: AddressEntries.Sort Method (Outlook)
keywords: vbaol11.chm37
f1_keywords:
- vbaol11.chm37
ms.prod: outlook
api_name:
- Outlook.AddressEntries.Sort
ms.assetid: 9b381837-9fe9-1041-8297-e8c8dbcdc2e4
ms.date: 06/08/2017
---


# AddressEntries.Sort Method (Outlook)

Sorts the collection of items by the specified property. The index for the collection is reset to 1 upon completion of this method.


## Syntax

 _expression_ . **Sort**( **_Property_** , **_Order_** )

 _expression_ A variable that represents an **AddressEntries** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Property_|Optional| **Variant**|The name of the property by which to sort, which may be enclosed in brackets, for example, "[CompanyName]". May not be a user-defined field, and may not be a multi-valued property, such as a category.|
| _Order_|Optional| **Variant**|The order for the specified address entries. Can be one of these  **OlSortOrder** constants: **olAscending** , **olDescending** , or **olSortNone** .|

## Remarks

 **Sort** only affects the order of items in a collection. It does not affect the order of items in an explorer view.


## See also


#### Concepts


[AddressEntries Object](addressentries-object-outlook.md)

