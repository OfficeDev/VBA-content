---
title: AddressEntries.Add Method (Outlook)
keywords: vbaol11.chm32
f1_keywords:
- vbaol11.chm32
ms.prod: outlook
api_name:
- Outlook.AddressEntries.Add
ms.assetid: b4c37547-8fbd-b1e4-40f3-5cba3cffd6e9
ms.date: 06/08/2017
---


# AddressEntries.Add Method (Outlook)

Adds a new entry to the  **[AddressEntries](addressentries-object-outlook.md)** collection.


## Syntax

 _expression_ . **Add**( **_Type_** , **_Name_** , **_Address_** )

 _expression_ An **AddressEntries** object that represents the new entry.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **String**|The type of the new entry.|
| _Name_|Optional| **Variant**|The name of the new entry.|
| _Address_|Optional| **Variant**|The address.|

### Return Value

An  **[AddressEntry](addressentry-object-outlook.md)** object that represents the new entry.


## Remarks

New entries or changes to existing entries are not persisted in the collection until after calling the  **[Update](addressentry-update-method-outlook.md)** method.


## See also


#### Concepts


[AddressEntries Object](addressentries-object-outlook.md)

