---
title: Results.Sort Method (Outlook)
keywords: vbaol11.chm511
f1_keywords:
- vbaol11.chm511
ms.prod: outlook
api_name:
- Outlook.Results.Sort
ms.assetid: d897f4c9-ef58-cdb4-ca9e-d179af12f2b5
ms.date: 06/08/2017
---


# Results.Sort Method (Outlook)

Sorts the collection of items by the specified property. The index for the collection is reset to 1 upon completion of this method.


## Syntax

 _expression_ . **Sort**( **_Property_** , **_Descending_** )

 _expression_ A variable that represents a **Results** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Property_|Required| **String**|The name of the property by which to sort, which may be enclosed in brackets (for example, "[CompanyName]"). May not be a user-defined field, and may not be a multi-valued property, such as a category.|
| _Descending_|Optional| **Variant**| **True** to sort in descending order. The default value is **False** (ascending).|

## Remarks

 **Sort** only affects the order of items in a collection. It does not affect the order of items in an explorer view.


## See also


#### Concepts


[Results Object](results-object-outlook.md)

