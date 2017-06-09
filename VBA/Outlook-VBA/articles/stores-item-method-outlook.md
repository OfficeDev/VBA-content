---
title: Stores.Item Method (Outlook)
keywords: vbaol11.chm819
f1_keywords:
- vbaol11.chm819
ms.prod: outlook
api_name:
- Outlook.Stores.Item
ms.assetid: b516241a-7baf-b04b-027d-25de80058fbe
ms.date: 06/08/2017
---


# Stores.Item Method (Outlook)

Returns a  **[Store](store-object-outlook.md)** object that is specified by _Index_ . Read-only.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Stores** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Either an  **Integer** that specifies a one-based index into the **Stores** collection, or a **String** value that specifies the **[DisplayName](store-displayname-property-outlook.md)** of a **Store** in the **Stores** collection.|

### Return Value

A  **Store** object in the parent **[Stores](stores-object-outlook.md)** collection, as specified by _Index_ .


## Remarks

The  **Store.DisplayName** property is the default property of a **Store** .

If  _Index_ is a string and no item can be found by that name, an error will be returned.


## See also


#### Concepts


[Stores Object](stores-object-outlook.md)

