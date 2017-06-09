---
title: Categories.Item Method (Outlook)
keywords: vbaol11.chm2436
f1_keywords:
- vbaol11.chm2436
ms.prod: outlook
api_name:
- Outlook.Categories.Item
ms.assetid: 7bdf22ec-7c77-1f1f-e4fd-77bdcc0ea288
ms.date: 06/08/2017
---


# Categories.Item Method (Outlook)

Returns a  **[Category](category-object-outlook.md)** object from the collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Categories** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Either a  **Long** value representing the index number of the object, or a **String** value representing either the **[Name](category-name-property-outlook.md)** or **[CategoryID](category-categoryid-property-outlook.md)** property value of an object in the collection.|

### Return Value

A  **Category** object that represents the specified object.


## Remarks

If the name of a category is specified in  _Index_, this method returns the first  **Category** object that matches the specified value. If a match cannot be found, the method returns **Null** ( **Nothing** in Visual Basic.)


## See also


#### Concepts


[Categories Object](categories-object-outlook.md)

