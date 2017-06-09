---
title: Categories.Remove Method (Outlook)
keywords: vbaol11.chm2438
f1_keywords:
- vbaol11.chm2438
ms.prod: outlook
api_name:
- Outlook.Categories.Remove
ms.assetid: 8c16b02e-0297-9f36-7cb7-20e6ab0c286b
ms.date: 06/08/2017
---


# Categories.Remove Method (Outlook)

Removes an object from the collection.


## Syntax

 _expression_ . **Remove**( **_Index_** )

 _expression_ A variable that represents a **Categories** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Either a  **Long** value representing the index number of the object, or a **String** value representing either the **[Name](category-name-property-outlook.md)** or **[CategoryID](category-categoryid-property-outlook.md)** property value of an object in the collection.|

## Remarks

If the name of a category is specified in  _Index_, this method removes the first  **Category** object that matches the specified value.


## See also


#### Concepts


[Categories Object](categories-object-outlook.md)

