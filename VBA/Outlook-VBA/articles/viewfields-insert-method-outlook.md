---
title: ViewFields.Insert Method (Outlook)
keywords: vbaol11.chm2553
f1_keywords:
- vbaol11.chm2553
ms.prod: outlook
api_name:
- Outlook.ViewFields.Insert
ms.assetid: a975a030-76c9-e877-8df7-601094998fd1
ms.date: 06/08/2017
---


# ViewFields.Insert Method (Outlook)

Creates a new  **[ViewField](viewfield-object-outlook.md)** object and inserts it at the specified index within the **[ViewFields](viewfields-object-outlook.md)** collection.


## Syntax

 _expression_ . **Insert**( **_PropertyName_** , **_Index_** )

 _expression_ A variable that represents a **ViewFields** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PropertyName_|Required| **String**|The name of the property to which the new object is associated.|
| _Index_|Required| **Variant**|Either a one-based index number at which to insert the new object, or a value used to match the  **[ViewXMLSchemaName](viewfield-viewxmlschemaname-property-outlook.md)** property value of an object in the collection where the new object is to be inserted.|

### Return Value

A  **ViewField** object that represents the new view field.


## See also


#### Concepts


[ViewFields Object](viewfields-object-outlook.md)

