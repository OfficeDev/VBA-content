---
title: OrderFields.Item Method (Outlook)
keywords: vbaol11.chm2677
f1_keywords:
- vbaol11.chm2677
ms.prod: outlook
api_name:
- Outlook.OrderFields.Item
ms.assetid: 0738f59e-8eda-18af-1aee-13d566c248db
ms.date: 06/08/2017
---


# OrderFields.Item Method (Outlook)

Returns an  **[OrderField](orderfield-object-outlook.md)** object from the collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents an **OrderFields** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The value can be a one-based integer that indexes an  **OrderField** object in the **[OrderFields](orderfields-object-outlook.md)** collection, a string that matches the **[ViewXMLSchemaName](orderfield-viewxmlschemaname-property-outlook.md)** property value of an **OrderField** object in the collection, or a field name as displayed in the Field Chooser.|

### Return Value

An  **OrderField** object that represents the specified object.


## See also


#### Concepts


[OrderFields Object](orderfields-object-outlook.md)

