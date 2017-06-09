---
title: OrderFields.Insert Method (Outlook)
keywords: vbaol11.chm2681
f1_keywords:
- vbaol11.chm2681
ms.prod: outlook
api_name:
- Outlook.OrderFields.Insert
ms.assetid: b866034d-4999-ebab-7f18-5fd63f169564
ms.date: 06/08/2017
---


# OrderFields.Insert Method (Outlook)

Creates a new  **[OrderField](orderfield-object-outlook.md)** object and inserts it at the specified index within the **[OrderFields](orderfields-object-outlook.md)** collection.


## Syntax

 _expression_ . **Insert**( **_PropertyName_** , **_Index_** , **_IsDescending_** )

 _expression_ A variable that represents an **OrderFields** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PropertyName_|Required| **String**|The name of the property to which the new object is associated.|
| _Index_|Required| **Variant**|Either the index number at which to insert the new object, or a value used to match the  **[ViewXMLSchemaName](orderfield-viewxmlschemaname-property-outlook.md)** property value of an object in the collection at where the new object is to be inserted.|
| _IsDescending_|Optional| **Boolean**|The value used to set the  **[IsDescending](orderfield-isdescending-property-outlook.md)** property of the new **OrderField** object. If this value is not specified, the default value of the **IsDescending** property is used.|

### Return Value

An  **OrderField** object that represents the new order field.


## See also


#### Concepts


[OrderFields Object](orderfields-object-outlook.md)

