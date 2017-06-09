---
title: ValidationRules.Item Property (Visio)
keywords: vis_sdr.chm18313765
f1_keywords:
- vis_sdr.chm18313765
ms.prod: visio
api_name:
- Visio.ValidationRules.Item
ms.assetid: 4133f9ba-ca20-104a-5a30-7de37b978706
ms.date: 06/08/2017
---


# ValidationRules.Item Property (Visio)

Returns the  **[ValidationRule](validationrule-object-visio.md)** object that has the specified index position. The **Item** property is the default property for all collections. Read-only.


## Syntax

 _expression_ . **Item**( **_NameUOrIndex_** )

 _expression_ A variable that represents a **[ValidationRules](validationrules-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NameUOrIndex_|Required| **Variant**|The index number of the object in its collection.|

### Return Value

 **ValidationRule**


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```
objRet = object(index)
```


