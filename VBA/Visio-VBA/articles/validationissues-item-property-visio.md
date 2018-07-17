---
title: ValidationIssues.Item Property (Visio)
keywords: vis_sdr.chm18513765
f1_keywords:
- vis_sdr.chm18513765
ms.prod: visio
api_name:
- Visio.ValidationIssues.Item
ms.assetid: b8fb6413-4da7-f600-e730-f1e1b21e34fe
ms.date: 06/08/2017
---


# ValidationIssues.Item Property (Visio)

Returns the  **[ValidationIssue](validationissue-object-visio.md)** object that has the specified name or index position. The **Item** property is the default property for all collections. Read-only.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **[ValidationIssues](validationissues-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number of the object in its collection.|

### Return Value

 **ValidationIssue**


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression, because it is the default property for all collections.


