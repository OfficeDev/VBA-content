---
title: Columns.Item Method (Outlook)
keywords: vbaol11.chm2740
f1_keywords:
- vbaol11.chm2740
ms.prod: outlook
api_name:
- Outlook.Columns.Item
ms.assetid: d9abb503-32ea-d98b-bc43-d818c8b72883
ms.date: 06/08/2017
---


# Columns.Item Method (Outlook)

Obtains a  **[Column](column-object-outlook.md)** object specified by _Index_ .


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **[Columns](columns-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|A 1-based index value that can be either a  **Long** representing the column index for the **Columns** collection or a **String** representing the **[Name](column-name-property-outlook.md)** of the **Column** .|

### Return Value

 A **Column** object that represents the column matching the _Index_ in the **[Table](table-object-outlook.md)** . Returns the error, "Array index out of bounds" if _Index_ is an invalid **Long** integer. Returns **Null** ( **Nothing** in Visual Basic) if _Index_ is a **String** representing a column name that cannot be found in the **Table** .


## See also


#### Concepts


[Columns Object](columns-object-outlook.md)

