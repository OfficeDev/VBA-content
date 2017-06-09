---
title: Range.AddComment Method (Excel)
keywords: vbaxl10.chm144222
f1_keywords:
- vbaxl10.chm144222
ms.prod: excel
api_name:
- Excel.Range.AddComment
ms.assetid: 89bbacad-4655-bcc1-8010-2ab367cc7b31
ms.date: 06/08/2017
---


# Range.AddComment Method (Excel)

Adds a comment to the range.


## Syntax

 _expression_ . **AddComment**( **_Text_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Optional| **Variant**|The comment text.|

### Return Value

Comment


## Example

This example adds a comment to cell E5 on worksheet one.


```vb
Worksheets(1).Range("E5").AddComment "Current Sales"
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

