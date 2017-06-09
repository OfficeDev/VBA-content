---
title: ListTemplate.Convert Method (Word)
keywords: vbawd10.chm160366693
f1_keywords:
- vbawd10.chm160366693
ms.prod: word
api_name:
- Word.ListTemplate.Convert
ms.assetid: 5b25c80e-a39c-3bcb-5c5f-bb9001e1ca86
ms.date: 06/08/2017
---


# ListTemplate.Convert Method (Word)

Converts a multiple-level list to a single-level list, or vice versa.


## Syntax

 _expression_ . **Convert**( **_Level_** )

 _expression_ Required. A variable that represents a **[ListTemplate](listtemplate-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Level_|Optional| **Variant**|The level to use for formatting the new list. When converting a multiple-level list to a single-level list, this argument can be a number from 1 through 9. When converting a single-level list to a multiple-level list, 1 is the only valid value. If this argument is omitted, 1 is the default value.|

## Remarks

You cannot use the  **Convert** method on a list template that is derived from the **ListGalleries** collection.


## Example

This example converts the first list template in the active document. If the list template is multiple-level, it becomes single-level, or vice versa.


```vb
ActiveDocument.ListTemplates(1).Convert
```


## See also


#### Concepts


[ListTemplate Object](listtemplate-object-word.md)

