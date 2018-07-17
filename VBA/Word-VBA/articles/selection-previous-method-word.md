---
title: Selection.Previous Method (Word)
keywords: vbawd10.chm158662762
f1_keywords:
- vbawd10.chm158662762
ms.prod: word
api_name:
- Word.Selection.Previous
ms.assetid: 85679323-fe2c-f37a-5373-2c9e6d8494eb
ms.date: 06/08/2017
---


# Selection.Previous Method (Word)

Moves the selected text by the specified number of units, and returns a  **Range** object relative to the collapsed selection.


## Syntax

 _expression_ . **Previous**( **_Unit_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **Variant**|Specifies the type of unit by which to move the selection. Can be one of the  **[WdUnits](wdunits-enumeration-word.md)** constants.|
| _Count_|Optional| **Variant**|The number of units by which you want to move. The default value is 1.|

### Return Value

Range


## Remarks

If the selection is just after the specified Unit, the selection is moved to the previous unit. For example, if the selection is just after a word (before the trailing space), the following instruction moves the selection backward to the previous word.


```
Selection.Previous(Unit:=wdWord, Count:=1).Select
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

