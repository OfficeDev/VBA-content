---
title: Range.Previous Method (Word)
keywords: vbawd10.chm157155434
f1_keywords:
- vbawd10.chm157155434
ms.prod: word
api_name:
- Word.Range.Previous
ms.assetid: ee1135ec-6f88-ec52-c3cc-0fb8183ac4cd
ms.date: 06/08/2017
---


# Range.Previous Method (Word)

Returns the previous range a relative to the specified range.


## Syntax

 _expression_ . **Previous**( **_Unit_** , **_Count_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **Variant**|The type of units by which to count. Can be any  **WdUnits** constant.|
| _Count_|Optional| **Variant**|The number of units by which you want to move back. The default value is 1.|

### Return Value

Range


## Remarks

If the  **Range** object is just after the specified Unit, the **Range** object is moved to the previous unit. For example, if the **Range** object is just after a word (before the trailing space), the following instruction moves the **Range** object backward to the previous word.


```vb
ActiveDocument.Words(2).Previous(Unit:=wdWord, Count:=1).Select
```


## Example

This example applies bold formatting to the first word in the active document.


```vb
ActiveDocument.Words(2) _ 
 .Previous(Unit:=wdWord, Count:=1).Bold = True
```


## See also


#### Concepts


[Range Object](range-object-word.md)

