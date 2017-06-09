---
title: Range.ConvertHangulAndHanja Method (Word)
keywords: vbawd10.chm157155549
f1_keywords:
- vbawd10.chm157155549
ms.prod: word
api_name:
- Word.Range.ConvertHangulAndHanja
ms.assetid: 2b640faf-da3c-a3b6-976b-d7dca3cb710f
ms.date: 06/08/2017
---


# Range.ConvertHangulAndHanja Method (Word)

Converts the specified range from hangul to hanja or vice versa.


## Syntax

 _expression_ . **ConvertHangulAndHanja**( **_ConversionsMode_** , **_FastConversion_** , **_CheckHangulEnding_** , **_EnableRecentOrdering_** , **_CustomDictionary_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ConversionsMode_|Optional| **Variant**|Sets the direction for the conversion between hangul and hanja. Can be either of the following  **WdMultipleWordConversionsMode** constants: **wdHangulToHanja** or **wdHanjaToHangul** . The default value is the current value of the **MultipleWordConversionsMode** property.|
| _FastConversion_|Optional| **Variant**| **True** if Microsoft Word automatically converts a word with only one suggestion for conversion. The default value is the current value of the **HangulHanjaFastConversion** property.|
| _CheckHangulEnding_|Optional| **Variant**| **True** if Word automatically detects hangul endings and ignores them. The default value is the current value of the **CheckHangulEndings** property. This argument is ignored if the ConversionsMode argument is set to **wdHanjaToHangul** .|
| _EnableRecentOrdering_|Optional| **Variant**| **True** if Word displays the most recently used words at the top of the suggestions list. The default value is the current value of the **EnableHangulHanjaRecentOrdering** property.|
| _CustomDictionary_|Optional| **Variant**|The name of a custom hangul-hanja conversion dictionary. Use this argument to use a custom dictionary with hangul-hanja conversions not contained in the main dictionary.|

## Example

This example converts the current selection from hangul to hanja.


```vb
Selection.Range.ConvertHangulAndHanja _ 
 ConversionsMode:=wdHangulToHanja, _ 
 FastConversion:=True, _ 
 EnableRecentOrdering:= True
```


## See also


#### Concepts


[Range Object](range-object-word.md)

