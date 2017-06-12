---
title: Range.PhoneticGuide Method (Word)
keywords: vbawd10.chm157155552
f1_keywords:
- vbawd10.chm157155552
ms.prod: word
api_name:
- Word.Range.PhoneticGuide
ms.assetid: f720cf42-4d61-977c-8e09-6346a48afecf
ms.date: 06/08/2017
---


# Range.PhoneticGuide Method (Word)

Adds phonetic guides to the specified range.


## Syntax

 _expression_ . **PhoneticGuide**( **_Text_** , **_Alignment_** , **_Raise_** , **_FontSize_** , **_FontName_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Required| **String**|The phonetic text to add.|
| _Alignment_|Optional| **WdPhoneticGuideAlignmentType**|The alignment of the added phonetic text.|
| _Raise_|Optional| **Long**|The distance (in points) from the top of the text in the specified range to the top of the phonetic text. If no value is specified, Microsoft Word automatically sets the phonetic text at an optimum distance above the specified range.|
| _FontSize_|Optional| **Long**|The font size to use for the phonetic text. If no value is specified, Word uses a font size 50 percent smaller than the text in the specified range.|
| _FontName_|Optional| **String**|The name of the font to use for the phonetic text. If no value is specified, Word uses the same font as the text in the specified range.|

## Remarks

For more information on using Word with East Asian languages, see Word features for East Asian languages.


## Example

This example adds a phonetic guide to the selected phrase "tres chic."


```
Selection.Range.PhoneticGuide Text:="tray sheek", _ 
 Alignment:=wdPhoneticGuideAlignmentCenter, _ 
 Raise:=11, FontSize:=7
```


## See also


#### Concepts


[Range Object](range-object-word.md)

