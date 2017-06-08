---
title: Fields.AddPhoneticGuide Method (Publisher)
keywords: vbapb10.chm6029320
f1_keywords:
- vbapb10.chm6029320
ms.prod: publisher
api_name:
- Publisher.Fields.AddPhoneticGuide
ms.assetid: 9b64e505-3aa7-040f-f791-f2dbeaf6860e
ms.date: 06/08/2017
---


# Fields.AddPhoneticGuide Method (Publisher)

Returns a  **[Field](field-object-publisher.md)** object that represents phonetic text added to the specified range.


## Syntax

 _expression_. **AddPhoneticGuide**( **_Range_**,  **_Text_**,  **_Alignment_**,  **_Raise_**,  **_FontName_**,  **_FontSize_**)

 _expression_A variable that represents a  **Fields** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Range|Required| **TextRange**|The text in the publication over which the phonetic text is displayed.|
|Text|Required| **String**|The phonetic text to add.|
|Alignment|Optional| **PbPhoneticGuideAlignmentType**|The alignment of the added phonetic text.|
|Raise|Optional| **Variant**|The distance (in points) from the top of the text in the specified range to the top of the phonetic text. If no value is specified, Microsoft Publisher automatically sets the phonetic text at an optimum distance above the specified range.|
|FontName|Optional| **String**|The name of the font to use for the phonetic text. If no value is specified, Publisher uses the same font as the text in the specified range.|
|FontSize|Optional| **Variant**|The font size to use for the phonetic text. Default is 10 point.|

### Return Value

Field


## Remarks

The Alignment parameter can be one of the  **PbPhoneticGuideAlignmentType** constants declared in the Microsoft Publisher type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **pbPhoneticGuideAlignmentCenter**|Centers phonetic text over the specified range.|
| **pbPhoneticGuideAlignmentDefault**|Centers phonetic text over the specified range. The default.|
| **pbPhoneticGuideAlignmentLeft**| Left-aligns phonetic text with the specified range.|
| **pbPhoneticGuideAlignmentOneTwoOne**|Adjusts the inside and outside spacing of the phonetic text in a 1:2:1 ratio.|
| **pbPhoneticGuideAlignmentRight**|Right-aligns phonetic text with the specified range.|
| **pbPhoneticGuideAlignmentZeroOneZero**|Adjusts the inside and outside spacing of the phonetic text in a 0:1:0 ratio.|

## Example

This example adds a phonetic guide to the selected phrase "very nice."


```vb
Sub PhoneticGuide() 
 Selection.TextRange.Fields.AddPhoneticGuide _ 
 Range:=Selection.TextRange, Text:="ver-E nIs", _ 
 Alignment:=pbPhoneticGuideAlignmentCenter, _ 
 Raise:=11, FontSize:=7 
End Sub
```


