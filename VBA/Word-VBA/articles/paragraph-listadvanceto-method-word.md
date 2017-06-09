---
title: Paragraph.ListAdvanceTo Method (Word)
keywords: vbawd10.chm156696912
f1_keywords:
- vbawd10.chm156696912
ms.prod: word
api_name:
- Word.Paragraph.ListAdvanceTo
ms.assetid: 41b60f22-74b1-60f6-40ad-4107074a57ee
ms.date: 06/08/2017
---


# Paragraph.ListAdvanceTo Method (Word)

Sets the list levels for a paragraph in a list.


## Syntax

 _expression_ . **ListAdvanceTo**( **_Level1_** , **_Level2_** , **_Level3_** , **_Level4_** , **_Level5_** , **_Level6_** , **_Level7_** , **_Level8_** , **_Level9_** )

 _expression_ An expression that returns a **[Paragraph](paragraph-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Level1 to Level9_|Optional| **Integer**|Specifies the list level.|

## Remarks

Microsoft Word adjusts the numbering value to be used for numbering each list level for the specified paragraph by inserting the necessary intervening paragraphs as hidden text between the specified paragraph and the previous paragraph.


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

