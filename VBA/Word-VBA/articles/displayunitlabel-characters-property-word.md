---
title: DisplayUnitLabel.Characters Property (Word)
keywords: vbawd10.chm94568450
f1_keywords:
- vbawd10.chm94568450
ms.prod: word
api_name:
- Word.DisplayUnitLabel.Characters
ms.assetid: 9557eeff-306c-8f12-98c0-9a6e633b9511
ms.date: 06/08/2017
---


# DisplayUnitLabel.Characters Property (Word)

Returns a  **[ChartCharacters](chartcharacters-object-word.md)** object that represents a range of characters within the object text. You can use the **ChartCharacters** object to format characters within a text string.


## Syntax

 _expression_ . **Characters**( **_Start_** , **_Length_** )

 _expression_ A variable that represents a **[DisplayUnitLabel](displayunitlabel-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional| **Variant**|The first character to be returned. If this argument is either 1 or omitted, this property returns a range of characters starting with the first character.|
| _Length_|Optional| **Variant**|The number of characters to be returned. If this argument is omitted, this property returns the remainder of the string (everything after the Start character).|

## Remarks

The  **ChartCharacters** object is not a collection.


## See also


#### Concepts


[DisplayUnitLabel Object](displayunitlabel-object-word.md)

