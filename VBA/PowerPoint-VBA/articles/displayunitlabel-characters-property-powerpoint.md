---
title: DisplayUnitLabel.Characters Property (PowerPoint)
keywords: vbapp10.chm699002
f1_keywords:
- vbapp10.chm699002
ms.prod: powerpoint
api_name:
- PowerPoint.DisplayUnitLabel.Characters
ms.assetid: 58a26e5b-3526-7eb9-f9c6-a6d83708bad8
ms.date: 06/08/2017
---


# DisplayUnitLabel.Characters Property (PowerPoint)

Returns a  **[ChartCharacters](chartcharacters-object-powerpoint.md)** object that represents a range of characters within the object text. You can use the **ChartCharacters** object to format characters within a text string.


## Syntax

 _expression_. **Characters**( **_Start_**, **_Length_** )

 _expression_ A variable that represents a **[DisplayUnitLabel](displayunitlabel-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Variant**|The first character to be returned. If this argument is either 1 or omitted, this property returns a range of characters starting with the first character.|
| _Length_|Optional|**Variant**|The number of characters to be returned. If this argument is omitted, this property returns the remainder of the string (everything after the Start character).|

## Remarks

The  **ChartCharacters** object is not a collection.


## See also


#### Concepts


[DisplayUnitLabel Object](displayunitlabel-object-powerpoint.md)

