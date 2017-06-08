---
title: AxisTitle.Characters Property (PowerPoint)
keywords: vbapp10.chm683002
f1_keywords:
- vbapp10.chm683002
ms.prod: powerpoint
api_name:
- PowerPoint.AxisTitle.Characters
ms.assetid: 8b1b9dc9-6aa3-872f-964a-fe623feff6fa
ms.date: 06/08/2017
---


# AxisTitle.Characters Property (PowerPoint)

Returns a  **[ChartCharacters](chartcharacters-object-powerpoint.md)** object that represents a range of characters within the object text. You can use the **ChartCharacters** object to format characters within a text string.


## Syntax

 _expression_. **Characters**( **_Start_**, **_Length_** )

 _expression_ A variable that represents an **[AxisTitle](axistitle-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Variant**|The first character to be returned. If this argument is either 1 or omitted, this property returns a range of characters starting with the first character.|
| _Length_|Optional|**Variant**|The number of characters to be returned. If this argument is omitted, this property returns the remainder of the string (everything after the Start character).|

## Remarks

The  **ChartCharacters** object is not a collection.


## See also


#### Concepts


[AxisTitle Object](axistitle-object-powerpoint.md)

